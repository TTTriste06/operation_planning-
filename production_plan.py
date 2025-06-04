import re
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill, Font
from datetime import datetime
from collections import defaultdict
from openpyxl.styles import numbers

def init_monthly_fields(main_plan_df: pd.DataFrame) -> list[int]:
    """
    自动识别主计划中预测字段的月份，添加 HEADER_TEMPLATE 中的所有月度字段列。
    初始化为 ""。
    
    返回：
    - forecast_months: 所有识别出的月份列表（升序）
    """
    HEADER_TEMPLATE = [
        "销售数量", "销售金额", "成品投单计划", "半成品投单计划", "投单计划调整",
        "成品可行投单", "半成品可行投单", "成品实际投单", "半成品实际投单",
        "回货计划", "回货计划调整", "PC回货计划", "回货实际"
    ]

    month_pattern = re.compile(r"^(\d{1,2})月预测$")
    forecast_months = sorted({
        int(match.group(1)) for col in main_plan_df.columns
        if isinstance(col, str) and (match := month_pattern.match(col.strip()))
    })

    if not forecast_months:
        return []

    start_month = datetime.today().month
    end_month = max(forecast_months) - 1

    for m in range(start_month, end_month + 1):
        for header in HEADER_TEMPLATE:
            col = f"{m}月{header}"
            if col not in main_plan_df.columns:
                main_plan_df[col] = ""

    return forecast_months

def safe_col(df: pd.DataFrame, col: str) -> pd.Series:
    """确保列为数字，若不存在则返回 0"""
    return pd.to_numeric(df[col], errors="coerce").fillna(0) if col in df.columns else pd.Series(0, index=df.index)

def generate_monthly_fg_plan(main_plan_df: pd.DataFrame, forecast_months: list[int]) -> pd.DataFrame:
    """
    生成每月“成品投单计划”列，规则：
    - 第一个月：InvPart + max(预测, 未交) + max(预测, 未交)（下月） - 成品仓 - 成品在制
    - 后续月份：max(预测, 未交)（下月） + （上月投单 - 上月实际投单）
    
    参数：
    - main_plan_df: 主计划表（含所有字段）
    - forecast_months: 所有月份的列表（int 类型，如 [6, 7, 8, ...]）

    返回：
    - main_plan_df: 添加了成品投单计划字段的 DataFrame
    """

    df_plan = pd.DataFrame(index=main_plan_df.index)

    for idx, month in enumerate(forecast_months[:-1]):  # 最后一个月不生成
        this_month = f"{month}月"
        next_month = f"{forecast_months[idx + 1]}月"
        prev_month = f"{forecast_months[idx - 1]}月" if idx > 0 else None

        # 构造字段名
        col_forecast_this = f"{month}月预测"
        col_order_this = f"未交订单 2025-{month:02d}"
        col_forecast_next = f"{forecast_months[idx + 1]}月预测"
        col_order_next = f"未交订单 2025-{forecast_months[idx + 1]:02d}"
        col_target = f"{month}月成品投单计划"
        col_actual_prod = f"{prev_month}成品实际投单"
        col_target_prev = f"{prev_month}成品投单计划" if prev_month else None

        # 安全提取列，如果缺失则填 0
        def get(col):
            return pd.to_numeric(main_plan_df[col], errors="coerce").fillna(0) if col in main_plan_df.columns else pd.Series(0, index=main_plan_df.index)
        
        def get_plan(col):
            return pd.to_numeric(df_plan[col], errors="coerce").fillna(0) if col in df_plan.columns else pd.Series(0, index=main_plan_df.index)

        if idx == 0:
            df_plan[col_target] = (
                get("InvPart") +
                pd.concat([get(col_forecast_this), get(col_order_this)], axis=1).max(axis=1) +
                pd.concat([get(col_forecast_next), get(col_order_next)], axis=1).max(axis=1) -
                get("成品仓") -
                get("成品在制")
            )
        else:
            df_plan[col_target] = (
                pd.concat([get(col_forecast_next), get(col_order_next)], axis=1).max(axis=1) +
                (get_plan(col_target_prev) - get(col_actual_prod))
            )

    plan_cols_in_summary = [col for col in main_plan_df.columns if "成品投单计划" in col and "半成品" not in col]
    
    # 回填到主计划中
    if len(plan_cols_in_summary) != df_plan.shape[1]:
        st.error(f"❌ 写入失败：df_plan 有 {df_plan.shape[1]} 列，summary 中有 {len(plan_cols_in_summary)} 个 '成品投单计划' 列")
    else:
        # ✅ 将 df_plan 的列按顺序填入 summary_preview
        for i, col in enumerate(plan_cols_in_summary):
            main_plan_df[col] = df_plan.iloc[:, i]

    return main_plan_df

def aggregate_actual_fg_orders(main_plan_df: pd.DataFrame, df_order: pd.DataFrame, forecast_months: list[int]) -> pd.DataFrame:
    """
    从下单明细中抓取“成品实际投单”并写入 main_plan_df，每月写入“X月成品实际投单”列。
    
    参数：
    - main_plan_df: 主计划表，需包含“品名”列
    - df_order: 下单明细，含“下单日期”、“回货明细_回货品名”、“回货明细_回货数量”
    - forecast_months: 月份列表，例如 [6, 7, 8]

    返回：
    - main_plan_df: 添加了成品实际投单列的 DataFrame
    """
    if df_order.empty or not forecast_months:
        return main_plan_df

    df_order = df_order.copy()
    df_order = df_order[["下单日期", "回货明细_回货品名", "回货明细_回货数量"]].dropna()
    df_order["回货明细_回货品名"] = df_order["回货明细_回货品名"].astype(str).str.strip()
    df_order["下单月份"] = pd.to_datetime(df_order["下单日期"], errors="coerce").dt.month

    # 筛选出主计划中存在的品名
    valid_parts = set(main_plan_df["品名"].astype(str))
    df_order = df_order[df_order["回货明细_回货品名"].isin(valid_parts)]

    # 初始化结果表
    order_summary = pd.DataFrame({"品名": main_plan_df["品名"].astype(str)})
    for m in forecast_months:
        col = f"{m}月成品实际投单"
        order_summary[col] = 0

    # 累加每一行订单数量至对应月份列
    for _, row in df_order.iterrows():
        part = row["回货明细_回货品名"]
        qty = row["回货明细_回货数量"]
        month = row["下单月份"]
        col_name = f"{month}月成品实际投单"
        if month in forecast_months:
            match_idx = order_summary[order_summary["品名"] == part].index
            if not match_idx.empty:
                order_summary.loc[match_idx[0], col_name] += qty
    st.write("order_summary")
    st.write(order_summary)
    # 回填结果到主计划表
    for col in order_summary.columns[1:]:
        main_plan_df[col] = order_summary[col]

    return main_plan_df
































def apply_monthly_grouped_headers(ws):
    """
    自动合并主计划中按“月_字段”格式的列，如“6月销售数量”，将每月字段统一合并为“6月”标题。
    """
    header_row = [cell.value for cell in ws[2]]  # 第2行是字段名
    pattern = re.compile(r"^(\d{1,2})月(.*)$")

    # group: {month -> [col_idx, ...]}
    monthly_groups = defaultdict(list)

    for i, col in enumerate(header_row):
        if not isinstance(col, str):
            continue
        match = pattern.match(col.strip())
        if match:
            month = int(match.group(1))
            monthly_groups[month].append(i + 1)  # openpyxl 列号从1开始

    # 遍历每个识别到的月份
    for month, cols in sorted(monthly_groups.items()):
        start_col = min(cols)
        end_col = max(cols)

        if end_col >= start_col:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)

        cell = ws.cell(row=1, column=start_col)
        cell.value = f"{month}月"
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FFFF00")
