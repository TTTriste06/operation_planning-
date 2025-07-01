import re
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill, Font
from datetime import datetime
from collections import defaultdict
from openpyxl.styles import numbers
from sheet_add import clean_df

def init_monthly_fields(main_plan_df: pd.DataFrame, start_date: datetime = None) -> list[int]:
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

    today = pd.Timestamp(start_date.replace(day=1)) if start_date else pd.Timestamp(datetime.today().replace(day=1))
    start_month = today.month
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
        col_sales_this = f"{month}月销售数量"
        
        # 安全提取列，如果缺失则填 0
        def get(col):
            return pd.to_numeric(main_plan_df[col], errors="coerce").fillna(0) if col in main_plan_df.columns else pd.Series(0, index=main_plan_df.index)
        
        def get_plan(col):
            return pd.to_numeric(df_plan[col], errors="coerce").fillna(0) if col in df_plan.columns else pd.Series(0, index=main_plan_df.index)

        if idx == 0:
            cond = (get(col_order_this) + get(col_sales_this) > get(col_forecast_this))
            for row_idx in main_plan_df.index:
                name = main_plan_df.at[row_idx, "品名"] if "品名" in main_plan_df.columns else f"Row {row_idx}"
                v_invpart = get("InvPart").at[row_idx]
                v_fg_inv = get("成品仓").at[row_idx]
                v_fg_wip = get("成品在制").at[row_idx]
                v_order_this = get(col_order_this).at[row_idx]
                v_sales_this = get(col_sales_this).at[row_idx]
                v_forecast_this = get(col_forecast_this).at[row_idx]
                v_forecast_next = get(col_forecast_next).at[row_idx]
                v_order_next = get(col_order_next).at[row_idx]
                max_next = max(v_forecast_next, v_order_next)

                if cond.at[row_idx]:
                    formula = f" {v_invpart} + {v_order_this} + max({v_forecast_next}, {v_order_next}) - {v_fg_inv} - {v_fg_wip}"
                    result = v_invpart + v_order_this + max_next - v_fg_inv - v_fg_wip
                else:
                    formula = f" {v_invpart} + {v_forecast_this} - {v_sales_this} + max({v_forecast_next}, {v_order_next}) - {v_fg_inv} - {v_fg_wip}"
                    result = v_invpart + v_forecast_this - v_sales_this + max_next - v_fg_inv - v_fg_wip
                df_plan.at[row_idx, col_target] = result
        else:
            for row_idx in main_plan_df.index:
                name = main_plan_df.at[row_idx, "品名"] if "品名" in main_plan_df.columns else f"Row {row_idx}"
                v_prev_plan = get_plan(col_target_prev).at[row_idx]
                v_actual = get(col_actual_prod).at[row_idx]
                v_forecast_next = get(col_forecast_next).at[row_idx]
                v_order_next = get(col_order_next).at[row_idx]

                formula = f"{v_prev_plan} - {v_actual} + {v_forecast_next}"
                result = v_prev_plan + max(v_forecast_next, v_order_next)
                df_plan.at[row_idx, col_target] = result
                
    plan_cols_in_summary = [col for col in main_plan_df.columns if "成品投单计划" in col and "半成品" not in col]
    
    # 回填到主计划中
    if len(plan_cols_in_summary) != df_plan.shape[1]:
        st.error(f"❌ 写入失败：df_plan 有 {df_plan.shape[1]} 列，summary 中有 {len(plan_cols_in_summary)} 个 '成品投单计划' 列")
    else:
        # ✅ 将 df_plan 的列按顺序填入 summary_preview
        for i, col in enumerate(plan_cols_in_summary):
            main_plan_df[col] = df_plan.iloc[:, i]

    return main_plan_df

def generate_monthly_semi_plan(main_plan_df: pd.DataFrame, forecast_months: list[int],
                                mapping_df: pd.DataFrame) -> pd.DataFrame:
    """
    半成品投单计划
    """
    tmp = mapping_df[["新品名", "半成品"]].copy()
    tmp = clean_df(tmp)                               
    tmp = tmp[tmp["半成品"].notna() & (tmp["半成品"].astype(str).str.strip() != "")]
    combined_names = pd.Series(
        tmp["新品名"].astype(str).str.strip().tolist() +
        tmp["半成品"].astype(str).str.strip().tolist()
    ).dropna().unique().tolist()             

    mask = main_plan_df["品名"].astype(str).str.strip().isin(combined_names)

    df_plan = pd.DataFrame(index=main_plan_df.index)

    for idx, month in enumerate(forecast_months[:-1]):  # 最后一个月不生成
        this_month = f"{month}月"
        next_month = f"{forecast_months[idx + 1]}月"
        prev_month = f"{forecast_months[idx - 1]}月" if idx > 0 else None

        col_target = f"{this_month}半成品投单计划"
        col_plan_this = f"{this_month}成品投单计划"
        col_forecast_next = f"{next_month}预测"
        col_order_next = f"未交订单 2025-{forecast_months[idx + 1]:02d}"
        col_actual_prod = f"{prev_month}半成品实际投单"
        col_target_prev = f"{prev_month}半成品投单计划" if prev_month else None
        

        def get(col):
            return pd.to_numeric(main_plan_df[col], errors="coerce").fillna(0) if col in main_plan_df.columns else pd.Series(0, index=main_plan_df.index)

        def get_plan(col):
            return pd.to_numeric(df_plan[col], errors="coerce").fillna(0) if col in df_plan.columns else pd.Series(0, index=main_plan_df.index)

        if idx == 0:
            plan_this = get(col_plan_this)
            sfg = get("半成品仓")
            sfg_wip = get("半成品在制")
            
            result = plan_this - sfg - sfg_wip
            df_plan[col_target] = result
        else:
            for row_idx in main_plan_df.index:
                prev_plan = get_plan(col_target_prev).at[row_idx]
                actual_prod = get(col_actual_prod).at[row_idx]
                forecast_next = get(col_forecast_next).at[row_idx]
                order_next = get(col_order_next).at[row_idx]
                result = prev_plan + max(forecast_next, order_next)
                df_plan.at[row_idx, col_target] = result

    plan_cols_in_summary = [col for col in main_plan_df.columns if "半成品投单计划" in col]

    if len(plan_cols_in_summary) != df_plan.shape[1]:
        st.error(f"❌ 写入失败：df_plan 有 {df_plan.shape[1]} 列，summary 中有 {len(plan_cols_in_summary)} 个 '半成品投单计划' 列")
    else:
        for i, col in enumerate(plan_cols_in_summary):
            main_plan_df[col] = ""
            col_in_df_plan = df_plan.columns[i] if i < len(df_plan.columns) else None
            if col_in_df_plan:
                main_plan_df.loc[mask, col] = df_plan.loc[mask, col_in_df_plan]
                
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
                
    # 回填结果到主计划表
    for col in order_summary.columns[1:]:
        main_plan_df[col] = order_summary[col]

    return main_plan_df

def aggregate_actual_sfg_orders(main_plan_df: pd.DataFrame, df_order: pd.DataFrame, mapping_df: pd.DataFrame, forecast_months: list[int]) -> pd.DataFrame:
    """
    提取“半成品实际投单”数据并写入主计划表，依据“赛卓-新旧料号”中“半成品”字段进行反查。

    参数：
    - main_plan_df: 主计划 DataFrame，需包含“品名”列
    - df_order: 下单明细，含“下单日期”、“回货明细_回货品名”、“回货明细_回货数量”
    - mapping_df: 新旧料号表，含“半成品”字段和“新品名”
    - forecast_months: 月份整数列表

    返回：
    - main_plan_df: 写入了“X月半成品实际投单”的 DataFrame
    """
    if df_order.empty or mapping_df.empty or not forecast_months:
        return main_plan_df

    df_order = df_order.copy()
    df_order = df_order[["下单日期", "回货明细_回货品名", "回货明细_回货数量"]].dropna()
    df_order["回货明细_回货品名"] = df_order["回货明细_回货品名"].astype(str).str.strip()
    df_order["下单月份"] = pd.to_datetime(df_order["下单日期"], errors="coerce").dt.month

    # 生成半成品 → 新品名 映射字典
    semi_mapping = mapping_df[mapping_df["半成品"].notna() & (mapping_df["半成品"] != "")]
    semi_dict = dict(zip(semi_mapping["半成品"].astype(str).str.strip(), semi_mapping["新品名"].astype(str).str.strip()))

    # 初始化结果 DataFrame
    sfg_summary = pd.DataFrame({"品名": main_plan_df["品名"].astype(str)})
    for m in forecast_months:
        sfg_summary[f"{m}月半成品实际投单"] = 0

    # 逐行分配
    for _, row in df_order.iterrows():
        part = row["回货明细_回货品名"]
        qty = row["回货明细_回货数量"]
        month = row["下单月份"]
        col_name = f"{month}月半成品实际投单"

        if part in semi_dict and month in forecast_months:
            new_part = semi_dict[part]
            match_idx = sfg_summary[sfg_summary["品名"] == new_part].index
            if not match_idx.empty:
                sfg_summary.loc[match_idx[0], col_name] += qty

    # 写入主计划
    for col in sfg_summary.columns[1:]:
        main_plan_df[col] = sfg_summary[col]

    return main_plan_df

def aggregate_actual_arrivals(main_plan_df: pd.DataFrame, df_arrival: pd.DataFrame, forecast_months: list[int]) -> pd.DataFrame:
    """
    从“到货明细”中提取回货实际数量并填入主计划表。

    参数：
    - main_plan_df: 主计划 DataFrame（需包含“品名”列）
    - df_arrival: 到货明细 DataFrame，含“到货日期”、“品名”、“允收数量”
    - forecast_months: 月份整数列表，如 [6, 7, 8]

    返回：
    - main_plan_df: 添加了“X月回货实际”的列
    """
    if df_arrival.empty or not forecast_months:
        return main_plan_df

    # 保留有效列并清洗
    df_arrival = df_arrival[["到货日期", "品名", "允收数量"]].dropna()
    df_arrival["品名"] = df_arrival["品名"].astype(str).str.strip()
    df_arrival["到货月份"] = pd.to_datetime(df_arrival["到货日期"], errors="coerce").dt.month

    # 初始化结果表
    result_df = pd.DataFrame({"品名": main_plan_df["品名"].astype(str)})
    for m in forecast_months:
        result_df[f"{m}月回货实际"] = 0

    # 汇总每月数据
    for _, row in df_arrival.iterrows():
        part = row["品名"]
        qty = row["允收数量"]
        month = row["到货月份"]
        col = f"{month}月回货实际"
        if month in forecast_months:
            match_idx = result_df[result_df["品名"] == part].index
            if not match_idx.empty:
                result_df.loc[match_idx[0], col] += qty

    # 写入主计划表
    for col in result_df.columns[1:]:
        main_plan_df[col] = result_df[col]

    return main_plan_df

def aggregate_sales_quantity_and_amount(main_plan_df: pd.DataFrame, df_sales: pd.DataFrame, forecast_months: list[int]) -> pd.DataFrame:
    """
    将销货明细中的销售数量和销售金额按照月份填入主计划表。

    参数：
    - main_plan_df: 主计划 DataFrame（含“品名”列）
    - df_sales: 销货明细 DataFrame，含“交易日期”、“品名”、“数量”、“原币金额”
    - forecast_months: 月份列表，如 [6, 7, 8]

    返回：
    - main_plan_df: 添加了“X月销售数量”和“X月销售金额”的列
    """
    if df_sales.empty or not forecast_months:
        return main_plan_df

    df_sales = df_sales[["交易日期", "品名", "数量", "原币金额"]].dropna()
    df_sales["品名"] = df_sales["品名"].astype(str).str.strip()
    df_sales["销售月份"] = pd.to_datetime(df_sales["交易日期"], errors="coerce").dt.month

    result_qty = pd.DataFrame({"品名": main_plan_df["品名"].astype(str)})
    result_amt = pd.DataFrame({"品名": main_plan_df["品名"].astype(str)})
    for m in forecast_months:
        result_qty[f"{m}月销售数量"] = 0
        result_amt[f"{m}月销售金额"] = 0

    for _, row in df_sales.iterrows():
        part = row["品名"]
        qty = row["数量"]
        amt = row["原币金额"]
        month = row["销售月份"]
        if month in forecast_months:
            col_qty = f"{month}月销售数量"
            col_amt = f"{month}月销售金额"
            match_idx = result_qty[result_qty["品名"] == part].index
            if not match_idx.empty:
                result_qty.loc[match_idx[0], col_qty] += qty
                result_amt.loc[match_idx[0], col_amt] += amt

    for col in result_qty.columns[1:]:
        main_plan_df[col] = result_qty[col]

    for col in result_amt.columns[1:]:
        main_plan_df[col] = result_amt[col]

    return main_plan_df

def generate_monthly_adjust_plan(main_plan_df: pd.DataFrame) -> pd.DataFrame:
    """
    根据已有字段直接填充投单计划调整列。
    第一个月为空，后续为公式字符串。
    """
    adjust_cols = [col for col in main_plan_df.columns if "投单计划调整" in col]
    fg_plan_cols = [col for col in main_plan_df.columns if "成品投单计划" in col and "半成品" not in col]
    fg_actual_cols = [col for col in main_plan_df.columns if "成品实际投单" in col and "半成品" not in col]

    if not adjust_cols or not fg_plan_cols or not fg_actual_cols:
        raise ValueError("❌ 缺少必要的列：投单计划调整 / 成品投单计划 / 成品实际投单")

    for i, col in enumerate(adjust_cols):
        if i == 0:
            # 第一个月为空字符串
            main_plan_df[col] = ""
        else:
            # 后续月：写入公式
            curr_plan_col = fg_plan_cols[i] if i < len(fg_plan_cols) else None
            prev_plan_col = fg_plan_cols[i - 1]
            prev_actual_col = fg_actual_cols[i - 1]

            # 获取 Excel 的列号（+1 因为 openpyxl 是从 1 开始）
            col_curr_plan = get_column_letter(main_plan_df.columns.get_loc(curr_plan_col) + 1)
            col_prev_plan = get_column_letter(main_plan_df.columns.get_loc(prev_plan_col) + 1)
            col_prev_actual = get_column_letter(main_plan_df.columns.get_loc(prev_actual_col) + 1)

            def build_formula(row_idx: int) -> str:
                row_num = row_idx + 3  # 数据起始于 Excel 第 3 行
                return f"={col_curr_plan}{row_num}+({col_prev_plan}{row_num}-{col_prev_actual}{row_num})"

            main_plan_df[col] = [build_formula(i) for i in range(len(main_plan_df))]
    return main_plan_df

def generate_monthly_return_adjustment(main_plan_df: pd.DataFrame) -> pd.DataFrame:
    """
    填写“回货计划调整”列：
    - 第一个月为空
    - 后续月份：= 本月回货计划 + (上月成品实际投单 - 上月投单计划调整)
    """
    adjust_return_cols = [col for col in main_plan_df.columns if "回货计划调整" in col]
    return_plan_cols = [col for col in main_plan_df.columns if "回货计划" in col and "调整" not in col and "PC" not in col]
    actual_plan_cols = [col for col in main_plan_df.columns if "成品实际投单" in col and "半成品" not in col]
    adjust_plan_cols = [col for col in main_plan_df.columns if "投单计划调整" in col]

    for i in range(len(adjust_return_cols)):
        col_adjust = adjust_return_cols[i]
        col_return = return_plan_cols[i]

        # 本月列索引
        col_idx_return = main_plan_df.columns.get_loc(col_return) + 1
        col_idx_adjust = main_plan_df.columns.get_loc(col_adjust) + 1

        # 上月列（用于差值计算）
        if i > 0:
            col_idx_prev_actual = main_plan_df.columns.get_loc(actual_plan_cols[i - 1]) + 1
            col_idx_prev_adjust = main_plan_df.columns.get_loc(adjust_plan_cols[i - 1]) + 1

        for row in range(3, len(main_plan_df) + 3):  # 第3行起是数据行
            if i == 0:
                main_plan_df.at[row - 3, col_adjust] = ""
            else:
                # 本月回货计划
                col_r = get_column_letter(col_idx_return)
                # 上月：成品实际投单与投单计划调整
                col_prev_actual = get_column_letter(col_idx_prev_actual)
                col_prev_adjust = get_column_letter(col_idx_prev_adjust)

                formula = f"={col_r}{row} + ({col_prev_actual}{row} - {col_prev_adjust}{row})"
                main_plan_df.at[row - 3, col_adjust] = formula

    return main_plan_df

def generate_monthly_return_plan(main_plan_df: pd.DataFrame) -> pd.DataFrame:
    """
    回货计划填写逻辑：
    - 第一个月为空；
    - 从第二个月开始，等于同一行第当前列前18列的值（通过公式表示）。
    """
    # 找出所有“回货计划”列（不含“调整”）
    return_plan_cols = [col for col in main_plan_df.columns if "回货计划" in col and "调整" not in col and "PC" not in col]
    
    # 处理每一个回货计划列
    for i, col in enumerate(return_plan_cols):
        if i == 0:
            # 第一个月为空
            main_plan_df[col] = ""
        else:
            # 获取该列在 DataFrame 中的位置
            col_idx = main_plan_df.columns.get_loc(col)
            prev_18_idx = col_idx - 18
            if prev_18_idx < 0:
                raise ValueError(f"❌ 第{i+1}月回货计划前18列不存在，列索引越界。")

            # 获取引用的前18列名
            ref_col = main_plan_df.columns[prev_18_idx]

            # 构造 Excel 公式：=INDIRECT(ADDRESS(ROW(), col_index))
            col_letter = get_column_letter(prev_18_idx + 1)  # Excel 列号从 1 开始
            main_plan_df[col] = f"={col_letter}" + (main_plan_df.index + 3).astype(str)

    return main_plan_df
    
def format_monthly_grouped_headers(ws):
    """
    从AC列开始，每13列为一个月块：
    - 合并第1行写“x月”
    - 去掉第2行每列前缀的“x月”
    - 每月块用不同背景色填充前两行
    """
    # 自动查找第2行中第一个“销售数量”字段的列号
    start_col = None
    for col_idx in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=2, column=col_idx).value
        if isinstance(cell_value, str) and "销售数量" in cell_value:
            start_col = col_idx
            break
    
    if start_col is None:
        raise ValueError("❌ 未在第2行找到“销售数量”字段，无法定位起始列")

    row_1 = 1
    row_2 = 2
    max_col = ws.max_column

    # 几组可循环的浅色背景（Excel兼容性好的十六进制RGB）
    fill_colors = [
        "FFF2CC",  # 浅黄色
        "D9EAD3",  # 浅绿色
        "D0E0E3",  # 浅蓝色
        "F4CCCC",  # 浅红色
        "EAD1DC",  # 浅紫色
        "CFE2F3",  # 浅青色
        "FFE599",  # 明亮黄
    ]

    month_pattern = re.compile(r"^(\d{1,2})月(.+)")
    col = start_col
    month_index = 0

    while col <= max_col:
        month_title = None
        fill_color = PatternFill(start_color=fill_colors[month_index % len(fill_colors)],
                                 end_color=fill_colors[month_index % len(fill_colors)],
                                 fill_type="solid")

        for offset in range(13):
            curr_col = col + offset
            cell = ws.cell(row=row_2, column=curr_col)
            value = cell.value
            if isinstance(value, str):
                match = month_pattern.match(value.strip())
                if match:
                    if month_title is None:
                        month_title = match.group(1)
                    cell.value = match.group(2)

            # 填充第2行颜色
            ws.cell(row=row_2, column=curr_col).fill = fill_color

        if month_title:
            # 合并第1行
            start_letter = get_column_letter(col)
            end_letter = get_column_letter(col + 12)
            ws.merge_cells(f"{start_letter}{row_1}:{end_letter}{row_1}")
            top_cell = ws.cell(row=row_1, column=col)
            top_cell.value = f"{month_title}月"
            top_cell.alignment = Alignment(horizontal="center", vertical="center")
            top_cell.font = Font(bold=True)
            top_cell.fill = fill_color

        col += 13
        month_index += 1


def highlight_production_plan_cells(ws, df):
    """
    根据规则给所有“成品投单计划”列标色：
    - < 0：红色
    - < 安全库存：黄色
    - > 2 * 安全库存：橙色
    """
    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    orange_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")

    # 获取列位置
    plan_cols = [col for col in df.columns if "成品投单计划" in col and "半成品" not in col]
    safety_col = "InvPart"
    if safety_col not in df.columns:
        raise ValueError("❌ 缺少“安全库存”列，无法对成品投单计划进行标色。")

    for col in plan_cols:
        col_idx = df.columns.get_loc(col) + 1  # openpyxl是1-based
        for i, val in enumerate(df[col]):
            row_idx = i + 3  # 因为第1行是合并标题，第2行是字段名
            safety = df.at[i, safety_col]

            # 进行数值判断（确保为float）
            try:
                val = float(val)
                safety = float(safety)
            except:
                continue

            cell = ws.cell(row=row_idx, column=col_idx)

            if val < 0:
                cell.fill = red_fill
            elif val < safety:
                cell.fill = yellow_fill
            elif val > 2 * safety:
                cell.fill = orange_fill

def drop_last_forecast_month_columns(main_plan_df: pd.DataFrame, forecast_months: list[int]) -> pd.DataFrame:
    """
    删除 AC列（第29列）后所有含有最后一个预测月份的字段列，如 '12月销售数量' 等。
    """
    if not forecast_months:
        return main_plan_df  # 无预测月份，不处理

    last_valid_month = forecast_months[-1]
    last_month_str = f"{last_valid_month}月"

    # 起始列为 AC = 第29列，0-based index 为 28
    fixed_part = main_plan_df.iloc[:, :28]
    dynamic_part = main_plan_df.iloc[:, 28:]

    # 仅保留不包含最后预测月的列
    dynamic_part = dynamic_part.loc[:, ~dynamic_part.columns.str.contains(fr"^{last_month_str}")]

    # 合并回主表
    cleaned_df = pd.concat([fixed_part, dynamic_part], axis=1)

    return cleaned_df

