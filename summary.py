import re
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment

def merge_safety_inventory(summary_df: pd.DataFrame, safety_df: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """
    将安全库存表中 InvWaf 和 InvPart 信息按 '品名' 合并到汇总表中，仅根据 '品名' 匹配。
    对相同品名做 sum 汇总；未匹配的填 0。
    """
    safety_df = safety_df.rename(columns={"ProductionNO.": "品名"}).copy()
    safety_df.columns = safety_df.columns.str.strip()
    safety_df["品名"] = safety_df["品名"].astype(str).str.strip()
    safety_df["InvWaf"] = pd.to_numeric(safety_df["InvWaf"], errors="coerce").fillna(0)
    safety_df["InvPart"] = pd.to_numeric(safety_df["InvPart"], errors="coerce").fillna(0)

    safety_grouped = safety_df.groupby("品名", as_index=False)[["InvWaf", "InvPart"]].sum()

    summary_df["品名"] = summary_df["品名"].astype(str).str.strip()
    merged = summary_df.merge(safety_grouped, on="品名", how="left")

    matched_keys = set(safety_grouped["品名"])
    used_keys = set(merged[~merged[["InvWaf", "InvPart"]].isna().all(axis=1)]["品名"])
    unmatched_keys = list(matched_keys - used_keys)

    merged["InvWaf"] = merged["InvWaf"].fillna(0)
    merged["InvPart"] = merged["InvPart"].fillna(0)

    return merged, unmatched_keys


def merge_safety_header(ws: Worksheet, df: pd.DataFrame):
    """
    将“InvWaf”和“InvPart”两列的上方合并写入“安全库存”标题。
    """
    try:
        invwaf_col_idx = df.columns.get_loc("InvWaf") + 1  # openpyxl是1-indexed
        invpart_col_idx = df.columns.get_loc("InvPart") + 1

        start_col = get_column_letter(invwaf_col_idx)
        end_col = get_column_letter(invpart_col_idx)

        # 合并单元格
        ws.merge_cells(f"{start_col}1:{end_col}1")
        ws[f"{start_col}1"] = "安全库存"
        ws[f"{start_col}1"].alignment = Alignment(horizontal="center", vertical="center")
    except Exception as e:
        st.error(f"⚠️ 安全库存表头合并失败: {e}")

def append_unfulfilled_summary_columns_by_date(main_plan_df: pd.DataFrame, df_unfulfilled: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """
    将未交订单按预交货日分为历史与未来月份，并添加至主计划 DataFrame。
    返回合并后的主计划表和未匹配品名列表（df_unfulfilled 中存在但主计划中没有的）。
    """
    today = pd.Timestamp(datetime.today().replace(day=1))
    final_month = pd.Timestamp("2025-11-01")
    future_months = pd.period_range(today.to_period("M"), final_month.to_period("M"), freq="M")
    future_cols = [f"未交订单 {str(p)}" for p in future_months]

    df = df_unfulfilled.copy()
    df["预交货日"] = pd.to_datetime(df["预交货日"], errors="coerce")
    df["未交订单数量"] = pd.to_numeric(df["未交订单数量"], errors="coerce").fillna(0)
    df["品名"] = df["品名"].astype(str).str.strip()
    df["月份"] = df["预交货日"].dt.to_period("M")

    # 合并重复品名+月份记录
    df = df.groupby(["品名", "月份"], as_index=False)["未交订单数量"].sum()
    df["是否历史"] = df["月份"] < today.to_period("M")

    df_hist = df[df["是否历史"]].groupby("品名", as_index=False)["未交订单数量"].sum()
    df_hist = df_hist.rename(columns={"未交订单数量": "历史未交订单"})

    df_future = df[~df["是否历史"]].copy()
    df_future["月份"] = df_future["月份"].astype(str)
    df_pivot = df_future.pivot_table(index="品名", columns="月份", values="未交订单数量", aggfunc="sum").fillna(0)
    df_pivot.columns = [f"未交订单 {col}" for col in df_pivot.columns]
    df_pivot = df_pivot.reset_index()

    for col in future_cols:
        if col not in df_pivot.columns:
            df_pivot[col] = 0

    df_merged = pd.merge(df_hist, df_pivot, on="品名", how="outer").fillna(0)
    df_merged["总未交订单"] = df_merged["历史未交订单"] + df_merged[future_cols].sum(axis=1)

    ordered_cols = ["品名", "总未交订单", "历史未交订单"] + future_cols
    df_merged = df_merged[ordered_cols]

    main_plan_df["品名"] = main_plan_df["品名"].astype(str).str.strip()
    result = pd.merge(main_plan_df, df_merged, on="品名", how="left")

    for col in ordered_cols[1:]:
        if col in result.columns:
            result[col] = result[col].fillna(0)

    # ✅ 未匹配品名（df_unfulfilled 中有，但 main_plan_df 中没有）
    all_unfulfilled_names = set(df_unfulfilled["品名"].dropna().astype(str).str.strip())
    all_main_names = set(main_plan_df["品名"].dropna().astype(str).str.strip())
    unmatched = sorted(list(all_unfulfilled_names - all_main_names))

    return result, unmatched

def merge_unfulfilled_order_header(sheet):
    """
    自动检测以“未交订单 ”开头的列，在第一行合并并写入“未交订单”，居中。
    
    参数:
    - sheet: openpyxl worksheet 对象
    """
    # 第2行是列名行（默认 DataFrame 用 dataframe_to_rows 写入时）
    header_row = list(sheet.iter_rows(min_row=2, max_row=2, values_only=True))[0]

    # 找出所有“未交订单 yyyy-mm”列的索引
    unfulfilled_cols = [
        idx for idx, col in enumerate(header_row, start=1)
        if isinstance(col, str) and col.startswith("未交订单 ")
    ]

    if not unfulfilled_cols:
        return  # 没有未交订单列，不处理

    start_col = min(unfulfilled_cols)
    end_col = max(unfulfilled_cols)

    # 合并单元格范围
    merge_range = f"{get_column_letter(start_col)}1:{get_column_letter(end_col)}1"
    sheet.merge_cells(merge_range)

    # 设置合并单元格的值与居中格式
    cell = sheet.cell(row=1, column=start_col)
    cell.value = "未交订单"
    cell.alignment = Alignment(horizontal="center", vertical="center")


def append_forecast_to_summary(summary_df: pd.DataFrame, forecast_df: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """
    从预测表中提取当月及未来的预测信息（仅按“品名”匹配），合并至 summary_df。
    返回合并后的表格和未匹配的品名列表。

    参数:
    - summary_df: 主计划 DataFrame（需含 '品名'）
    - forecast_df: 原始预测表（需含 '生产料号' 及预测列）

    返回:
    - result: 合并后的 DataFrame
    - unmatched_keys: list[str]，未匹配的品名
    """
    today = datetime.today()
    this_month_int = today.month  

    # ✅ 统一列名
    forecast_df = forecast_df.rename(columns={"生产料号": "品名"}).copy()
    forecast_df["品名"] = forecast_df["品名"].astype(str).str.strip()

    # ✅ 识别预测列（仅保留“x月预测”且月份 >= 当前月）    
    # 获取所有“x月预测”列，且月份合法
    month_cols = [
        col for col in forecast_df.columns
        if isinstance(col, str) and col.endswith("月预测") and "月" in col and col[:col.index("月")].isdigit()
    ]
    
    # 保留当前月及以后的预测列
    future_month_cols = [
        col for col in month_cols
        if int(col[:col.index("月")]) >= this_month_int
    ]


    if not future_month_cols:
        st.warning("⚠️ 未找到当月或未来月份的预测列（格式应为“5月预测”）")
        return summary_df, []

    # ✅ 汇总相同品名的预测值
    forecast_df[future_month_cols] = forecast_df[future_month_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    forecast_grouped = forecast_df.groupby("品名", as_index=False)[future_month_cols].sum()

    # ✅ 合并到主计划
    summary_df["品名"] = summary_df["品名"].astype(str).str.strip()
    result = summary_df.merge(forecast_grouped, on="品名", how="left")

    # ✅ 填补新预测列中的 NaN 为 0（不影响原有列）
    for col in future_month_cols:
        if col in result.columns:
            result[col] = result[col].fillna(0)

    # ✅ 找出未匹配品名
    forecast_keys = set(forecast_grouped["品名"])
    summary_keys = set(summary_df["品名"])
    unmatched_keys = sorted(list(forecast_keys - summary_keys))

    return result, unmatched_keys

def merge_forecast_header(sheet):
    """
    自动检测以“月预测”结尾的列（如“6月预测”、“7月预测”），
    在第一行合并这些列的单元格并写入“预测”，设置居中。
    """
    header_row = list(sheet.iter_rows(min_row=2, max_row=2, values_only=True))[0]

    # 找到所有“月预测”结尾的列索引
    forecast_cols = [
        idx for idx, col in enumerate(header_row, start=1)
        if isinstance(col, str) and col.endswith("月预测")
    ]

    if not forecast_cols:
        return  # 没有预测列，不处理

    start_col = min(forecast_cols)
    end_col = max(forecast_cols)

    # 合并单元格
    merge_range = f"{get_column_letter(start_col)}1:{get_column_letter(end_col)}1"
    sheet.merge_cells(merge_range)

    # 设置内容与样式
    cell = sheet.cell(row=1, column=start_col)
    cell.value = "预测"
    cell.alignment = Alignment(horizontal="center", vertical="center")
    

def merge_finished_inventory_with_warehouse_types(summary_df: pd.DataFrame, finished_inventory_df: pd.DataFrame, mapping_df: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """
    从成品库存中提取“HOLD仓”、“成品仓”、“半成品仓”数据，根据“品名”合并到 summary_df 中；
    若 mapping_df 中存在“半成品”→“新品名”映射，也一并添加半成品库存数据。

    返回合并后的 DataFrame 与未匹配的品名列表。
    """
    warehouse_cols = ["HOLD仓", "成品仓", "半成品仓"]

    # 初始化三列
    for col in warehouse_cols:
        if col not in summary_df.columns:
            summary_df[col] = 0

    # 清洗数据
    finished_inventory_df = finished_inventory_df.copy()
    finished_inventory_df["品名"] = finished_inventory_df["品名"].astype(str).str.strip()
    finished_inventory_df["仓库名称"] = finished_inventory_df["仓库名称"].astype(str).str.strip()
    finished_inventory_df["数量"] = pd.to_numeric(finished_inventory_df["数量"], errors="coerce").fillna(0)

    # 分组聚合：每个品名在每个仓库的总数量
    grouped = finished_inventory_df.groupby(["品名", "仓库名称"], as_index=False)["数量"].sum()

    # 合并进主计划
    summary_df["品名"] = summary_df["品名"].astype(str).str.strip()
    for _, row in grouped.iterrows():
        pname, warehouse, qty = row["品名"], row["仓库名称"], row["数量"]
        if warehouse in warehouse_cols and pname in summary_df["品名"].values:
            summary_df.loc[summary_df["品名"] == pname, warehouse] += qty

    # 处理“半成品 → 新品名”映射：将“半成品”的半成品仓，加到新品名对应行的“半成品仓”
    mapping_df = mapping_df.copy()
    mapping_df["半成品"] = mapping_df["半成品"].astype(str).str.strip()
    mapping_df["新品名"] = mapping_df["新品名"].astype(str).str.strip()
    half_mappings = mapping_df[mapping_df["半成品"] != ""]

    for _, row in half_mappings.iterrows():
        old_name = row["半成品"]
        new_name = row["新品名"]
        if old_name in summary_df["品名"].values and new_name in summary_df["品名"].values:
            delta = summary_df.loc[summary_df["品名"] == old_name, "半成品仓"].sum()
            summary_df.loc[summary_df["品名"] == new_name, "半成品仓"] += delta

    # 查找未匹配品名（只查成品库存里的）
    unmatched = sorted(list(set(grouped["品名"]) - set(summary_df["品名"])))

    return summary_df, unmatched


def merge_inventory_header(sheet):
    """
    合并“HOLD仓”、“成品仓”、“半成品仓”标题，写入“库存”，居中
    """
    header_row = list(sheet.iter_rows(min_row=2, max_row=2, values_only=True))[0]

    inventory_cols = [
        idx for idx, col in enumerate(header_row, start=1)
        if col in ["HOLD仓", "成品仓", "半成品仓"]
    ]
    if not inventory_cols:
        return

    start_col = min(inventory_cols)
    end_col = max(inventory_cols)
    sheet.merge_cells(f"{get_column_letter(start_col)}1:{get_column_letter(end_col)}1")
    cell = sheet.cell(row=1, column=start_col)
    cell.value = "库存"
    cell.alignment = Alignment(horizontal="center", vertical="center")


