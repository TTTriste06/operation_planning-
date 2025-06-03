import re
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

def merge_safety_inventory(summary_df: pd.DataFrame, safety_df: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """
    将安全库存表中 InvWaf 和 InvPart 信息按 '品名' 合并到汇总表中，仅根据 '品名' 匹配。
    如果未匹配，则填入 0（便于后续计算）。

    参数:
    - summary_df: 汇总后的 DataFrame，含 '品名'
    - safety_df: 安全库存表，含 'ProductionNO.'、'InvWaf'、'InvPart'

    返回:
    - merged: 合并后的 DataFrame（含 InvWaf 和 InvPart，空值填 0）
    - unmatched_keys: list of 未被匹配的品名
    """

    # ✅ 清洗并重命名列
    safety_df = safety_df.rename(columns={"ProductionNO.": "品名"}).copy()
    safety_df.columns = safety_df.columns.str.strip()
    safety_df["品名"] = safety_df["品名"].astype(str).str.strip()

    # 只保留需要列 + 去重
    safety_df = safety_df[["品名", "InvWaf", "InvPart"]].drop_duplicates()

    # 转换为数值型（空值也会变成 NaN，稍后会填为 0）
    safety_df["InvWaf"] = pd.to_numeric(safety_df["InvWaf"], errors="coerce")
    safety_df["InvPart"] = pd.to_numeric(safety_df["InvPart"], errors="coerce")

    # 清洗主计划的品名
    summary_df["品名"] = summary_df["品名"].astype(str).str.strip()

    # ✅ 合并
    merged = summary_df.merge(safety_df, on="品名", how="left")

    # ✅ 记录未匹配品名
    matched_keys = set(safety_df["品名"])
    used_keys = set(merged[~merged[["InvWaf", "InvPart"]].isna().all(axis=1)]["品名"])
    unmatched_keys = list(matched_keys - used_keys)

    # ✅ 空值填 0，便于后续计算
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

def append_unfulfilled_summary_columns_by_date(main_plan_df: pd.DataFrame, df_unfulfilled: pd.DataFrame) -> pd.DataFrame:
    """
    将未交订单按预交货日分为历史与未来月份，并添加至主计划 DataFrame。

    参数:
    - main_plan_df: 主计划 DataFrame，需含 '品名'
    - df_unfulfilled: 未交订单 DataFrame，含 '品名', '未交订单数量', '预交货日'

    返回:
    - main_plan_df: 合并后的 DataFrame，包含“总未交订单”、“历史未交订单”、“未交订单 yyyy-mm”等列
    """
    today = pd.Timestamp(datetime.today().replace(day=1))  # 当前月份的第一天
    final_month = pd.Timestamp("2025-11-01")  # 最晚月份（可根据需求修改）
    future_months = pd.period_range(today.to_period("M"), final_month.to_period("M"), freq="M")
    future_cols = [f"未交订单 {str(p)}" for p in future_months]

    # 清洗数据
    df = df_unfulfilled.copy()
    df["预交货日"] = pd.to_datetime(df["预交货日"], errors="coerce")
    df["未交订单数量"] = pd.to_numeric(df["未交订单数量"], errors="coerce").fillna(0)
    df["品名"] = df["品名"].astype(str).str.strip()
    df["月份"] = df["预交货日"].dt.to_period("M")

    # 历史未交订单
    df["是否历史"] = df["月份"] < today.to_period("M")
    df_hist = df[df["是否历史"]].groupby("品名")["未交订单数量"].sum().rename("历史未交订单").reset_index()

    # 未来未交订单
    df_future = df[~df["是否历史"]].copy()
    df_future["月份"] = df_future["月份"].astype(str)
    df_pivot = df_future.pivot_table(index="品名", columns="月份", values="未交订单数量", aggfunc="sum").fillna(0)
    df_pivot.columns = [f"未交订单 {col}" for col in df_pivot.columns]
    df_pivot = df_pivot.reset_index()

    # 确保列完整
    for col in future_cols:
        if col not in df_pivot.columns:
            df_pivot[col] = 0

    # 合并历史与未来
    df_merged = pd.merge(df_hist, df_pivot, on="品名", how="outer").fillna(0)

    # 总未交订单
    df_merged["总未交订单"] = df_merged["历史未交订单"] + df_merged[future_cols].sum(axis=1)

    # 整理列顺序
    ordered_cols = ["品名", "总未交订单", "历史未交订单"] + future_cols
    df_merged = df_merged[ordered_cols]

    # 合并进主计划表
    main_plan_df["品名"] = main_plan_df["品名"].astype(str).str.strip()
    result = pd.merge(main_plan_df, df_merged, on="品名", how="left").fillna(0)

    return result

