import pandas as pd
import re
import streamlit as st
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
