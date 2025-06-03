import pandas as pd
import re
import streamlit as st
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


def merge_safety_inventory(main_plan_df: pd.DataFrame, safety_df: pd.DataFrame):
    st.write(safety_df)
    safety_df = safety_df.rename(columns={"ProductionNO.": "品名"}).copy()
    st.write(safety_df)
    safety_df = safety_df[["品名", "InvWaf", "InvPart"]].drop_duplicates()
    st.write(safety_df)
    
    all_keys = set(safety_df["品名"].dropna().astype(str).str.strip())

    merged = main_plan_df.merge(safety_df, on="品名", how="left")

    used_keys = set(
        merged[~merged[["InvWaf", "InvPart"]].isna().all(axis=1)]["品名"]
        .dropna().astype(str).str.strip()
    )

    unmatched_keys = list(all_keys - used_keys)
    
    st.write(merged)
    
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
    except Exception as e:
        st.error(f"⚠️ 安全库存表头合并失败: {e}")
