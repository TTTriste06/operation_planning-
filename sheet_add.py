import re
import pandas as pd
import streamlit as st
from config import FIELD_MAPPINGS

def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    将 DataFrame 中的 NaN 替换为空字符串，并移除全为空的列（可选）。
    """
    return df.fillna("").replace("nan", "")  # 防止 "nan" 字符串也被写入


def append_all_standardized_sheets(writer: pd.ExcelWriter, main_tables: dict, additional_tables: dict):
    """
    将 main_tables + additional_tables 中所有标准化命名的表写入 Excel，每个作为一个 sheet。

    参数：
    - writer: pd.ExcelWriter
    - main_tables: dict，通常为 self.dataframes（标准化后的主数据）
    - additional_tables: dict，通常为 self.additional_sheets（辅助数据）
    """
    combined_tables = {**main_tables, **additional_tables}

    for sheet_name, df in combined_tables.items():
        try:
            if isinstance(df, pd.DataFrame) and not df.empty:
                cleaned_df = clean_df(df)
                cleaned_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        except Exception as e:
            print(f"❌ 写入 sheet [{sheet_name}] 失败：{e}")
