import re
import pandas as pd
import streamlit as st
from config import FIELD_MAPPINGS
from excel_utils import adjust_column_width

def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    清洗 DataFrame：
    - 将 NaN 和 'nan' 替换为空字符串；
    - 去除字符串前后空格；
    """
    df = df.fillna("").replace("nan", "")
    df = df.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)
    return df


def append_all_standardized_sheets(writer: pd.ExcelWriter, main_tables: dict, additional_tables: dict):
    """
    将 main_tables 和 additional_tables 中的所有标准化命名的 DataFrame 写入 Excel Sheet，
    并对每个 Sheet 自动执行 NaN 清洗 + 列宽调整。
    """
    combined_tables = {**main_tables, **additional_tables}

    for sheet_name, df in combined_tables.items():
        try:
            if isinstance(df, pd.DataFrame) and not df.empty:
                cleaned_df = clean_df(df)
                cleaned_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                adjust_column_width(writer, sheet_name, cleaned_df)
        except Exception as e:
            print(f"❌ 写入 sheet [{sheet_name}] 失败：{e}")
