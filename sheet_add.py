import re
import pandas as pd
import streamlit as st
from config import FIELD_MAPPINGS
from openpyxl.utils import get_column_letter

def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    清洗 DataFrame：
    - 将 NaN 和 'nan' 替换为空字符串；
    - 去除字符串前后空格；
    """
    df = df.fillna("").replace("nan", "")
    df = df.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)
    return df

def adjust_column_width(writer, sheet_name: str, df):
    """
    自动调整指定 sheet 的列宽，使每列适应其内容长度。
    
    参数:
    - writer: pd.ExcelWriter 实例（engine='openpyxl'）
    - sheet_name: str，目标工作表名称
    - df: 原始写入的 DataFrame，用于列宽计算
    """
    ws = writer.book[sheet_name]
    
    for i, col in enumerate(df.columns, 1):  # 1-based indexing
        max_len = max(
            df[col].astype(str).map(len).max(),
            len(str(col))  # header 长度
        )
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = max_len + 2  # 适度留白


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
