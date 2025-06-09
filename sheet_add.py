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

def append_all_standardized_sheets(writer: pd.ExcelWriter, 
                                            uploaded_files: dict, 
                                            additional_sheets: dict):
    """
    从原始上传文件（uploaded_files 和 additional_sheets）中直接读取未被更改的 Excel 内容，
    并写入每个 Sheet（NaN 清洗 + 自动列宽）。
    """
    all_files = {**uploaded_files, **additional_sheets}

    for filename, file_obj in all_files.items():
        try:
            # 尝试读取该文件中的所有 sheet
            xls = pd.ExcelFile(file_obj)
            for sheet in xls.sheet_names:
                df = xls.parse(sheet)
                if isinstance(df, pd.DataFrame) and not df.empty:
                    cleaned_df = clean_df(df)
                    safe_sheet_name = f"{filename[:15]}-{sheet[:15]}"[:31]  # 避免超过限制
                    cleaned_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                    adjust_column_width(writer, safe_sheet_name, cleaned_df)
        except Exception as e:
            print(f"❌ 读取或写入文件 [{filename}] 的 sheet 失败：{e}")
