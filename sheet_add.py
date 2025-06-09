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
    将 uploaded_files 和 additional_sheets 中所有内容写入 Excel：
    - 若是原始 Excel 文件：提取全部 sheet 写入；
    - 若是 DataFrame：直接写入；
    - sheet 命名支持自动识别并重命名；
    - 自动执行 NaN 清洗和列宽调整。
    """
    all_files = {**uploaded_files, **additional_sheets}

    # 🔁 自动重命名规则
    rename_map = {
        "未交订单": "赛卓-未交订单",
        "成品在制": "赛卓-成品在制",
        "成品库存": "赛卓-成品库存",
        "CP在制": "赛卓-CP在制",
        "晶圆库存": "赛卓-晶圆库存",
        "到货明细": "赛卓-到货明细",
        "下单明细": "赛卓-下单明细",
        "销货明细": "赛卓-销货明细"
    }

    for filename, file_obj in all_files.items():
        try:
            # ✅ Case 1: file_obj 是 DataFrame
            if isinstance(file_obj, pd.DataFrame):
                cleaned_df = clean_df(file_obj)
                sheet_name = filename[:31]
                for keyword, std_name in rename_map.items():
                    if keyword in sheet_name:
                        sheet_name = std_name
                        break
                cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)
                adjust_column_width(writer, sheet_name, cleaned_df)

            # ✅ Case 2: file_obj 是 Excel 文件对象，遍历其所有 sheet
            else:
                xls = pd.ExcelFile(file_obj)
                for sheet in xls.sheet_names:
                    df = xls.parse(sheet)
                    if isinstance(df, pd.DataFrame) and not df.empty:
                        cleaned_df = clean_df(df)
                        base_name = f"{filename[:15]}-{sheet[:15]}"
                        sheet_name = base_name
                        for keyword, std_name in rename_map.items():
                            if keyword in sheet:
                                sheet_name = std_name
                                break
                        sheet_name = sheet_name[:31]
                        cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        adjust_column_width(writer, sheet_name, cleaned_df)

        except Exception as e:
            print(f"❌ 读取或写入文件 [{filename}] 的 sheet 失败：{e}")
