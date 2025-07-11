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
    all_files = {**additional_sheets, **uploaded_files}
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
            # 遍历 rename_map 的 key，看是否被包含在 filename 中
            for key, new_name in rename_map.items():
                if key in filename:
                    filename = new_name
                    break  # 命中后退出

            if isinstance(file_obj, pd.DataFrame):
                cleaned_df = clean_df(file_obj)
                sheet_name = filename[:31]
                cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)
                adjust_column_width(writer, sheet_name, cleaned_df)
            else:
                xls = pd.ExcelFile(file_obj)
                for sheet in xls.sheet_names:
                    df = xls.parse(sheet)
                    if isinstance(df, pd.DataFrame) and not df.empty:
                        cleaned_df = clean_df(df)
                        safe_sheet_name = filename[:31]
                        cleaned_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                        adjust_column_width(writer, safe_sheet_name, cleaned_df)
        except Exception as e:
            print(f"❌ 读取或写入文件 [{filename}] 的 sheet 失败：{e}")


def append_original_cp_sheets(writer, cp_dataframes: dict):
    """
    将所有原始 CP 文件表追加为带格式的 sheet，命名为 FAB_WIP_上华1厂、FAB_WIP_DB 等
    """
    fab_name_map = {
        "上华1": "上华1厂",
        "上华2": "上华2厂",
        "上华5": "上华5厂",
        "DB": "DB",
        "华虹": "华虹",
        "先进": "先进积塔"
    }

    for key, df in cp_dataframes.items():
        # 自动识别厂名
        matched_name = None
        for prefix in fab_name_map:
            if key.startswith(prefix):
                matched_name = fab_name_map[prefix]
                break

        # 如果匹配不到就跳过或使用原名
        if not matched_name:
            st.warning(f"⚠️ 未识别的厂名称：{key}，将使用原始名")
            matched_name = key

        sheet_name = f"FAB_WIP_{matched_name}"
        clean_name = sheet_name[:31]  # Excel sheet name 限长

        try:
            df.to_excel(writer, sheet_name=clean_name, index=False)
        except Exception as e:
            st.warning(f"⚠️ 写入 sheet `{clean_name}` 时出错：{e}")
