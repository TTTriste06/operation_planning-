import re
import pandas as pd
import streamlit as st
from config import FIELD_MAPPINGS

def append_all_source_sheets_to_excel(writer, uploaded_files: dict, additional_sheets: dict):
    """
    将上传的原始主文件 + 额外辅助文件（additional_sheets）都写入 Excel 文件。
    使用键名作为 sheet 名（如 “赛卓-预测”）。
    """
    all_sources = {}

    # ✅ 收集 uploaded_files 中含有中文关键词的文件
    for key, file in uploaded_files.items():
        st.write(key)
        for std_name in FIELD_MAPPINGS.keys():
            if std_name in key:
                all_sources[std_name] = file
                break

    # ✅ 收集 additional_sheets（如“赛卓-预测”、“赛卓-安全库存”等 DataFrame）
    for key, df in additional_sheets.items():
        st.write(key)
        if key not in all_sources:  # 不重复写
            all_sources[key] = df

    st.write("all")
    # ✅ 统一写入 Excel
    for sheet_name, content in all_sources.items():
        st.write(sheet_name)
        try:
            if isinstance(content, pd.DataFrame):
                content.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            else:
                df = pd.read_excel(content, sheet_name=0)
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        except Exception as e:
            print(f"❌ 写入 sheet [{sheet_name}] 失败: {e}")
