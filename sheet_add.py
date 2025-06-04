import pandas as pd
import re

def append_uploaded_sheets_to_excel_by_mapping(writer, uploaded_files: dict, field_mappings: dict):
    """
    将上传的原始文件写入 Excel，每个文件作为一个 sheet，名称为 FIELD_MAPPINGS 中定义的 key。
    
    参数:
    - writer: pd.ExcelWriter
    - uploaded_files: dict[str, file-like]
    - field_mappings: dict[str, dict]，其中 key 为目标 sheet 名
    """
    for sheet_name, mapping in field_mappings.items():
        # 在上传文件中找到与 sheet_name 对应的文件名（支持中文命名）
        match_file = None
        for file_name in uploaded_files:
            if sheet_name in file_name:
                match_file = uploaded_files[file_name]
                break

        if match_file:
            try:
                df = pd.read_excel(match_file, sheet_name=0)
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            except Exception as e:
                print(f"❌ 写入 {sheet_name} 失败: {e}")
        else:
            print(f"⚠️ 未找到匹配文件用于写入 sheet: {sheet_name}")
