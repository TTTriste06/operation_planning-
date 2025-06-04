import pandas as pd
import re

def append_uploaded_sheets_to_excel_by_mapping(writer, uploaded_files: dict):
    """
    将标准命名的 12 个上传文件写入 Excel，每个为一个 sheet，使用中文名作为 sheet 名。
    要求文件名包含以下任一 key：
    - "赛卓-未交订单", "赛卓-成品在制", ..., "赛卓-供应商-PC"
    """
    expected_sheet_names = [
        "赛卓-未交订单", "赛卓-成品在制", "赛卓-成品库存", "赛卓-CP在制", "赛卓-晶圆库存",
        "赛卓-到货明细", "赛卓-下单明细", "赛卓-销货明细", "赛卓-新旧料号",
        "赛卓-预测", "赛卓-安全库存", "赛卓-供应商-PC"
    ]

    for sheet_name in expected_sheet_names:
        matched_file = None
        for filename in uploaded_files:
            if sheet_name in filename:
                matched_file = uploaded_files[filename]
                break

        if matched_file:
            try:
                df = pd.read_excel(matched_file, sheet_name=0)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            except Exception as e:
                print(f"❌ 无法写入 sheet [{sheet_name}]，错误: {e}")
        else:
            print(f"⚠️ 未找到匹配文件写入 sheet: {sheet_name}")
