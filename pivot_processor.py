import pandas as pd
from io import BytesIO
from datetime import datetime
from config import FILE_KEYWORDS, OUTPUT_FILENAME_PREFIX
from mapping_utils import clean_mapping_headers
from data_utils import extract_required_columns

class PivotProcessor:
    def __init__(self):
        self.dataframes = {}
        self.additional_sheets = {}

    def classify_files(self, uploaded_files):
        """
        根据关键词识别上传的主数据文件，并赋予标准中文名称
        """
        for file in uploaded_files:
            filename = file.name
            for keyword, standard_name in FILE_KEYWORDS.items():
                if keyword in filename:
                    self.dataframes[standard_name] = pd.read_excel(file)
                    break

    def process(self):
        """
        执行实际数据处理逻辑：提取关键字段、格式化数据，为后续透视或合并做准备。
        """
        extracted = {}
    
        # 主数据处理
        for sheet_name, df in self.dataframes.items():
            try:
                df_extracted = extract_required_columns(sheet_name, df)
                extracted[sheet_name] = df_extracted
            except Exception as e:
                print(f"❌ 提取 `{sheet_name}` 失败: {e}")
    
        # 辅助数据处理（如预测、安全库存）
        for sheet_name, df in self.additional_sheets.items():
            try:
                df_extracted = extract_required_columns(sheet_name, df)
                extracted[sheet_name] = df_extracted
            except Exception as e:
                print(f"❌ 提取辅助 `{sheet_name}` 失败: {e}")
    
        # 👉 可将提取后的结果合并成汇总表或分别处理
        result_summary = pd.DataFrame({
            "表名": list(extracted.keys()),
            "行数": [len(df) for df in extracted.values()],
            "字段": [", ".join(df.columns) for df in extracted.values()]
        })
        st.write(result_summary)
    
        return result_summary


    def export_to_excel(self, df):
        output = BytesIO()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{OUTPUT_FILENAME_PREFIX}_{timestamp}.xlsx"
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="运营主计划", index=False)
        output.seek(0)
        return filename, output

    def set_additional_data(self, sheets_dict):
        """
        设置辅助数据表，如 预测、安全库存、新旧料号 等
        """
        self.additional_sheets = sheets_dict or {}
    
        # ✅ 对新旧料号进行列名清洗
        mapping_df = self.additional_sheets.get("赛卓-新旧料号")
        if mapping_df is not None and not mapping_df.empty:
            try:
                st.write(mapping_df)
                cleaned = clean_mapping_headers(mapping_df)
                self.additional_sheets["赛卓-新旧料号"] = cleaned
                st.write(mapping_df)
            except Exception as e:
                raise ValueError(f"❌ 新旧料号表清洗失败：{e}")
