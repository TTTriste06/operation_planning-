import pandas as pd
from io import BytesIO
from datetime import datetime
from config import FILE_KEYWORDS, OUTPUT_FILENAME_PREFIX
from mapping_utils import clean_mapping_headers

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
        # 🚧 TODO: 在此处添加你的数据处理逻辑
        result = pd.DataFrame({"示例列": ["此处将放置处理结果"]})
        return result

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
                cleaned = clean_mapping_headers(mapping_df)
                self.additional_sheets["赛卓-新旧料号"] = cleaned
                st.write(mapping_df)
            except Exception as e:
                raise ValueError(f"❌ 新旧料号表清洗失败：{e}")
