import pandas as pd
from datetime import datetime
from config import FILE_KEYWORDS, OUTPUT_FILENAME_PREFIX
from io import BytesIO

class PivotProcessor:
    def __init__(self):
        self.dataframes = {}

    def classify_files(self, uploaded_files):
        for file in uploaded_files:
            filename = file.name
            for key, keyword in FILE_KEYWORDS.items():
                if keyword in filename:
                    self.dataframes[key] = pd.read_excel(file)
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
