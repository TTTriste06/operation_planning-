import streamlit as st
from ui import setup_sidebar, get_uploaded_files
from pivot_processor import PivotProcessor
from github_utils import upload_to_github, load_or_fallback_from_github

def main():
    st.set_page_config(page_title="运营主计划生成器", layout="wide")
    setup_sidebar()

    uploaded_core_files, forecast_file, safety_file, mapping_file, supplier_file, start = get_uploaded_files()

    if start:
        if len(uploaded_core_files) < 6:
            st.error("❌ 请上传 6 个主要文件")
            return

        github_files = {
            "预测.xlsx": forecast_file,
            "安全库存.xlsx": safety_file,
            "新旧料号.xlsx": mapping_file,
            "供应商-PC.xlsx": supplier_file
        }

        additional_sheets = {}

        for name, file in github_files.items():
        if file:
            file_bytes = file.read()
            upload_to_github(BytesIO(file_bytes), name)
            additional_sheets[name.split(".")[0]] = pd.read_excel(BytesIO(file_bytes))
        else:
            df = load_or_fallback_from_github(key=reverse_lookup(name))
            additional_sheets[name.split(".")[0]] = df

        processor = PivotProcessor()
        processor.classify_files(uploaded_core_files)

        # 将附加数据传入处理器
        processor.set_additional_data(additional_sheets)

        result_df = processor.process()
        
        filename, output = processor.export_to_excel(result_df)

        st.success(f"✅ 成功生成：{filename}")
        st.download_button("📥 下载 Excel 文件", data=output, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
