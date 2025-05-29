import streamlit as st
from ui import setup_sidebar, get_uploaded_files
from pivot_processor import PivotProcessor

def main():
    st.set_page_config(page_title="运营主计划生成器", layout="wide")
    setup_sidebar()

    uploaded_files, start = get_uploaded_files()

    if start:
        if len(uploaded_files) < 6:
            st.error("❌ 请上传所有 6 个主要文件再点击生成！")
            return

        processor = PivotProcessor()
        processor.classify_files(uploaded_files)

        result_df = processor.process()
        filename, output = processor.export_to_excel(result_df)

        st.success(f"✅ 成功生成：{filename}")
        st.download_button("📥 下载 Excel 文件", data=output, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
