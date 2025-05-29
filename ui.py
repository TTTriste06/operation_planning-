import streamlit as st

def setup_sidebar():
    with st.sidebar:
        st.title("📤 Excel 数据处理与汇总")
        st.markdown("上传所需的六个 Excel 文件，系统将生成汇总后的“运营主计划”。")

def get_uploaded_files():
    uploaded_files = st.file_uploader("上传 6 个文件", type=["xlsx"], accept_multiple_files=True)
    start = st.button("✅ 生成运营主计划")
    return uploaded_files, start
