import streamlit as st
from streamlit import file_uploader

def setup_sidebar():
    with st.sidebar:
        st.title(" ")

def get_uploaded_files():
    st.header("📤 Excel 数据处理与汇总")
    
    uploaded_core_files = file_uploader("📂 上传 6 个主数据文件", type=["xlsx"], accept_multiple_files=True)

    st.markdown("### 📎 上传 4 个辅助文件（可选，用于合并与匹配）")
    forecast_file = file_uploader("📗 上传预测文件（例如：预测.xlsx）", type=["xlsx"], key="forecast")
    safety_file = file_uploader("📙 上传安全库存文件", type=["xlsx"], key="safety")
    mapping_file = file_uploader("📘 上传新旧料号文件", type=["xlsx"], key="mapping")
    supplier_file = file_uploader("📕 上传供应商-PC 文件", type=["xlsx"], key="supplier")

    start = st.button("✅ 生成运营主计划")
    return uploaded_core_files, forecast_file, safety_file, mapping_file, supplier_file, start
