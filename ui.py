import streamlit as st
import pandas as pd
from config import CONFIG
from dateutil.relativedelta import relativedelta
from datetime import date


def setup_sidebar():
    with st.sidebar:
        st.title("欢迎使用数据汇总工具")
        st.markdown("---")
        st.markdown("### 功能简介：")
        st.markdown("- 上传 5 个主数据表")
        st.markdown("- 上传辅助数据（预测、安全库存、新旧料号）")
        st.markdown("- 自动生成汇总 Excel 文件")

def get_uploaded_files():
    st.header("📤 Excel 数据处理与汇总")

    # 📅 手动输入历史截止月份
    manual_month = st.text_input("📅 输入历史数据截止月份（格式: YYYY-MM，可留空表示不筛选）")
    CONFIG["selected_month"] = manual_month.strip() if manual_month.strip() else None
    
    # 📂 上传主要文件
    uploaded_files = st.file_uploader(
        "📂 上传 5 个核心 Excel 文件（未交订单/成品在制/成品库存/晶圆库存/CP在制）",
        type=["xlsx"],
        accept_multiple_files=True,
        key="main_files"
    )
    uploaded_dict = {file.name: file for file in uploaded_files}
    st.write("✅ 已上传主文件：", list(uploaded_dict.keys()))

    # 📁 上传辅助文件
    st.subheader("📁 上传辅助文件（如无更新可跳过）")
    forecast_file = st.file_uploader("📈 上传预测文件", type="xlsx", key="forecast")
    safety_file = st.file_uploader("🔐 上传安全库存文件", type="xlsx", key="safety")
    mapping_file = st.file_uploader("🔁 上传新旧料号对照表", type="xlsx", key="mapping")

    # 📦 上传扩展文件（新增 3 个）
    st.subheader("📦 上传运营额外明细文件")
    arrival_file = st.file_uploader("🚚 上传到货明细", type="xlsx", key="arrival")
    order_file = st.file_uploader("📝 上传下单明细", type="xlsx", key="order")
    sales_file = st.file_uploader("💰 上传销货明细", type="xlsx", key="sales")

    # 🚀 生成按钮
    start = st.button("🚀 生成汇总 Excel")

    return uploaded_dict, forecast_file, safety_file, mapping_file, arrival_file, order_file, sales_file, start
