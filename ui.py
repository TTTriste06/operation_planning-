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
    
    # 用户手动输入月份（可为空）
    manual_month = st.text_input("📅 输入历史数据截止月份（格式: YYYY-MM，可留空表示不筛选）")
    if manual_month.strip():
        CONFIG["selected_month"] = manual_month.strip()
        st.write(CONFIG["selected_month"])
    else:
        CONFIG["selected_month"] = None

    # 🗓️ 生成起始月份选项（从一年前到六个月后）
    today = date.today()
    start_month = today - relativedelta(months=12)
    end_month = today + relativedelta(months=6)
    month_options = pd.date_range(start=start_month, end=end_month, freq="MS").to_list()
    
    # 📅 月份选择框
    selected_month = st.selectbox(
        "📆 请选择排产计划起始月份",
        month_options,
        format_func=lambda x: x.strftime("%Y年%m月")
    )
    CONFIG["selected_plan_month"] = selected_month

        
    uploaded_files = st.file_uploader(
        "📂 上传 5 个核心 Excel 文件（未交订单/成品在制/成品库存/晶圆库存/CP在制）",
        type=["xlsx"],
        accept_multiple_files=True,
        key="main_files"
    )

    uploaded_dict = {}    
    for file in uploaded_files:
        uploaded_dict[file.name] = file

    # 输出上传文件名调试
    st.write("✅ 已上传文件名：", list(uploaded_dict.keys()))

    # 额外上传的 3 个文件
    st.subheader("📁 上传额外文件（可用储存的文件）")
    forecast_file = st.file_uploader("📈 上传预测文件", type="xlsx", key="forecast")
    safety_file = st.file_uploader("🔐 上传安全库存文件", type="xlsx", key="safety")
    mapping_file = st.file_uploader("🔁 上传新旧料号对照表", type="xlsx", key="mapping")
  

    start = st.button("🚀 生成汇总 Excel")
    return uploaded_dict, forecast_file, safety_file, mapping_file, start
