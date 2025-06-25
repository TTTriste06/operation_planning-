import streamlit as st
import pandas as pd
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

    # 📅 输入历史截止月份
    # manual_month = st.text_input("📅 输入历史数据截止月份（格式: YYYY-MM，可留空表示不筛选）")
    # CONFIG["selected_month"] = manual_month.strip() if manual_month.strip() else None

    # ✅ 合并上传框：所有主+明细文件统一上传
    all_files = st.file_uploader(
        "📁 上传主文件: 未交订单/成品在制/成品库存/CP在制/晶圆库存/下单明细/销货明细/到货明细（支持多选）",
        type=["xlsx"],
        accept_multiple_files=True,
        key="all_files"
    )

    # 将所有文件统一收集到 uploaded_files 字典
    uploaded_files = {}
    if all_files:
        for file in all_files:
            uploaded_files[file.name] = file
        st.success(f"✅ 共上传 {len(uploaded_files)} 个文件：")
        st.write(list(uploaded_files.keys()))
    else:
        st.info("📂 尚未上传文件。")

    # 📁 上传辅助文件（可选）
    st.subheader("📁 上传辅助文件（如无更新可跳过）")
    forecast_file = st.file_uploader("📈 上传预测文件", type="xlsx", key="forecast")
    safety_file = st.file_uploader("🔐 上传安全库存文件", type="xlsx", key="safety")
    mapping_file = st.file_uploader("🔁 上传新旧料号对照表", type="xlsx", key="mapping")
    pc_file = st.file_uploader("🔁 上传PC-供应商表", type="xlsx", key="pc")

    # 🚀 生成按钮
    start = st.button("🚀 生成汇总 Excel")

    return uploaded_files, forecast_file, safety_file, mapping_file, pc_file, start, selected_year, selected_month
