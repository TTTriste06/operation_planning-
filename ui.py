import streamlit as st
import pandas as pd
from dateutil.relativedelta import relativedelta
from datetime import date
from datetime import datetime

def setup_sidebar():
    with st.sidebar:
        st.title("功能简介")
        st.markdown("---")
        st.markdown("- 起始日期可以改变预测、未交订单和投单计划的起始月份")
        st.markdown("- 8个主文件每次都必须上传，每个文件一定要包含对应的关键字")
        st.markdown("- 等上方加载的标识消失后再进行下载")

def get_uploaded_files():
    st.header("📤 Excel 数据处理与汇总")

    # 📅 添加主计划起始时间选择器
    st.subheader("📅 选择主计划起始时间")
    selected_date = st.date_input(
        "选择一个起始日期", 
        value=datetime.today()  # 默认选当月1号
    )

    # ✅ 合并上传框：所有主+明细文件统一上传
    st.subheader("📁 上传主文件")
    all_files = st.file_uploader(
        "关键字：未交订单/成品在制/成品库存/CP在制/晶圆库存/下单明细/销货明细/到货明细（支持多选）",
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

    # 📁 上传辅助文件
    st.subheader("📁 上传辅助文件（如无更新可跳过）")
    forecast_file = st.file_uploader("📈 上传预测文件", type="xlsx", key="forecast")
    safety_file = st.file_uploader("🔐 上传安全库存文件", type="xlsx", key="safety")
    mapping_file = st.file_uploader("🔁 上传新旧料号对照表", type="xlsx", key="mapping")
    pc_file = st.file_uploader("🔁 上传PC-供应商表", type="xlsx", key="pc")


    # 🚀 生成按钮
    start = st.button("🚀 生成汇总 Excel")

    return uploaded_files, forecast_file, safety_file, mapping_file, pc_file, selected_date, start
