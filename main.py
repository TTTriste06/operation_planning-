import streamlit as st
from io import BytesIO
from datetime import datetime
import pandas as pd

from pivot_processor import PivotProcessor
from ui import setup_sidebar, get_uploaded_files
from github_utils import load_file_with_github_fallback
from urllib.parse import quote

def main():
    st.set_page_config(page_title="Excel数据透视汇总工具", layout="wide")
    setup_sidebar()

    # 获取上传文件
    uploaded_files, forecast_file, safety_file, mapping_file, pc_file, selected_date, start = get_uploaded_files()
    
    if start:            
        if len(uploaded_files) < 8:
            st.error("❌ 请上传 8 个核心文件（未交订单/成品在制/成品库存/CP在制/晶圆库存/下单明细/销货明细/到货明细）！")
            return
            
        # 加载辅助表
        df_forecast = load_file_with_github_fallback("forecast", forecast_file)
        df_safety = load_file_with_github_fallback("safety", safety_file)
        df_mapping = load_file_with_github_fallback("mapping", mapping_file)
        df_pc = load_file_with_github_fallback("pc", pc_file)
        
        additional_sheets = {
            "赛卓-预测": df_forecast,
            "赛卓-安全库存": df_safety,
            "赛卓-新旧料号": df_mapping,
            "赛卓-供应商-PC": df_pc
        }

        # 初始化处理器
        buffer = BytesIO()
        processor = PivotProcessor()
        processor.set_additional_data(additional_sheets)
        processor.process(uploaded_files, buffer, additional_sheets, start_date=selected_date)

        # 下载文件按钮
        file_name = f"运营数据订单-在制-库存汇总报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.success("✅ 汇总完成！你可以下载结果文件：")
        st.download_button(
            label="📥 下载 Excel 汇总报告",
            data=buffer.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Sheet 预览
        try:
            buffer.seek(0)
            with pd.ExcelFile(buffer, engine="openpyxl") as xls:
                sheet_names = xls.sheet_names
                tabs = st.tabs(sheet_names)
                for i, sheet_name in enumerate(sheet_names):
                    try:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        with tabs[i]:
                            st.subheader(f"📄 {sheet_name}")
                            st.dataframe(df, use_container_width=True)
                    except Exception as e:
                        with tabs[i]:
                            st.error(f"❌ 无法读取工作表 `{sheet_name}`: {e}")
        except Exception as e:
            st.warning(f"⚠️ 无法预览生成的 Excel 文件：{e}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("❌ Streamlit app crashed:", e)
        traceback.print_exc()

