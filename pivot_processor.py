import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

from config import FILE_KEYWORDS, OUTPUT_FILENAME_PREFIX, FIELD_MAPPINGS
from excel_utils import adjust_column_width
from mapping_utils import (
    clean_mapping_headers, 
    replace_all_names_with_mapping, 
    apply_mapping_and_merge, 
    apply_extended_substitute_mapping
)
from data_utils import (
    extract_info, 
    fill_spec_and_wafer_info, 
    fill_packaging_info
)
from summary import (
    merge_safety_inventory,
    merge_safety_header
)

class PivotProcessor:
    def process(self, uploaded_files: dict, output_buffer, additional_sheets: dict = None):
        """
        替换品名、新建主计划表，并直接写入 Excel 文件（含列宽调整、标题行）。
        """
        
        # === 标准化上传文件名 ===
        self.dataframes = {}
        for filename, file_obj in uploaded_files.items():
            matched = False
            for keyword, standard_name in FILE_KEYWORDS.items():
                if keyword in filename:
                    self.dataframes[standard_name] = pd.read_excel(file_obj)
                    matched = True
                    break
            if not matched:
                st.warning(f"⚠️ 上传文件 `{filename}` 未识别关键词，跳过")
        
        self.additional_sheets = additional_sheets
        mapping_df = additional_sheets.get("赛卓-新旧料号")
        if mapping_df is None or mapping_df.empty:
            raise ValueError("❌ 缺少新旧料号映射表，无法进行品名替换。")
            
        # === 构建主计划 ===
        headers = ["晶圆品名", "规格", "品名", "封装厂", "封装形式", "PC"]
        main_plan_df = pd.DataFrame(columns=headers)

        ## == 品名 ==
        df_unfulfilled = self.dataframes.get("赛卓-未交订单")
        df_forecast = self.additional_sheets.get("赛卓-预测")

        name_unfulfilled = []
        name_forecast = []

        if df_unfulfilled is not None and not df_unfulfilled.empty:
            col_name = FIELD_MAPPINGS["赛卓-未交订单"]["品名"]
            name_unfulfilled = df_unfulfilled[col_name].astype(str).str.strip().tolist()

        if df_forecast is not None and not df_forecast.empty:
            col_name = FIELD_MAPPINGS["赛卓-预测"]["品名"]
            name_forecast = df_forecast[col_name].astype(str).str.strip().tolist()

        all_names = pd.Series(name_unfulfilled + name_forecast)
        all_names = replace_all_names_with_mapping(all_names, mapping_df)


        main_plan_df = main_plan_df.reindex(index=range(len(all_names)))
        if not all_names.empty:
            main_plan_df["品名"] = all_names.values

        ## == 规格和品名 ==
        main_plan_df = fill_spec_and_wafer_info(
            main_plan_df,
            self.dataframes,
            self.additional_sheets,
            FIELD_MAPPINGS
        )

        ## == 封装厂，封装形式和PC ==
        main_plan_df = fill_packaging_info(
            main_plan_df,
            dataframes=self.dataframes,
            additional_sheets=self.additional_sheets
        )

        ## == 替换新旧料号、替代料号 ==
        for sheet_name, df in {
            **self.dataframes,
            **{k: v for k, v in additional_sheets.items() if k not in ["赛卓-新旧料号", "赛卓-供应商-PC"]}
        }.items():
            if sheet_name not in FIELD_MAPPINGS:
                st.warning(f"⚠️ {sheet_name} 未在 FIELD_MAPPINGS 注册，跳过替换")
                continue
        
            field_map = FIELD_MAPPINGS[sheet_name]
            if "品名" not in field_map:
                st.warning(f"⚠️ {sheet_name} 的 FIELD_MAPPINGS 中未定义 '品名'，跳过")
                continue
        
            actual_name_col = field_map["品名"]
            if actual_name_col not in df.columns:
                st.warning(f"⚠️ {sheet_name} 中找不到列：{actual_name_col}，跳过")
                continue
        
            try:
                df, _ = apply_mapping_and_merge(df, mapping_df, field_map={"品名": actual_name_col})
                df, _ = apply_extended_substitute_mapping(df, mapping_df, field_map={"品名": actual_name_col})
        
                if sheet_name in self.dataframes:
                    self.dataframes[sheet_name] = df
                else:
                    self.additional_sheets[sheet_name] = df
            except Exception as e:
                st.error(f"❌ 替换 {sheet_name} 中的品名失败：{e}")

        ## == 安全库存 ==
        safety_df = additional_sheets.get("赛卓-安全库存")
        if safety_df is not None and not safety_df.empty:
            main_plan_df, unmatched_safety = merge_safety_inventory(main_plan_df, safety_df)
            st.success("✅ 已合并安全库存数据")
            if unmatched_safety:
                st.warning(f"⚠️ 以下品名未在安全库存中匹配到：{unmatched_safety}")

        
        # === 写入 Excel 文件（主计划）===
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            main_plan_df.to_excel(writer, sheet_name="主计划", index=False, startrow=1)
        
            ws = writer.book["主计划"]
            ws.cell(row=1, column=1, value=f"主计划生成时间：{timestamp}")
        
            merge_safety_header(ws, main_plan_df)  # 🔷 合并标题
            adjust_column_width(ws)

        output_buffer.seek(0)

        
    def set_additional_data(self, sheets_dict):
        """
        设置辅助数据表，如 预测、安全库存、新旧料号 等
        """
        self.additional_sheets = sheets_dict or {}
    
        # ✅ 对新旧料号进行列名清洗
        mapping_df = self.additional_sheets.get("赛卓-新旧料号")
        if mapping_df is not None and not mapping_df.empty:
            try:
                cleaned = clean_mapping_headers(mapping_df)
                self.additional_sheets["赛卓-新旧料号"] = cleaned
            except Exception as e:
                raise ValueError(f"❌ 新旧料号表清洗失败：{e}")
