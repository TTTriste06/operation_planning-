import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime

from config import FILE_KEYWORDS, OUTPUT_FILENAME_PREFIX, FIELD_MAPPINGS
from excel_utils import adjust_column_width
from mapping_utils import clean_mapping_headers, apply_mapping_and_merge, apply_extended_substitute_mapping


class PivotProcessor:
    def __init__(self):
        self.dataframes = {}
        self.additional_sheets = {}

    def classify_files(self, uploaded_files):
        """
        根据关键词识别上传的主数据文件，并赋予标准中文名称
        """
        for file in uploaded_files:
            filename = file.name
            for keyword, standard_name in FILE_KEYWORDS.items():
                if keyword in filename:
                    self.dataframes[standard_name] = pd.read_excel(file)
                    break

    def process(self):
        """
        处理上传数据，替换新旧料号、替代料号。
        """
        # 提取新旧料号
        mapping_df = self.additional_sheets.get("赛卓-新旧料号")
        if mapping_df is None or mapping_df.empty:
            raise ValueError("❌ 缺少新旧料号映射表，无法进行品名替换。")

        # 替换品名 要处理的所有表（主表 + 辅助表，除了 mapping 和 supplier）
        all_tables = {
            **self.dataframes,
            **{k: v for k, v in self.additional_sheets.items() if k not in ["赛卓-新旧料号", "赛卓-供应商-PC"]}
        }
    
        for sheet_name, df in all_tables.items():
            if sheet_name not in FIELD_MAPPINGS:
                st.warning(f"⚠️ {sheet_name} 没有在 FIELD_MAPPINGS 中注册，跳过替换")
                continue
    
            field_map = FIELD_MAPPINGS[sheet_name]
            if "品名" not in field_map:
                st.warning(f"⚠️ {sheet_name} 的 FIELD_MAPPINGS 中未定义 '品名' 映射，跳过")
                continue
    
            actual_name_col = field_map["品名"]
            if actual_name_col not in df.columns:
                st.warning(f"⚠️ {sheet_name} 中找不到列：{actual_name_col}，跳过")
                continue
    
            try:
                df, keys_main = apply_mapping_and_merge(df, mapping_df, field_map={"品名": actual_name_col})
                df, keys_sub = apply_extended_substitute_mapping(df, mapping_df, field_map={"品名": actual_name_col})
                # 更新处理后的表
                if sheet_name in self.dataframes:
                    self.dataframes[sheet_name] = df
                else:
                    self.additional_sheets[sheet_name] = df
            except Exception as e:
                st.error(f"❌ 替换 {sheet_name} 中的品名失败：{e}")

        """
        处理主计划表
        """
        # 新建主计划df
        headers = ["晶圆品名", "规格", "品名", "封装厂", "封装形式", "PC"]
        main_plan_df = pd.DataFrame(columns=headers)

        # 品名：提取未交订单和预测中的品名

    
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
    
        # ✅ 合并并去重（即使两个来源都空也不报错）
        all_names = pd.Series(name_unfulfilled + name_forecast).dropna().drop_duplicates().sort_values()
    
        # ✅ 创建空 DataFrame 并填入品名
        main_plan_df = main_plan_df.reindex(index=range(len(all_names)))
        if not all_names.empty:
            main_plan_df["品名"] = all_names.values
    
        return {"主计划": main_plan_df}


    def export_to_excel(self, sheet_dict: dict):
        """
        将多个 DataFrame 写入 Excel 文件，并将表头写在第 2 行（第 1 行为空）。
        """
        output = BytesIO()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{OUTPUT_FILENAME_PREFIX}_{timestamp}.xlsx"
    
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet_name, df in sheet_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1 if sheet_name == "主计划" else 0)
        
            workbook = writer.book
            if "主计划" in writer.sheets:
                worksheet = writer.sheets["主计划"]
                adjust_column_width(worksheet)

        output.seek(0)
        return filename, output



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
