import re
import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

from config import FILE_KEYWORDS, OUTPUT_FILENAME_PREFIX, FIELD_MAPPINGS, pivot_config, RENAME_MAP
from excel_utils import adjust_column_width
from mapping_utils import (
    clean_mapping_headers, 
    replace_all_names_with_mapping, 
    apply_mapping_and_merge, 
    apply_extended_substitute_mapping,
    apply_all_name_replacements
)
from data_utils import (
    extract_info, 
    fill_spec_and_wafer_info, 
    fill_packaging_info
)
from summary import (
    merge_safety_inventory,
    merge_safety_header,
    append_unfulfilled_summary_columns_by_date,
    merge_unfulfilled_order_header, 
    append_forecast_to_summary,
    merge_forecast_header,
    merge_finished_inventory_with_warehouse_types,
    merge_inventory_header,
    append_product_in_progress,
    merge_product_in_progress_header
)
from production_plan import (
    init_monthly_fields,
    generate_monthly_fg_plan,
    aggregate_actual_fg_orders,
    aggregate_actual_sfg_orders,
    aggregate_actual_arrivals,
    aggregate_sales_quantity_and_amount,
    generate_monthly_semi_plan,
    generate_monthly_adjust_plan,
    generate_monthly_return_adjustment,
    generate_monthly_return_plan,
    format_monthly_grouped_headers,
    highlight_production_plan_cells,
    drop_last_forecast_month_columns
)
from sheet_add import clean_df, append_all_standardized_sheets
from pivot_generator import generate_monthly_pivots, standardize_uploaded_keys

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
        mapping_df = self.additional_sheets.get("赛卓-新旧料号")
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

        ## == 规格和晶圆 ==
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
        st.write(self.additional_sheets["赛卓-安全库存"])
        df_new = self.additional_sheets["赛卓-安全库存"]
        st.write(df_new)
        df_new, _ = apply_mapping_and_merge(df_new, mapping_df, FIELD_MAPPINGS["赛卓-安全库存"])
        st.write(self.additional_sheets["赛卓-安全库存"])
        df_new, _ = apply_extended_substitute_mapping(df_new, mapping_df, FIELD_MAPPINGS["赛卓-安全库存"])
        additional_sheets["赛卓-安全库存"] = df_new

        st.write(self.additional_sheets["赛卓-安全库存"])

        df_new = self.additional_sheets["赛卓-预测"]
        df_new, _ = apply_mapping_and_merge(df_new, mapping_df, FIELD_MAPPINGS["赛卓-预测"])
        df_new, _ = apply_extended_substitute_mapping(df_new, mapping_df, FIELD_MAPPINGS["赛卓-预测"])
        additional_sheets["赛卓-预测"] = df_new

        df_new = self.dataframes["赛卓-未交订单"]
        df_new, _ = apply_mapping_and_merge(df_new, mapping_df, FIELD_MAPPINGS["赛卓-未交订单"])
        df_new, _ = apply_extended_substitute_mapping(df_new, mapping_df, FIELD_MAPPINGS["赛卓-未交订单"])
        additional_sheets["赛卓-未交订单"] = df_new

        df_new = self.dataframes["赛卓-成品库存"]
        df_new, _ = apply_mapping_and_merge(df_new, mapping_df, FIELD_MAPPINGS["赛卓-成品库存"])
        df_new, _ = apply_extended_substitute_mapping(df_new, mapping_df, FIELD_MAPPINGS["赛卓-成品库存"])
        additional_sheets["赛卓-成品库存"] = df_new

        df_new = self.dataframes["赛卓-成品在制"]
        df_new, _ = apply_mapping_and_merge(df_new, mapping_df, FIELD_MAPPINGS["赛卓-成品在制"])
        df_new, _ = apply_extended_substitute_mapping(df_new, mapping_df, FIELD_MAPPINGS["赛卓-成品在制"])
        additional_sheets["赛卓-成品在制"] = df_new

        df_new = self.dataframes["赛卓-CP在制"]
        df_new, _ = apply_mapping_and_merge(df_new, mapping_df, FIELD_MAPPINGS["赛卓-CP在制"])
        df_new, _ = apply_extended_substitute_mapping(df_new, mapping_df, FIELD_MAPPINGS["赛卓-CP在制"])
        additional_sheets["赛卓-CP在制"] = df_new

        df_new = self.dataframes["赛卓-晶圆库存"]
        df_new, _ = apply_mapping_and_merge(df_new, mapping_df, FIELD_MAPPINGS["赛卓-晶圆库存"])
        df_new, _ = apply_extended_substitute_mapping(df_new, mapping_df, FIELD_MAPPINGS["赛卓-晶圆库存"])
        additional_sheets["赛卓-晶圆库存"] = df_new



            


        """
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
                st.write("3")
                st.write(additional_sheets)
                if sheet_name in self.dataframes:
                    self.dataframes[sheet_name] = df
                    st.write("4")
                    st.write(additional_sheets)
                else:
                    additional_sheets[sheet_name] = df
                    st.write("5")
                    st.write(additional_sheets)
            except Exception as e:
                st.error(f"❌ 替换 {sheet_name} 中的品名失败：{e}")
        """
        ## == 安全库存 ==
        safety_df = self.additional_sheets.get("赛卓-安全库存")
        if safety_df is not None and not safety_df.empty:
            main_plan_df, unmatched_safety = merge_safety_inventory(main_plan_df, safety_df)
            st.success("✅ 已合并安全库存数据")
        
        ## == 未交订单 ==
        unfulfilled_df = self.dataframes.get("赛卓-未交订单")
        if unfulfilled_df is not None and not unfulfilled_df.empty:
            main_plan_df, unmatched_unfulfilled = append_unfulfilled_summary_columns_by_date(main_plan_df, unfulfilled_df)
            st.success("✅ 已合并未交订单数据")


        ## == 预测 ==
        forecast_df = self.additional_sheets.get("赛卓-预测")
        if forecast_df is not None and not forecast_df.empty:
            main_plan_df, unmatched_forecast = append_forecast_to_summary(main_plan_df, forecast_df)
            st.success("✅ 已合并预测数据")

        ## == 成品库存 ==
        finished_df = self.dataframes.get("赛卓-成品库存")
        if finished_df is not None and not finished_df.empty:
            main_plan_df, unmatched_finished = merge_finished_inventory_with_warehouse_types(main_plan_df, finished_df, mapping_df)
            st.success("✅ 已合并成品库存数据")

        ## == 成品在制 ==
        product_in_progress_df = self.dataframes.get("赛卓-成品在制")
        if product_in_progress_df is not None and not product_in_progress_df.empty:
            main_plan_df, unmatched_in_progress = append_product_in_progress(main_plan_df, product_in_progress_df, mapping_df)
            st.success("✅ 已合并成品在制数据")

        # === 投单计划 ===
        forecast_months = init_monthly_fields(main_plan_df)

        # 成品&半成品实际投单
        df_order = self.dataframes.get("赛卓-下单明细", pd.DataFrame())
        main_plan_df = aggregate_actual_fg_orders(main_plan_df, df_order, forecast_months)
        main_plan_df = aggregate_actual_sfg_orders(main_plan_df, df_order, mapping_df, forecast_months)

        # 回货实际
        df_arrival = self.dataframes.get("赛卓-到货明细", pd.DataFrame())
        main_plan_df = aggregate_actual_arrivals(main_plan_df, df_arrival, forecast_months)

        # 销售数量&销售金额
        df_sales = self.dataframes.get("赛卓-销货明细", pd.DataFrame())
        main_plan_df = aggregate_sales_quantity_and_amount(main_plan_df, df_sales, forecast_months)

        # 成品投单计划
        main_plan_df = generate_monthly_fg_plan(main_plan_df, forecast_months)

        # 半成品投单计划
        main_plan_df = generate_monthly_semi_plan(main_plan_df, forecast_months, mapping_df)

        # 投单计划调整
        main_plan_df = generate_monthly_adjust_plan(main_plan_df)

        # 回货计划
        main_plan_df = generate_monthly_return_plan(main_plan_df)

        
        # 回货计划调整
        main_plan_df = generate_monthly_return_adjustment(main_plan_df)

        
        # 检查
        main_plan_df = drop_last_forecast_month_columns(main_plan_df, forecast_months)

        
        # === 写入 Excel 文件（主计划）===
        timestamp = datetime.now().strftime("%Y%m%d")
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            main_plan_df = clean_df(main_plan_df)
            main_plan_df.to_excel(writer, sheet_name="主计划", index=False, startrow=1)
            append_all_standardized_sheets(writer, uploaded_files, self.additional_sheets)
            
            # 替换上传文件 key 为标准名
            standardized_files = standardize_uploaded_keys(uploaded_files, RENAME_MAP)
            
            # 将 UploadedFile 读取为 DataFrame
            parsed_dataframes = {
                filename: pd.read_excel(file)  # 或提前 parse 完成的 DataFrame dict
                for filename, file in standardized_files.items()
            }
            
            pivot_tables = generate_monthly_pivots(parsed_dataframes, pivot_config)
            
            for sheet_name, df in pivot_tables.items():
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                
            
            # 写完后手动调整所有透视表 sheet 的列宽
            for sheet_name, df in pivot_tables.items():
                ws = writer.book[sheet_name]
                for col_cells in ws.columns:
                    max_length = 0
                    col_letter = col_cells[0].column_letter
                    for cell in col_cells:
                        try:
                            if cell.value:
                                max_length = max(max_length, len(str(cell.value)))
                        except:
                            pass
                    adjusted_width = max_length * 1.2 + 10
                    ws.column_dimensions[col_letter].width = min(adjusted_width, 50)
            
            ws = writer.book["主计划"]
            ws.cell(row=1, column=1, value=f"主计划生成时间：{timestamp}")
            

            legend_cell = ws.cell(row=1, column=3)
            legend_cell.value = (
                "🟥 < 0    "
                "🟨 < 安全库存    "
                "🟧 > 2 × 安全库存"
            )
            legend_cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")

            merge_safety_header(ws, main_plan_df)
            merge_unfulfilled_order_header(ws)
            merge_forecast_header(ws)
            merge_inventory_header(ws)
            merge_product_in_progress_header(ws)

            format_monthly_grouped_headers(ws)
            highlight_production_plan_cells(ws, main_plan_df)


            adjust_column_width(ws)


            # 设置字体加粗，行高也调高一点
            bold_font = Font(bold=True)
            ws.row_dimensions[2].height = 35
    
            # 遍历这一行所有已用到的列，对单元格字体加粗、居中、垂直居中
            max_col = ws.max_column
            for col_idx in range(1, max_col + 1):
                cell = ws.cell(row=2, column=col_idx)
                cell.font = bold_font
                # 垂直水平居中
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # 自动筛选
            last_col_letter = get_column_letter(ws.max_column)
            ws.auto_filter.ref = f"A2:{last_col_letter}2"
        
            # 冻结
            ws.freeze_panes = "D1"



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
