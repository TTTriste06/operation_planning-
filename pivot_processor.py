import re
import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill

from config import FILE_KEYWORDS, OUTPUT_FILENAME_PREFIX, FIELD_MAPPINGS, pivot_config, RENAME_MAP
from excel_utils import adjust_column_width, highlight_replaced_names_in_main_sheet, reorder_main_plan_by_unfulfilled_sheet
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

        # === 标准化新旧料号表 ===
        self.additional_sheets = additional_sheets
        mapping_df = self.additional_sheets.get("赛卓-新旧料号")
        if mapping_df is None or mapping_df.empty:
            raise ValueError("❌ 缺少新旧料号映射表，无法进行品名替换。")

        # 创建新的 mapping_semi：仅保留“半成品”字段非空的行
        mapping_semi = mapping_df[~mapping_df["半成品"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        st.write(mapping_semi)
        
        # 去除“品名”为空的行
        mapping_new = mapping_df[~mapping_df["新品名"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_new = mapping_new[~mapping_new["旧品名"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        st.write(mapping_new)
        
        # 去除“替代品名1”为空的行，并保留指定字段
        rename_dict = {
            "新晶圆品名": "NewWaferName",
            "新规格": "NewSpec",
            "新品名": "NewName",
            "封装厂": "PackagingFactory",
            "PC": "PC",
            "半成品": "SemiProduct",
            "备注": "Remark",
            "替代晶圆1": "AltWafer1",
            "替代规格1": "AltSpec1",
            "替代品名1": "AltName1"
        }

        
        mapping_sub1 = mapping_df[
            ["新晶圆品名", "新规格", "新品名", "封装厂", "PC", "半成品", "备注", "替代晶圆1", "替代规格1", "替代品名1"]
        ]
        mapping_sub1 = mapping_sub1[~mapping_df["替代品名1"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_sub1.columns = [
            "新晶圆品名", "新规格", "新品名", "封装厂", "PC", "半成品", "备注",
            "替代晶圆", "替代规格", "替代品名"
        ]


        mapping_sub2 = mapping_df[
            ["新晶圆品名", "新规格", "新品名", "封装厂", "PC", "半成品", "备注", "替代晶圆2", "替代规格2", "替代品名2"]
        ]
        mapping_sub2 = mapping_sub2[~mapping_df["替代品名2"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_sub2.columns = [
            "新晶圆品名", "新规格", "新品名", "封装厂", "PC", "半成品", "备注",
            "替代晶圆", "替代规格", "替代品名"
        ]

        mapping_sub3 = mapping_df[
            ["新晶圆品名", "新规格", "新品名", "封装厂", "PC", "半成品", "备注", "替代晶圆3", "替代规格3", "替代品名3"]
        ]
        mapping_sub3 = mapping_sub3[~mapping_df["替代品名3"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_sub3.columns = [
            "新晶圆品名", "新规格", "新品名", "封装厂", "PC", "半成品", "备注",
            "替代晶圆", "替代规格", "替代品名"
        ]
        

        mapping_sub4 = mapping_df[
            ["新晶圆品名", "新规格", "新品名", "封装厂", "PC", "半成品", "备注", "替代晶圆4", "替代规格4", "替代品名4"]
        ]
        mapping_sub4 = mapping_sub4[~mapping_df["替代品名4"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_sub4.columns = [
            "新晶圆品名", "新规格", "新品名", "封装厂", "PC", "半成品", "备注",
            "替代晶圆", "替代规格", "替代品名"
        ]

       

        
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
        all_names = replace_all_names_with_mapping(all_names, mapping_new, mapping_new)


        main_plan_df = main_plan_df.reindex(index=range(len(all_names)))
        if not all_names.empty:
            main_plan_df["品名"] = all_names.values

        ## == 规格和晶圆 ==
        main_plan_df = fill_spec_and_wafer_info(
            main_plan_df,
            self.dataframes,
            self.additional_sheets,
            mapping_semi, 
            FIELD_MAPPINGS
        )

        ## == 封装厂，封装形式和PC ==
        main_plan_df = fill_packaging_info(
            main_plan_df,
            dataframes=self.dataframes,
            additional_sheets=self.additional_sheets
        )

        
        ## == 替换新旧料号、替代料号 ==
        all_replaced_names = set()  # 用 set 累计替换的新品名
        df_new = self.additional_sheets["赛卓-安全库存"]
        df_new, replaced_main = apply_mapping_and_merge(df_new, mapping_new, FIELD_MAPPINGS["赛卓-安全库存"])
        df_new, replaced_sub1 = apply_extended_substitute_mapping(df_new, mapping_sub1, FIELD_MAPPINGS["赛卓-安全库存"])
        df_new, replaced_sub2 = apply_extended_substitute_mapping(df_new, mapping_sub2, FIELD_MAPPINGS["赛卓-安全库存"])
        df_new, replaced_sub3 = apply_extended_substitute_mapping(df_new, mapping_sub3, FIELD_MAPPINGS["赛卓-安全库存"])
        df_new, replaced_sub4 = apply_extended_substitute_mapping(df_new, mapping_sub4, FIELD_MAPPINGS["赛卓-安全库存"])
        self.additional_sheets["赛卓-安全库存"] = df_new
        all_replaced_names.update(replaced_main)
        all_replaced_names.update(replaced_sub1)
        all_replaced_names.update(replaced_sub2)
        all_replaced_names.update(replaced_sub3)
        all_replaced_names.update(replaced_sub4)
        
        df_new = self.additional_sheets["赛卓-预测"]
        df_new, replaced_main = apply_mapping_and_merge(df_new, mapping_new, FIELD_MAPPINGS["赛卓-预测"])
        df_new, replaced_sub1 = apply_extended_substitute_mapping(df_new, mapping_sub1, FIELD_MAPPINGS["赛卓-预测"])
        df_new, replaced_sub2 = apply_extended_substitute_mapping(df_new, mapping_sub2, FIELD_MAPPINGS["赛卓-预测"])
        df_new, replaced_sub3 = apply_extended_substitute_mapping(df_new, mapping_sub3, FIELD_MAPPINGS["赛卓-预测"])
        df_new, replaced_sub4 = apply_extended_substitute_mapping(df_new, mapping_sub4, FIELD_MAPPINGS["赛卓-预测"])
        self.additional_sheets["赛卓-预测"] = df_new
        all_replaced_names.update(replaced_main)
        all_replaced_names.update(replaced_sub1)
        all_replaced_names.update(replaced_sub2)
        all_replaced_names.update(replaced_sub3)
        all_replaced_names.update(replaced_sub4)

        df_new = self.dataframes["赛卓-未交订单"]
        df_new, replaced_main = apply_mapping_and_merge(df_new, mapping_new, FIELD_MAPPINGS["赛卓-未交订单"])
        df_new, replaced_sub1 = apply_extended_substitute_mapping(df_new, mapping_sub1, FIELD_MAPPINGS["赛卓-未交订单"])
        df_new, replaced_sub2 = apply_extended_substitute_mapping(df_new, mapping_sub2, FIELD_MAPPINGS["赛卓-未交订单"])
        df_new, replaced_sub3 = apply_extended_substitute_mapping(df_new, mapping_sub3, FIELD_MAPPINGS["赛卓-未交订单"])
        df_new, replaced_sub4 = apply_extended_substitute_mapping(df_new, mapping_sub4, FIELD_MAPPINGS["赛卓-未交订单"])
        self.dataframes["赛卓-未交订单"] = df_new
        all_replaced_names.update(replaced_main)
        all_replaced_names.update(replaced_sub1)
        all_replaced_names.update(replaced_sub2)
        all_replaced_names.update(replaced_sub3)
        all_replaced_names.update(replaced_sub4)

        df_new = self.dataframes["赛卓-成品库存"]
        df_new, replaced_main = apply_mapping_and_merge(df_new, mapping_new, FIELD_MAPPINGS["赛卓-成品库存"])
        df_new, replaced_sub1 = apply_extended_substitute_mapping(df_new, mapping_sub1, FIELD_MAPPINGS["赛卓-成品库存"])
        df_new, replaced_sub2 = apply_extended_substitute_mapping(df_new, mapping_sub2, FIELD_MAPPINGS["赛卓-成品库存"])
        df_new, replaced_sub3 = apply_extended_substitute_mapping(df_new, mapping_sub3, FIELD_MAPPINGS["赛卓-成品库存"])
        df_new, replaced_sub4 = apply_extended_substitute_mapping(df_new, mapping_sub4, FIELD_MAPPINGS["赛卓-成品库存"])
        self.dataframes["赛卓-成品库存"] = df_new
        all_replaced_names.update(replaced_main)
        all_replaced_names.update(replaced_sub1)
        all_replaced_names.update(replaced_sub2)
        all_replaced_names.update(replaced_sub3)
        all_replaced_names.update(replaced_sub4)

        df_new = self.dataframes["赛卓-成品在制"]
        df_new, replaced_main = apply_mapping_and_merge(df_new, mapping_new, FIELD_MAPPINGS["赛卓-成品在制"])
        df_new, replaced_sub1 = apply_extended_substitute_mapping(df_new, mapping_sub1, FIELD_MAPPINGS["赛卓-成品在制"])
        df_new, replaced_sub2 = apply_extended_substitute_mapping(df_new, mapping_sub2, FIELD_MAPPINGS["赛卓-成品在制"])
        df_new, replaced_sub3 = apply_extended_substitute_mapping(df_new, mapping_sub3, FIELD_MAPPINGS["赛卓-成品在制"])
        df_new, replaced_sub4 = apply_extended_substitute_mapping(df_new, mapping_sub4, FIELD_MAPPINGS["赛卓-成品在制"])
        self.dataframes["赛卓-成品在制"] = df_new
        all_replaced_names.update(replaced_main)
        all_replaced_names.update(replaced_sub1)
        all_replaced_names.update(replaced_sub2)
        all_replaced_names.update(replaced_sub3)
        all_replaced_names.update(replaced_sub4)

        df_new = self.dataframes["赛卓-CP在制"]
        df_new, replaced_main = apply_mapping_and_merge(df_new, mapping_new, FIELD_MAPPINGS["赛卓-CP在制"])
        df_new, replaced_sub1 = apply_extended_substitute_mapping(df_new, mapping_sub1, FIELD_MAPPINGS["赛卓-CP在制"])
        df_new, replaced_sub2 = apply_extended_substitute_mapping(df_new, mapping_sub2, FIELD_MAPPINGS["赛卓-CP在制"])
        df_new, replaced_sub3 = apply_extended_substitute_mapping(df_new, mapping_sub3, FIELD_MAPPINGS["赛卓-CP在制"])
        df_new, replaced_sub4 = apply_extended_substitute_mapping(df_new, mapping_sub4, FIELD_MAPPINGS["赛卓-CP在制"])
        self.dataframes["赛卓-CP在制"] = df_new
        all_replaced_names.update(replaced_main)
        all_replaced_names.update(replaced_sub1)
        all_replaced_names.update(replaced_sub2)
        all_replaced_names.update(replaced_sub3)
        all_replaced_names.update(replaced_sub4)

        df_new = self.dataframes["赛卓-晶圆库存"]
        df_new, replaced_main = apply_mapping_and_merge(df_new, mapping_new, FIELD_MAPPINGS["赛卓-晶圆库存"])
        df_new, replaced_sub1 = apply_extended_substitute_mapping(df_new, mapping_sub1, FIELD_MAPPINGS["赛卓-晶圆库存"])
        df_new, replaced_sub2 = apply_extended_substitute_mapping(df_new, mapping_sub2, FIELD_MAPPINGS["赛卓-晶圆库存"])
        df_new, replaced_sub3 = apply_extended_substitute_mapping(df_new, mapping_sub3, FIELD_MAPPINGS["赛卓-晶圆库存"])
        df_new, replaced_sub4 = apply_extended_substitute_mapping(df_new, mapping_sub4, FIELD_MAPPINGS["赛卓-晶圆库存"])
        self.dataframes["赛卓-晶圆库存"] = df_new
        all_replaced_names.update(replaced_main)
        all_replaced_names.update(replaced_sub1)
        all_replaced_names.update(replaced_sub2)
        all_replaced_names.update(replaced_sub3)
        all_replaced_names.update(replaced_sub4)

        all_replaced_names = sorted(all_replaced_names)


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
            main_plan_df, unmatched_finished = merge_finished_inventory_with_warehouse_types(main_plan_df, finished_df, mapping_semi)
            st.success("✅ 已合并成品库存数据")

        ## == 成品在制 ==
        product_in_progress_df = self.dataframes.get("赛卓-成品在制")
        if product_in_progress_df is not None and not product_in_progress_df.empty:
            main_plan_df, unmatched_in_progress = append_product_in_progress(main_plan_df, product_in_progress_df, mapping_semi)
            st.success("✅ 已合并成品在制数据")

        # === 投单计划 ===
        forecast_months = init_monthly_fields(main_plan_df)

        # 成品&半成品实际投单
        df_order = self.dataframes.get("赛卓-下单明细", pd.DataFrame())
        main_plan_df = aggregate_actual_fg_orders(main_plan_df, df_order, forecast_months)
        main_plan_df = aggregate_actual_sfg_orders(main_plan_df, df_order, mapping_semi, forecast_months)

        # 回货实际
        df_arrival = self.dataframes.get("赛卓-到货明细", pd.DataFrame())
        main_plan_df = aggregate_actual_arrivals(main_plan_df, df_arrival, forecast_months)

        # 销售数量&销售金额
        df_sales = self.dataframes.get("赛卓-销货明细", pd.DataFrame())
        main_plan_df = aggregate_sales_quantity_and_amount(main_plan_df, df_sales, forecast_months)

        # 成品投单计划
        main_plan_df = generate_monthly_fg_plan(main_plan_df, forecast_months)

        # 半成品投单计划
        main_plan_df = generate_monthly_semi_plan(main_plan_df, forecast_months, mapping_semi)

        # 投单计划调整
        main_plan_df = generate_monthly_adjust_plan(main_plan_df)

        # 回货计划
        main_plan_df = generate_monthly_return_plan(main_plan_df)

        
        # 回货计划调整
        main_plan_df = generate_monthly_return_adjustment(main_plan_df)

        
        # 检查
        main_plan_df = reorder_main_plan_by_unfulfilled_sheet(main_plan_df, unfulfilled_df)
        main_plan_df = drop_last_forecast_month_columns(main_plan_df, forecast_months)
        
        
        # === 写入 Excel 文件（主计划）===
        timestamp = datetime.now().strftime("%Y%m%d")
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            main_plan_df = clean_df(main_plan_df)
            main_plan_df.to_excel(writer, sheet_name="主计划", index=False, startrow=1)
            append_all_standardized_sheets(writer, uploaded_files, self.additional_sheets)
            
            # 透视表
            standardized_files = standardize_uploaded_keys(uploaded_files, RENAME_MAP)
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

            #写入主计划
            ws = writer.book["主计划"]
            ws.cell(row=1, column=1, value=f"主计划生成时间：{timestamp}")
            
            legend_cell = ws.cell(row=1, column=3)
            legend_cell.value = (
                "Red < 0    "
                "Yellow < 安全库存    "
                "Orange > 2 × 安全库存"
            )
            legend_cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            fill = PatternFill(start_color="FFFFC0CB", end_color="FFFFC0CB", fill_type="solid")
            legend_cell.fill = fill


            merge_safety_header(ws, main_plan_df)
            merge_unfulfilled_order_header(ws)
            merge_forecast_header(ws)
            merge_inventory_header(ws)
            merge_product_in_progress_header(ws)

            format_monthly_grouped_headers(ws)
            highlight_production_plan_cells(ws, main_plan_df)
            highlight_replaced_names_in_main_sheet(ws, all_replaced_names)


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
            ws.freeze_panes = "D3"

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

