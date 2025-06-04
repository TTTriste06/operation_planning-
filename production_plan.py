import re
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill, Font
from datetime import datetime
from collections import defaultdict
from openpyxl.styles import numbers

def init_monthly_fields(main_plan_df: pd.DataFrame) -> list[int]:
    """
    自动识别主计划中预测字段的月份，添加 HEADER_TEMPLATE 中的所有月度字段列。
    初始化为 ""。
    
    返回：
    - forecast_months: 所有识别出的月份列表（升序）
    """
    HEADER_TEMPLATE = [
        "销售数量", "销售金额", "成品投单计划", "半成品投单计划", "投单计划调整",
        "成品可行投单", "半成品可行投单", "成品实际投单", "半成品实际投单",
        "回货计划", "回货计划调整", "PC回货计划", "回货实际"
    ]

    month_pattern = re.compile(r"^(\d{1,2})月预测$")
    forecast_months = sorted({
        int(match.group(1)) for col in main_plan_df.columns
        if isinstance(col, str) and (match := month_pattern.match(col.strip()))
    })

    if not forecast_months:
        return []

    start_month = datetime.today().month
    end_month = max(forecast_months) - 1

    for m in range(start_month, end_month + 1):
        for header in HEADER_TEMPLATE:
            col = f"{m}月{header}"
            if col not in main_plan_df.columns:
                main_plan_df[col] = ""

    return forecast_months

def safe_col(df: pd.DataFrame, col: str) -> pd.Series:
    """确保列为数字，若不存在则返回 0"""
    return pd.to_numeric(df[col], errors="coerce").fillna(0) if col in df.columns else pd.Series(0, index=df.index)

def generate_monthly_fg_plan(main_plan_df: pd.DataFrame, forecast_months: list[int]) -> pd.DataFrame:
    """
    生成每月“成品投单计划”列，规则：
    - 第一个月：InvPart + max(预测, 未交) + max(预测, 未交)（下月） - 成品仓 - 成品在制
    - 后续月份：max(预测, 未交)（下月） + （上月投单 - 上月实际投单）
    
    参数：
    - main_plan_df: 主计划表（含所有字段）
    - forecast_months: 所有月份的列表（int 类型，如 [6, 7, 8, ...]）

    返回：
    - main_plan_df: 添加了成品投单计划字段的 DataFrame
    """

    df_plan = pd.DataFrame(index=main_plan_df.index)

    for idx, month in enumerate(forecast_months[:-1]):  # 最后一个月不生成
        this_month = f"{month}月"
        next_month = f"{forecast_months[idx + 1]}月"
        prev_month = f"{forecast_months[idx - 1]}月" if idx > 0 else None

        # 构造字段名
        col_forecast_this = f"{month}月预测"
        col_order_this = f"未交订单数量_2025-{month:02d}"
        col_forecast_next = f"{forecast_months[idx + 1]}月预测"
        col_order_next = f"未交订单数量_2025-{forecast_months[idx + 1]:02d}"
        col_target = f"{month}月成品投单计划"
        col_actual_prod = f"{this_month}_成品实际投单"
        col_target_prev = f"{prev_month}_成品投单计划" if prev_month else None

        if idx == 0:
            df_plan[col_target] = (
                safe_col(main_plan_df, "InvPart") +
                pd.DataFrame({
                    "f": safe_col(main_plan_df, col_forecast_this),
                    "o": safe_col(main_plan_df, col_order_this)
                }).max(axis=1) +
                pd.DataFrame({
                    "f": safe_col(main_plan_df, col_forecast_next),
                    "o": safe_col(main_plan_df, col_order_next)
                }).max(axis=1) -
                safe_col(main_plan_df, "成品仓") -
                safe_col(main_plan_df, "成品在制")
            )
        else:
            df_plan[col_target] = (
                pd.DataFrame({
                    "f": safe_col(main_plan_df, col_forecast_next),
                    "o": safe_col(main_plan_df, col_order_next)
                }).max(axis=1) +
                (safe_col(df_plan, col_target_prev) - safe_col(main_plan_df, col_actual_prod))
            )

    plan_cols_in_summary = [col for col in summary_preview.columns if "成品投单计划" in col and "半成品" not in col]
    
    # 回填到主计划中
    if len(plan_cols_in_summary) != df_plan.shape[1]:
        st.error(f"❌ 写入失败：df_plan 有 {df_plan.shape[1]} 列，summary 中有 {len(plan_cols_in_summary)} 个 '成品投单计划' 列")
    else:
        # ✅ 将 df_plan 的列按顺序填入 summary_preview
        for i, col in enumerate(plan_cols_in_summary):
            main_plan_df[col] = df_plan.iloc[:, i]

    return main_plan_df














def apply_monthly_grouped_headers(ws):
    """
    自动合并主计划中按“月_字段”格式的列，如“6月销售数量”，将每月字段统一合并为“6月”标题。
    """
    header_row = [cell.value for cell in ws[2]]  # 第2行是字段名
    pattern = re.compile(r"^(\d{1,2})月(.*)$")

    # group: {month -> [col_idx, ...]}
    monthly_groups = defaultdict(list)

    for i, col in enumerate(header_row):
        if not isinstance(col, str):
            continue
        match = pattern.match(col.strip())
        if match:
            month = int(match.group(1))
            monthly_groups[month].append(i + 1)  # openpyxl 列号从1开始

    # 遍历每个识别到的月份
    for month, cols in sorted(monthly_groups.items()):
        start_col = min(cols)
        end_col = max(cols)

        if end_col >= start_col:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)

        cell = ws.cell(row=1, column=start_col)
        cell.value = f"{month}月"
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FFFF00")



class MonthlyPlanGenerator:
    HEADER_TEMPLATE = [
        "销售数量", "销售金额", "成品投单计划", "半成品投单计划", "投单计划调整",
        "成品可行投单", "半成品可行投单", "成品实际投单", "半成品实际投单",
        "回货计划", "回货计划调整", "PC回货计划", "回货实际"
    ]

    def __init__(self, summary_df: pd.DataFrame):
        self.summary_df = summary_df
        self.forecast_months = self.extract_forecast_months()
        self.start_month = datetime.today().month
        self.end_month = max(self.forecast_months)

    def extract_forecast_months(self) -> list[int]:
        pattern = re.compile(r"(\d{1,2})月预测")
        months = [
            int(match.group(1)) for col in self.summary_df.columns
            if (match := pattern.match(str(col)))
        ]
        return sorted(set(months))

    def init_monthly_columns(self):
        for m in range(self.start_month, self.end_month + 1):
            for header in self.HEADER_TEMPLATE:
                new_col = f"{m}月{header}"
                if new_col not in self.summary_df.columns:
                    self.summary_df[new_col] = ""

    def safe_col(self, col: str) -> pd.Series:
        return pd.to_numeric(self.summary_df[col], errors="coerce").fillna(0) if col in self.summary_df.columns else pd.Series(0, index=self.summary_df.index)

    def compute_product_plan(self):
        df_plan = pd.DataFrame(index=self.summary_df.index)
        forecast = self.forecast_months

        for idx, month in enumerate(forecast[:-1]):
            next_month = forecast[idx + 1]
            col = f"{month}月成品投单计划"
            col_next = f"{next_month}月预测"
            col_order = f"未交订单数量_2025-{month}"
            col_order_next = f"未交订单数量_2025-{next_month}"
            col_actual = f"{month}月成品实际投单"

            if idx == 0:
                df_plan[col] = (
                    self.safe_col("InvPart") +
                    pd.DataFrame({
                        "f": self.safe_col(f"{month}月预测"),
                        "o": self.safe_col(col_order)
                    }).max(axis=1) +
                    pd.DataFrame({
                        "f": self.safe_col(col_next),
                        "o": self.safe_col(col_order_next)
                    }).max(axis=1) -
                    self.safe_col("数量_成品仓") -
                    self.safe_col("成品在制")
                )
            else:
                col_prev = f"{forecast[idx - 1]}月成品投单计划"
                df_plan[col] = (
                    pd.DataFrame({
                        "f": self.safe_col(col_next),
                        "o": self.safe_col(col_order_next)
                    }).max(axis=1) +
                    (self.safe_col(col_prev) - self.safe_col(col_actual))
                )

        # 写入
        for col in df_plan.columns:
            self.summary_df[col] = df_plan[col]

    def compute_semi_plan(self):
        df_semi = pd.DataFrame(index=self.summary_df.index)
        forecast = self.forecast_months

        for idx, month in enumerate(forecast[:-1]):
            col_fp = f"{month}月成品投单计划"
            col_sp = f"{month}月半成品投单计划"
            if idx == 0:
                df_semi[col_sp] = self.safe_col(col_fp) - self.safe_col("半成品在制")
            else:
                df_semi[col_sp] = 0  # 后面写公式

        for col in df_semi.columns:
            self.summary_df[col] = df_semi[col]

    def write_formulas_to_excel(self, ws, header_base: str, start_row: int = 3):
        """ 针对如‘半成品投单计划’这类字段补公式 """
        cols = [col for col in self.summary_df.columns if header_base in col]
        for i, col in enumerate(cols):
            if i == 0:
                continue  # 第一个月不写公式
            col_idx = self.summary_df.columns.get_loc(col) + 1
            prev_letter = get_column_letter(col_idx - 1)
            back13 = get_column_letter(col_idx - 13)
            back8 = get_column_letter(col_idx - 8)
            for row in range(start_row, len(self.summary_df) + start_row):
                ws.cell(row=row, column=col_idx).value = f"={prev_letter}{row}+({back13}{row}-{back8}{row})"

    def merge_monthly_headers(self, ws):
        header_row = list(self.summary_df.columns)
        for base in self.HEADER_TEMPLATE:
            cols = [
                idx for idx, col in enumerate(header_row, start=1)
                if isinstance(col, str) and col.endswith(base)
            ]
            if cols:
                start_col = min(cols)
                end_col = max(cols)
                ws.merge_cells(f"{get_column_letter(start_col)}1:{get_column_letter(end_col)}1")
                cell = ws.cell(row=1, column=start_col)
                cell.value = base
                cell.alignment = Alignment(horizontal="center", vertical="center")


class MonthlyFieldAggregator:
    def __init__(self, summary_df: pd.DataFrame, forecast_months: list[int]):
        self.summary_df = summary_df
        self.forecast_months = forecast_months

    def aggregate_sales(self, df_sales: pd.DataFrame):
        df_sales = df_sales.copy()
        df_sales["品名"] = df_sales["品名"].astype(str).str.strip()
        df_sales["销售月份"] = pd.to_datetime(df_sales["交易日期"], errors="coerce").dt.month
        valid_names = set(self.summary_df["品名"].astype(str))
        df_sales = df_sales[df_sales["品名"].isin(valid_names)]

        qty_df = pd.DataFrame({"品名": self.summary_df["品名"]})
        amt_df = pd.DataFrame({"品名": self.summary_df["品名"]})
        for m in self.forecast_months:
            qty_df[f"{m}月销售数量"] = 0
            amt_df[f"{m}月销售金额"] = 0

        for _, row in df_sales.iterrows():
            part, month = row["品名"], row["销售月份"]
            if month in self.forecast_months:
                qty_col = f"{month}月销售数量"
                amt_col = f"{month}月销售金额"
                idx = qty_df[qty_df["品名"] == part].index
                if not idx.empty:
                    qty_df.at[idx[0], qty_col] += row["数量"]
                    amt_df.at[idx[0], amt_col] += row["原币金额"]

        for col in qty_df.columns[1:]:
            self.summary_df[col] = qty_df[col]
        for col in amt_df.columns[1:]:
            self.summary_df[col] = amt_df[col]

    def aggregate_arrival(self, df_arrival: pd.DataFrame):
        df_arrival = df_arrival.copy()
        df_arrival["品名"] = df_arrival["品名"].astype(str).str.strip()
        df_arrival["到货月份"] = pd.to_datetime(df_arrival["到货日期"], errors="coerce").dt.month
        valid_names = set(self.summary_df["品名"].astype(str))
        df_arrival = df_arrival[df_arrival["品名"].isin(valid_names)]

        result = pd.DataFrame({"品名": self.summary_df["品名"]})
        for m in self.forecast_months:
            result[f"{m}月回货实际"] = 0

        for _, row in df_arrival.iterrows():
            part, month = row["品名"], row["到货月份"]
            if month in self.forecast_months:
                col = f"{month}月回货实际"
                idx = result[result["品名"] == part].index
                if not idx.empty:
                    result.at[idx[0], col] += row["允收数量"]

        for col in result.columns[1:]:
            self.summary_df[col] = result[col]

    def aggregate_orders(self, df_order: pd.DataFrame):
        df_order = df_order.copy()
        df_order["品名"] = df_order["回货明细_回货品名"].astype(str).str.strip()
        df_order["下单月份"] = pd.to_datetime(df_order["下单日期"], errors="coerce").dt.month
        valid_names = set(self.summary_df["品名"].astype(str))
        df_order = df_order[df_order["品名"].isin(valid_names)]

        result = pd.DataFrame({"品名": self.summary_df["品名"]})
        for m in self.forecast_months:
            result[f"{m}月成品实际投单"] = 0

        for _, row in df_order.iterrows():
            part, month = row["品名"], row["下单月份"]
            if month in self.forecast_months:
                col = f"{month}月成品实际投单"
                idx = result[result["品名"] == part].index
                if not idx.empty:
                    result.at[idx[0], col] += row["回货明细_回货数量"]

        for col in result.columns[1:]:
            self.summary_df[col] = result[col]
