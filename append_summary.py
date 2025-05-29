import pandas as pd
from typing import Tuple
from openpyxl.styles import PatternFill

def standardize(val):
    if val is None:
        return ""
    return str(val).strip().replace("\u3000", " ").strip('\'"“”‘’')

def append_forecast_unmatched_to_summary_by_keys(summary_df: pd.DataFrame, forecast_df: pd.DataFrame) -> Tuple[pd.DataFrame, list]:
    """
    添加预测中未匹配的记录到汇总，并返回新增行索引。
    """
    forecast_cols = [col for col in forecast_df.columns if "预测" in col]
    needed_cols = ["产品型号", "生产料号"] + forecast_cols
    forecast_subset = forecast_df[needed_cols].copy()

    unmatched = forecast_subset[~forecast_subset["生产料号"].isin(summary_df["品名"])].copy()

    unmatched["规格"] = unmatched["产品型号"]
    unmatched["品名"] = unmatched["生产料号"]
    unmatched["晶圆品名"] = ""

    summary_cols = summary_df.columns
    new_rows = unmatched[["晶圆品名", "规格", "品名"] + forecast_cols if forecast_cols else []]

    for col in summary_cols:
        if col not in new_rows.columns:
            new_rows[col] = ""

    new_rows = new_rows[summary_cols]
    
    summary_df = pd.concat([summary_df, new_rows], ignore_index=True)

    return summary_df
