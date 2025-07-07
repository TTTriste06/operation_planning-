import pandas as pd
import streamlit as st
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font


def extract_wafer_with_grossdie_raw(main_plan_df: pd.DataFrame, df_grossdie: pd.DataFrame) -> pd.DataFrame:
    """
    直接将晶圆品名与 df_grossdie 的“规格”列做匹配，如果匹配上则取该行的“GROSS DIE”。

    参数：
        df_grossdie: 原始 grossdie 表（不可清洗）
        main_plan_df: 主计划表，包含“晶圆品名”

    返回：
        DataFrame: 包含“晶圆品名”和“单片数量”的 DataFrame
    """
    # 提取唯一晶圆品名
    wafer_names = (
        main_plan_df["晶圆品名"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .reset_index(drop=True)
    )
    df_result = pd.DataFrame({"晶圆品名": wafer_names})

    # 匹配逻辑：晶圆品名是否出现在 grossdie 的规格列中
    def match_grossdie(wafer_name):
        matched = df_grossdie[df_grossdie["规格"] == wafer_name]
        if not matched.empty:
            return matched.iloc[0]["GROSS DIE"]
        return None

    df_result["单片数量"] = df_result["晶圆品名"].apply(match_grossdie)

    return df_result


def append_inventory_columns(df_unique_wafer: pd.DataFrame, main_plan_df: pd.DataFrame) -> pd.DataFrame:
    """
    将每个晶圆品名在 main_plan_df 中对应的 InvWaf 与 InvPart 求和后，填入 df_unique_wafer。

    参数：
        df_unique_wafer: 包含唯一“晶圆品名”的 DataFrame
        main_plan_df: 包含完整数据（包含“晶圆品名”, "InvWaf", "InvPart"）

    返回：
        更新后的 df_unique_wafer，新增列：InvWaf, InvPart
    """
    # 只保留必要列并转换类型
    wafer_inventory = (
        main_plan_df[["晶圆品名", "InvWaf", "InvPart"]]
        .copy()
        .dropna(subset=["晶圆品名"])
    )
    wafer_inventory["晶圆品名"] = wafer_inventory["晶圆品名"].astype(str).str.strip()

    # 求和：以晶圆品名为索引聚合
    inventory_sum = wafer_inventory.groupby("晶圆品名", as_index=False)[["InvWaf", "InvPart"]].sum()

    # 合并回 df_unique_wafer
    df_unique_wafer = df_unique_wafer.copy()
    df_unique_wafer["晶圆品名"] = df_unique_wafer["晶圆品名"].astype(str).str.strip()

    df_merged = pd.merge(df_unique_wafer, inventory_sum, on="晶圆品名", how="left")

    return df_merged


def append_wafer_inventory_by_warehouse(df_unique_wafer: pd.DataFrame, wafer_inventory_df: pd.DataFrame) -> pd.DataFrame:
    """
    根据“晶圆品名”匹配 wafer_inventory_df 中的“WAFER品名”，
    并将其数量按“仓库名称”展开成多列，汇总填入 df_unique_wafer。
    """
    # 标准化字段
    wafer_inventory_df = wafer_inventory_df.copy()
    wafer_inventory_df["WAFER品名"] = wafer_inventory_df["WAFER品名"].astype(str).str.strip()
    wafer_inventory_df["仓库名称"] = wafer_inventory_df["仓库名称"].astype(str).str.strip()

    # 过滤出匹配的晶圆品名
    matched_inventory = wafer_inventory_df[
        wafer_inventory_df["WAFER品名"].isin(df_unique_wafer["晶圆品名"])
    ].copy()

    # 将“数量”确保是数字
    matched_inventory["数量"] = pd.to_numeric(matched_inventory["数量"], errors="coerce").fillna(0)

    # 透视表：按“晶圆品名”和“仓库名称”聚合数量
    pivot_inventory = matched_inventory.pivot_table(
        index="WAFER品名",
        columns="仓库名称",
        values="数量",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # 重命名 WAFER品名 → 晶圆品名，方便 merge
    pivot_inventory = pivot_inventory.rename(columns={"WAFER品名": "晶圆品名"})

    # 合并到原表
    df_result = pd.merge(df_unique_wafer, pivot_inventory, on="晶圆品名", how="left")

    return df_result

def merge_wafer_inventory_columns(ws: Worksheet, df: pd.DataFrame):
    """
    查找所有以“仓”结尾的列，在第一行合并并写入“晶圆库存”。

    参数：
        ws: openpyxl 的 Worksheet 对象（例如“主计划”sheet）
        df: 对应 DataFrame，用于定位列位置
    """
    # 1. 找出所有以“仓”结尾的列名
    inventory_cols = [col for col in df.columns if str(col).strip().endswith("仓")]
    if not inventory_cols:
        return  # 没有匹配到“仓”列，无需处理

    # 2. 获取这些列在 DataFrame 中的索引位置（从0开始）转为 Excel 列号（从1开始）
    start_col_idx = df.columns.get_loc(inventory_cols[0]) + 1
    end_col_idx = df.columns.get_loc(inventory_cols[-1]) + 1

    # 3. 获取列字母（如 E, F）
    start_letter = get_column_letter(start_col_idx)
    end_letter = get_column_letter(end_col_idx)

    # 4. 合并单元格并写入标题“晶圆库存”
    title_cell = ws.cell(row=1, column=start_col_idx, value="晶圆库存")
    ws.merge_cells(start_row=1, start_column=start_col_idx, end_row=1, end_column=end_col_idx)
