import pandas as pd
import streamlit as st

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
