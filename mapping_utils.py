import pandas as pd
import streamlit as st

def clean_mapping_headers(mapping_df):
    """
    将新旧料号表的列名重命名为标准字段，按列数自动对齐；若列数超限则报错。
    """
    required_headers = [
        "旧规格", "旧品名", "旧晶圆品名",
        "新规格", "新品名", "新晶圆品名",
        "封装厂", "PC", "半成品", "备注",
        "替代规格1", "替代品名1", "替代晶圆1",
        "替代规格2", "替代品名2", "替代晶圆2",
        "替代规格3", "替代品名3", "替代晶圆3",
        "替代规格4", "替代品名4", "替代晶圆4"
    ]

    if mapping_df.shape[1] > len(required_headers):
        raise ValueError(f"❌ 新旧料号列数超出预期：共 {mapping_df.shape[1]} 列，最多支持 {len(required_headers)} 列")

    # ✅ 重命名当前列
    mapping_df.columns = required_headers[:mapping_df.shape[1]]

    # ✅ 仅保留这些列
    return mapping_df[required_headers[:mapping_df.shape[1]]]

def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True):
    """
    按品名字段替换主料号（新旧料号映射）
    """
    name_col = field_map["品名"]
    df[name_col] = df[name_col].astype(str).str.strip()

    mapping_df["旧品名"] = mapping_df["旧品名"].astype(str).str.strip()
    mapping_df["新品名"] = mapping_df["新品名"].astype(str).str.strip()

    df = df.copy()
    merged = df.merge(mapping_df[["旧品名", "新品名"]], how="left", left_on=name_col, right_on="旧品名")

    # ✅ 只替换有值的新品名
    mask = merged["新品名"].notna() & (merged["新品名"].str.strip() != "")
    merged.loc[mask, name_col] = merged.loc[mask, "新品名"]

    if verbose:
        st.success(f"✅ 新旧料号替换完成，共替换: {mask.sum()} 行")

    mapped_keys = set(merged.loc[mask, name_col])
    return merged.drop(columns=["旧品名", "新品名"], errors="ignore"), mapped_keys

def apply_extended_substitute_mapping(df, mapping_df, field_map, verbose=True):
    """
    替代料号品名替换（仅品名字段替换，无聚合合并）
    """
    name_col = field_map["品名"]
    df = df.copy()
    df[name_col] = df[name_col].astype(str).str.strip().str.replace("\n", "").str.replace("\r", "")

    matched_keys = set()

    for i in range(1, 5):
        sub_col = f"替代品名{i}"
        if sub_col not in mapping_df.columns or "新品名" not in mapping_df.columns:
            continue

        sub_df = mapping_df[[sub_col, "新品名"]].dropna()
        sub_df[sub_col] = sub_df[sub_col].astype(str).str.strip()
        sub_df["新品名"] = sub_df["新品名"].astype(str).str.strip()

        valid_sub = sub_df[
            (sub_df[sub_col] != "") &
            (sub_df["新品名"] != "")
        ]

        for _, row in valid_sub.iterrows():
            mask = df[name_col] == row[sub_col]
            if mask.any():
                df.loc[mask, name_col] = row["新品名"]
                matched_keys.update([row["新品名"]])

    if verbose:
        st.success(f"✅ 替代品名替换完成，共替换: {len(matched_keys)} 种")

    return df, matched_keys
