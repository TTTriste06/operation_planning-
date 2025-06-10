import pandas as pd
import streamlit as st

def standardize_uploaded_keys(uploaded_files: dict, rename_map: dict) -> dict:
    standardized = {}

    for filename, file_obj in uploaded_files.items():
        matched = False
        for key, standard_name in rename_map.items():
            if key in filename:
                standardized[standard_name] = file_obj
                matched = True
                break
        if not matched:
            standardized[filename] = file_obj  # 保留未匹配的
    return standardized

import pandas as pd

def generate_monthly_pivots(dataframes: dict, pivot_config: dict) -> dict:
    """
    为多个 DataFrame 根据配置生成透视表。

    支持:
    - 固定 index 字段：必须全部存在；
    - 日期字段格式化为 %Y-%m；
    - 多 value 列聚合；
    - 缺失字段自动填空。

    返回:
        dict[sheet_name -> pd.DataFrame]
    """
    pivot_tables = {}

    for filename, df in dataframes.items():
        if filename not in pivot_config:
            print(f"⚠️ 未找到 {filename} 的透视配置，跳过")
            continue

        config = pivot_config[filename]
        index = config.get("index", [])
        columns = config["columns"]
        values = config["values"]
        aggfunc = config.get("aggfunc", "sum")
        date_format = config.get("date_format")

        df = df.copy()

        # 日期字段格式化
        if date_format:
            try:
                df[columns] = pd.to_datetime(df[columns], errors='coerce')
                df = df.dropna(subset=[columns])
                df[columns] = df[columns].dt.to_period("M").astype(str)
            except Exception as e:
                print(f"❌ 日期字段格式化失败 [{filename}]：{e}")
                continue

        # 检查 index 字段是否都存在
        if not all(col in df.columns for col in index):
            print(f"⚠️ {filename} 缺少部分 index 字段，跳过")
            continue

        # 填空，确保 index 字段不会因为 NaN 而被排除
        for col in index:
            df[col] = df[col].astype(str).fillna("").replace("nan", "").str.strip()

        try:
            pivot = pd.pivot_table(
                df,
                index=index,
                columns=columns,
                values=values,
                aggfunc=aggfunc,
                fill_value=0,
                dropna=False
            )

            if isinstance(pivot.columns, pd.MultiIndex):
                pivot.columns = ['_'.join(map(str, col)).strip() for col in pivot.columns]

            pivot = pivot.reset_index()
            sheet_name = filename.replace(".xlsx", "-汇总")
            pivot_tables[sheet_name] = pivot

        except Exception as e:
            print(f"❌ [{filename}] 生成透视失败: {e}")

    return pivot_tables
