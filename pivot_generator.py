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

def generate_monthly_pivots(dataframes: dict, pivot_config: dict) -> dict:
    """
    根据配置为多个 DataFrame 生成透视表

    参数：
        dataframes: dict[str, pd.DataFrame]，key 为文件名或表名
        pivot_config: dict，包含每个表的 index, columns, values, aggfunc, date_format 等

    返回：
        dict[str, pd.DataFrame]，key 为新 sheet 名（加 -汇总）
    """
    pivot_tables = {}

    for filename, df in dataframes.items():
        if filename not in pivot_config:
            print(f"⚠️ 未找到 {filename} 的透视配置，跳过")
            continue

        config = pivot_config[filename]
        index = config["index"]
        columns = config["columns"]
        values = config["values"]
        aggfunc = config.get("aggfunc", "sum")
        date_format = config.get("date_format")

        df = df.copy()

        # 若列是日期，则格式化为月份
        if date_format:
            try:
                df[columns] = pd.to_datetime(df[columns], errors='coerce')
                df = df.dropna(subset=[columns])
                df[columns] = df[columns].dt.to_period("M").astype(str)
            except Exception as e:
                print(f"❌ 日期格式处理失败 [{filename}]：{e}")
                continue

        try:
            pivot = pd.pivot_table(
                df,
                index=index,
                columns=columns,
                values=values,
                aggfunc=aggfunc,
                fill_value=0
            )

            # 扁平化列名（适配多值透视）
            if isinstance(pivot.columns, pd.MultiIndex):
                pivot.columns = ['_'.join(map(str, col)).strip() for col in pivot.columns]

            pivot = pivot.reset_index()
            pivot_tables[f"{filename.replace('.xlsx', '')}-汇总"] = pivot

        except Exception as e:
            print(f"❌ 生成透视失败 [{filename}]：{e}")

    return pivot_tables
