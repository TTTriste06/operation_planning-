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
    st.write("✅ 开始生成透视表")
    st.write("📂 可用数据表：", list(dataframes.keys()))
    st.write("🧩 配置文件：", list(pivot_config.keys()))
    pivot_tables = {}

    for filename, df in dataframes.items():
        if filename not in pivot_config:
            st.warning(f"⚠️ 未找到 {filename} 的透视配置，跳过")
            continue

        config = pivot_config[filename]
        index = config["index"]
        columns = config["columns"]
        values = config["values"]
        aggfunc = config.get("aggfunc", "sum")
        date_format = config.get("date_format")

        df = df.copy()

        # 日期格式处理
        if date_format:
            try:
                col = columns[0] if isinstance(columns, list) else columns
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df = df.dropna(subset=[col])
                df[col] = df[col].dt.to_period("M").astype(str)
            except Exception as e:
                st.error(f"❌ 日期字段格式化失败 [{filename}]：{e}")
                continue

        # 检查 index 是否都在
        if not all(col in df.columns for col in index):
            st.warning(f"⚠️ {filename} 缺少部分 index 字段，跳过")
            continue
        try:
            pivot = pd.pivot_table(
                df,
                index=index,
                columns=col,
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
            st.error(f"❌ [{filename}] 生成透视失败: {e}")

    return pivot_tables
