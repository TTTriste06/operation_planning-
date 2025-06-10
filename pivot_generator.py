import pandas as pd

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
    为多个 DataFrame 根据配置生成透视表。

    支持:
    - 可选 index 字段（optional_index）: 有则参与分组，无则略过；
    - 日期字段格式化为 %Y-%m；
    - 多 value 列聚合；
    - 缺失字段自动跳过。

    返回:
        dict[sheet_name -> pd.DataFrame]
    """
    pivot_tables = {}

    for filename, df in dataframes.items():
        if filename not in pivot_config:
            print(f"⚠️ 未找到 {filename} 的透视配置，跳过")
            continue

        config = pivot_config[filename]
        required_index = config.get("index", [])
        optional_index = config.get("optional_index", [])
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

        # 动态 index 组装
        index = [col for col in required_index if col in df.columns]
        available_optional = [col for col in optional_index if col in df.columns]
        index += available_optional

        # 确保至少含“品名”
        if "品名" in df.columns and "品名" not in index:
            index.append("品名")

        if not index:
            print(f"⚠️ {filename} 缺少分组字段，跳过")
            continue
            
       # 确保透视前，所有 index 字段均非 NaN
        for col in index:
            if col in df.columns:
                df[col] = df[col].fillna("    ").astype(str).str.strip()
        try:
            pivot = pd.pivot_table(
                df,
                index=index,
                columns=columns,
                values=values,
                aggfunc=aggfunc,
                fill_value=0
            )

            # 展平多级列名（多 values 时）
            if isinstance(pivot.columns, pd.MultiIndex):
                pivot.columns = ['_'.join(map(str, col)).strip() for col in pivot.columns]

            pivot = pivot.reset_index()
            sheet_name = filename.replace(".xlsx", "-汇总")
            pivot_tables[sheet_name] = pivot

        except Exception as e:
            print(f"❌ [{filename}] 生成透视失败: {e}")

    return pivot_tables
