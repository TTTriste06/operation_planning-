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


def replace_all_names_with_mapping(all_names: pd.Series, mapping_df: pd.DataFrame) -> pd.Series:
    """
    对品名列表 all_names 应用新旧料号 + 替代料号替换，返回去重后的替换结果。

    参数：
        all_names: 原始品名列表（pd.Series）
        mapping_df: 新旧料号映射表，必须包含 '旧品名', '新品名', '替代品名1~4'

    返回：
        替换后的品名列表（pd.Series），已去重排序
    """
    if not isinstance(all_names, pd.Series) or mapping_df is None or mapping_df.empty:
        return all_names.dropna().drop_duplicates().sort_values().reset_index(drop=True)

    # 清洗新旧品名列
    mapping_df["旧品名"] = mapping_df["旧品名"].astype(str).str.strip()
    mapping_df["新品名"] = mapping_df["新品名"].astype(str).str.strip()

    # 新旧料号替换
    df_names = all_names.dropna().astype(str).str.strip().to_frame(name="品名")
    merged = df_names.merge(mapping_df[["旧品名", "新品名"]], how="left", left_on="品名", right_on="旧品名")
    merged["最终品名"] = merged["新品名"].where(
        merged["新品名"].notna() & (merged["新品名"].str.strip() != ""), merged["品名"]
    )
    all_names = merged["最终品名"]

    # 替代料号替换（替换前判断新品名是否为空）
    for i in range(1, 5):
        sub_col = f"替代品名{i}"
        if sub_col not in mapping_df.columns:
            continue

        mapping_df[sub_col] = mapping_df[sub_col].astype(str).str.strip()
        mapping_df["新品名"] = mapping_df["新品名"].astype(str).str.strip()

        valid_subs = mapping_df[
            mapping_df[sub_col].notna() &
            (mapping_df[sub_col] != "") &
            mapping_df["新品名"].notna() &
            (mapping_df["新品名"] != "")
        ]

        if not valid_subs.empty:
            sub_map = valid_subs.set_index(sub_col)["新品名"]
            all_names = all_names.replace(sub_map)

    return all_names.dropna().drop_duplicates().sort_values().reset_index(drop=True)


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

    """
    if verbose:
        st.success(f"✅ 新旧料号替换完成，共替换: {mask.sum()} 行")
    """

    mapped_keys = set(merged.loc[mask, name_col])
    return merged.drop(columns=["旧品名", "新品名"], errors="ignore"), mapped_keys

def apply_extended_substitute_mapping(df, mapping_df, field_map, verbose=True):
    """
    替代料号品名替换（仅品名字段替换，无聚合合并），避免重复替换并自动去重。

    参数：
        df: 原始 DataFrame（如安全库存、预测等）
        mapping_df: 新旧料号映射表
        field_map: 对应字段映射配置，如 {"品名": "品名"}
        verbose: 是否打印替换信息（可配合 Streamlit）

    返回：
        df: 替换后的 DataFrame，已去重
        matched_keys: 成功替换的新品名集合
    """
    name_col = field_map["品名"]
    df = df.copy()

    # 清洗品名列
    df[name_col] = df[name_col].astype(str).str.strip().str.replace("\n", "").str.replace("\r", "")

    matched_keys = set()
    already_replaced = set()

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
            old_name = row[sub_col]
            new_name = row["新品名"]

            # 只替换还未替换过的品名，防止重复替换
            mask = (df[name_col] == old_name) & (~df[name_col].isin(already_replaced))
            if mask.any():
                df.loc[mask, name_col] = new_name
                matched_keys.add(new_name)
                already_replaced.add(new_name)

    # 替换完成后按关键字段去重（根据你实际字段选择）
    key_fields = [col for col in ["晶圆品名", "规格", "品名"] if col in df.columns]
    if key_fields:
        df = df.drop_duplicates(subset=key_fields)
    else:
        df = df.drop_duplicates()
    """
    if verbose:
        print(f"✅ 替代品名替换完成，共替换 {len(matched_keys)} 种")
    """
    return df, matched_keys
