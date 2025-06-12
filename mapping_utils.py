import pandas as pd
import streamlit as st

def apply_all_name_replacements(df, mapping_new, mapping_sub, sheet_name, field_mappings, verbose=True):
    """
    对任意 DataFrame 表执行“新旧料号替换 + 替代料号替换”流程。
    会自动识别 FIELD_MAPPINGS 中定义的品名字段。

    参数：
        df: 要处理的 DataFrame（如预测、安全库存等）
        mapping_df: 新旧料号映射表，包含 '旧品名'、'新品名'、'替代品名1~4'
        sheet_name: 当前表名（必须出现在 field_mappings 中）
        field_mappings: 全局字段映射字典
        verbose: 是否输出替换信息

    返回：
        df: 替换后的 DataFrame
        all_mapped_keys: 所有被替换的新料号集合（主+替代）
    """
    if sheet_name not in field_mappings:
        raise ValueError(f"❌ FIELD_MAPPINGS 中未定义 {sheet_name} 的字段映射")

    field_map = field_mappings[sheet_name]

    if "品名" not in field_map:
        raise ValueError(f"❌ {sheet_name} 的字段映射中未指定 '品名'")

    actual_name_col = field_map["品名"]

    if actual_name_col not in df.columns:
        raise ValueError(f"❌ {sheet_name} 中未找到列：{actual_name_col}")

    # Step 1️⃣ 新旧料号替换
    df, mapped_main = apply_mapping_and_merge(df.copy(), mapping_new, {"品名": actual_name_col}, verbose=verbose)

    # Step 2️⃣ 替代品名替换
    df, mapped_sub = apply_extended_substitute_mapping(df, mapping_sub, {"品名": actual_name_col}, verbose=verbose)

    all_mapped_keys = mapped_main.union(mapped_sub)

    if verbose:
        print(f"✅ [{sheet_name}] 共完成替换: {len(all_mapped_keys)} 种新料号")

    return df, all_mapped_keys

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


def replace_all_names_with_mapping(all_names: pd.Series, mapping_new: pd.DataFrame, mapping_sub: pd.DataFrame) -> pd.Series:
    """
    对品名列表 all_names 应用新旧料号 + 替代料号替换，返回去重后的替换结果。

    参数：
        all_names: 原始品名列表（pd.Series）
        mapping_new: 新旧料号映射表，包含 '旧品名', '新品名'
        mapping_sub: 替代料号映射表，包含 '替代品名1~4', '新品名'

    返回：
        替换后的品名列表（pd.Series），已去重排序
    """
    if not isinstance(all_names, pd.Series):
        return all_names

    all_names = all_names.dropna().astype(str).str.strip()

    # 1️⃣ 新旧料号替换
    if mapping_new is not None and not mapping_new.empty:
        mapping_new = mapping_new.copy()
        mapping_new["旧品名"] = mapping_new["旧品名"].astype(str).str.strip()
        mapping_new["新品名"] = mapping_new["新品名"].astype(str).str.strip()

        df_names = all_names.to_frame(name="品名")
        merged = df_names.merge(
            mapping_new[["旧品名", "新品名"]],
            how="left",
            left_on="品名",
            right_on="旧品名"
        )
        merged["最终品名"] = merged["新品名"].where(
            merged["新品名"].notna() & (merged["新品名"].str.strip() != ""),
            merged["品名"]
        )
        all_names = merged["最终品名"]

    # 2️⃣ 替代料号替换
    if mapping_sub is not None and not mapping_sub.empty:
        mapping_sub = mapping_sub.copy()
        mapping_sub["新品名"] = mapping_sub["新品名"].astype(str).str.strip()

        for i in range(1, 5):
            sub_col = f"替代品名{i}"
            if sub_col not in mapping_sub.columns:
                continue

            mapping_sub[sub_col] = mapping_sub[sub_col].astype(str).str.strip()

            valid_subs = mapping_sub[
                mapping_sub[sub_col].notna() &
                (mapping_sub[sub_col] != "") &
                mapping_sub["新品名"].notna() &
                (mapping_sub["新品名"] != "")
            ]

            if not valid_subs.empty:
                sub_map = valid_subs.set_index(sub_col)["新品名"]
                all_names = all_names.replace(sub_map)

    # 去重排序后返回
    return all_names.dropna().drop_duplicates().sort_values().reset_index(drop=True)


def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True):
    name_col = field_map["品名"]
    df[name_col] = df[name_col].astype(str).str.strip()
    mapping_df["旧品名"] = mapping_df["旧品名"].astype(str).str.strip()
    mapping_df["新品名"] = mapping_df["新品名"].astype(str).str.strip()

    df = df[df[name_col] != ""].copy()

    merged = df.merge(mapping_df[["旧品名", "新品名"]], how="left", left_on=name_col, right_on="旧品名")
    mask = merged["新品名"].notna() & (merged["新品名"] != "")

    # 记录成功替换的新名字
    replaced_names = set(merged.loc[mask, "新品名"].dropna().unique())

    merged.loc[mask, name_col] = merged.loc[mask, "新品名"]
    merged["_由新旧料号映射"] = mask

    if verbose:
        st.write(f"✅ 新旧料号替换成功: {mask.sum()}，未匹配: {(~mask).sum()}")
        st.write(f"🧾 替换后的新品名共 {len(replaced_names)} 个")

    return merged.drop(columns=["旧品名", "新品名"]), list(replaced_names)

def apply_extended_substitute_mapping(df, mapping_df, field_map, replaced_main=None, verbose=True):
    name_col = field_map["品名"]
    df[name_col] = df[name_col].astype(str).str.strip()

    substitute_cols = [col for col in mapping_df.columns if "替代品名" in col]
    mapping_df = mapping_df.dropna(subset=substitute_cols, how='all')

    mapping_long = pd.melt(mapping_df, id_vars=["新品名"], value_vars=substitute_cols,
                           var_name="替代字段", value_name="替代品名").dropna(subset=["替代品名"])
    mapping_long["替代品名"] = mapping_long["替代品名"].astype(str).str.strip()

    df = df[df[name_col] != ""].copy()
    merged = df.merge(mapping_long[["替代品名", "新品名"]],
                      how="left", left_on=name_col, right_on="替代品名")
    mask = merged["新品名"].notna() & (merged["新品名"] != "")

    # 排除已经在主替换中替换过的（如传入了 replaced_main）
    if replaced_main:
        mask &= ~merged["新品名"].isin(replaced_main)

    # 记录成功替换的新名字
    replaced_substitute_names = set(merged.loc[mask, "新品名"].dropna().unique())

    merged.loc[mask, name_col] = merged.loc[mask, "新品名"]
    merged["_由替代料号映射"] = mask

    if verbose:
        st.write(f"✅ 替代料号替换成功: {mask.sum()}，未匹配: {(~mask).sum()}")
        st.write(f"🧾 替换后的新品名（替代料号）共 {len(replaced_substitute_names)} 个")

    return merged.drop(columns=["替代品名", "新品名"]), list(replaced_substitute_names)
