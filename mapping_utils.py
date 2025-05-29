import pandas as pd
import streamlit as st


def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True):
    spec_col = field_map["规格"]
    name_col = field_map["品名"]
    wafer_col = field_map["晶圆品名"]

    for col in [spec_col, name_col, wafer_col]:
        df[col] = df[col].astype(str).str.strip()
    for col in ["旧规格", "旧品名", "旧晶圆品名"]:
        mapping_df[col] = mapping_df[col].astype(str).str.strip()

    left_on = [spec_col, name_col, wafer_col]
    right_on = ["旧规格", "旧品名", "旧晶圆品名"]

    try:
        df_merged = df.merge(mapping_df, how="left", left_on=left_on, right_on=right_on)

        matched = df_merged["新规格"].notna()
        unmatched_count = (~matched).sum()

        # 生成布尔掩码：成功被新旧料号替换的行
        mask_None = (
            df_merged["新规格"].notna() & (df_merged["新规格"].astype(str).str.strip() != "") &
            df_merged["新品名"].notna() & (df_merged["新品名"].astype(str).str.strip() != "") &
            df_merged["新晶圆品名"].notna() & (df_merged["新晶圆品名"].astype(str).str.strip() != "")
        )

        df_merged["_由新旧料号映射"] = mask_None  # 标记列 ✅

        # 替换字段
        df_merged.loc[mask_None, spec_col] = df_merged.loc[mask_None, "新规格"]
        df_merged.loc[mask_None, name_col] = df_merged.loc[mask_None, "新品名"]
        df_merged.loc[mask_None, wafer_col] = df_merged.loc[mask_None, "新晶圆品名"]

        # 删除中间列
        drop_cols = ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
        df_cleaned = df_merged.drop(columns=[col for col in drop_cols if col in df_merged.columns])

        group_cols = [spec_col, name_col, wafer_col]
        numeric_cols = df_cleaned.select_dtypes(include="number").columns.tolist()
        sum_cols = [col for col in numeric_cols if col not in group_cols]

        df_grouped = df_cleaned.groupby(group_cols, as_index=False)[sum_cols].sum()

        other_cols = [col for col in df_cleaned.columns if col not in group_cols + sum_cols]
        if other_cols:
            df_first = df_cleaned.groupby(group_cols, as_index=False)[other_cols].first()
            df_grouped = pd.merge(df_grouped, df_first, on=group_cols, how="left")

        # ✅ 返回主键集合
        mapped_keys = set(
            tuple(df_merged.loc[idx, [spec_col, name_col, wafer_col]].values)
            for idx in df_merged.index[df_merged["_由新旧料号映射"]]
        )

        return df_grouped, mapped_keys

    except Exception as e:
        print(f"❌ 替换失败: {e}")
        return df, set()


def apply_mapping_and_merge_forecast(df, mapping_df, field_map, verbose=True):
    name_col = field_map["品名"]

    # 清洗字段
    df[name_col] = df[name_col].astype(str).str.strip()
    mapping_df["旧品名"] = mapping_df["旧品名"].astype(str).str.strip()
    mapping_df["新品名"] = mapping_df["新品名"].astype(str).str.strip()

    # 合并（仅按品名）
    try:
        df_merged = df.merge(mapping_df[["旧品名", "新品名"]], how="left", left_on=[name_col], right_on=["旧品名"])

        # 成功替换的标志
        mask_replace = df_merged["新品名"].notna() & (df_merged["新品名"].str.strip() != "")

        if verbose:
            st.write(f"✅ 替换成功行数: {mask_replace.sum()}")
            st.write(f"⚠️ 未匹配行数: {(~mask_replace).sum()}")

        df_merged["_由新旧料号映射"] = mask_replace

        # 替换品名
        df_merged.loc[mask_replace, name_col] = df_merged.loc[mask_replace, "新品名"]

        # 删除中间列
        df_cleaned = df_merged.drop(columns=[col for col in ["旧品名", "新品名"] if col in df_merged.columns])

        # 聚合字段
        group_cols = [name_col]
        numeric_cols = df_cleaned.select_dtypes(include="number").columns.tolist()
        sum_cols = [col for col in numeric_cols if col not in group_cols]

        df_grouped = df_cleaned.groupby(group_cols, as_index=False)[sum_cols].sum()

        # 保留其他非数值字段
        other_cols = [col for col in df_cleaned.columns if col not in group_cols + sum_cols]
        if other_cols:
            df_first = df_cleaned.groupby(group_cols, as_index=False)[other_cols].first()
            df_grouped = pd.merge(df_grouped, df_first, on=group_cols, how="left")

        # 返回映射键
        mapped_keys = set(
            tuple(df_merged.loc[idx, [name_col]].values)
            for idx in df_merged.index[df_merged["_由新旧料号映射"]]
        )

        return df_grouped, mapped_keys

    except Exception as e:
        st.error(f"❌ 替换失败: {e}")
        return df, set()


# 我们重新写一下apply_extended_substitute_mapping，首先找到K-V列不为空的所有行，只用保留新规格，新品名，新晶圆品名，和四组替代料号（替代规格, 替代品名, 替代晶圆），
# 如果这组替代料号不为空，就在需要合并的sheet中找到对应信息的行，替换替代规格, 替代品名, 替代晶圆为新规格，新品名，新晶圆品名，并与之前的新规格，新品名，新晶圆品名行合并

# 在合并两行的时候，只要这个sheet中对应的规格，品名，晶圆品名是一样的（每个sheet对应的名字不一样，FIELD_MAPPINGS有对照的），就需要合并这两行，如果别的非数字信息不一致，就取不是替换料号那一行的信息

def apply_extended_substitute_mapping(df, mapping_df, field_map, verbose=True):
    spec_col = field_map["规格"]
    name_col = field_map["品名"]
    wafer_col = field_map["晶圆品名"]

    # 清洗主字段
    for col in [spec_col, name_col, wafer_col]:
        df[col] = df[col].astype(str).str.strip().str.replace("\n", "").str.replace("\r", "")

    substitute_records = []

    for i in range(1, 5):
        sub_spec = f"替代规格{i}"
        sub_name = f"替代品名{i}"
        sub_wafer = f"替代晶圆{i}"

        for col in [sub_spec, sub_name, sub_wafer, "新规格", "新品名", "新晶圆品名"]:
            if col not in mapping_df.columns:
                mapping_df[col] = ""
            mapping_df[col] = mapping_df[col].astype(str).str.strip().str.replace("\n", "").str.replace("\r", "")

        filtered = mapping_df[
            mapping_df[[sub_spec, sub_name, sub_wafer, "新规格", "新品名", "新晶圆品名"]].notna().all(axis=1)
        ]

        for _, row in filtered.iterrows():
            substitute_records.append({
                "旧规格": row[sub_spec],
                "旧品名": row[sub_name],
                "旧晶圆品名": row[sub_wafer],
                "新规格": row["新规格"],
                "新品名": row["新品名"],
                "新晶圆品名": row["新晶圆品名"]
            })

    # 执行替换（不进行合并）
    matched_keys = set()
    for sub in substitute_records:
        mask = (
            (df[spec_col] == sub["旧规格"]) &
            (df[name_col] == sub["旧品名"]) &
            (df[wafer_col] == sub["旧晶圆品名"])
        )
        if mask.any():
            if verbose:
                st.write(
                    f"🔁 替换: ({sub['旧规格']}, {sub['旧品名']}, {sub['旧晶圆品名']}) → "
                    f"({sub['新规格']}, {sub['新品名']}, {sub['新晶圆品名']})，替换行数: {mask.sum()}"
                )
            df.loc[mask, spec_col] = sub["新规格"]
            df.loc[mask, name_col] = sub["新品名"]
            df.loc[mask, wafer_col] = sub["新晶圆品名"]
            matched_keys.update(
                tuple(x) for x in df.loc[mask, [spec_col, name_col, wafer_col]].values
            )

    if verbose:
        st.success(f"✅ 替代料号替换成功数: {len(matched_keys)}")

    return df, matched_keys
