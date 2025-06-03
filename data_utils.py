import pandas as pd
import streamlit as st
import re

def extract_info(df, mapping, fields=("规格", "晶圆品名")):
    if df is None or df.empty:
        return pd.DataFrame(columns=["品名"] + list(fields))
    cols = {"品名": mapping.get("品名")}
    for f in fields:
        if f in mapping:
            cols[f] = mapping[f]
    try:
        sub = df[[cols["品名"]] + list(cols.values())[1:]].copy()
        sub.columns = ["品名"] + [f for f in fields if f in cols]
        return sub.drop_duplicates(subset=["品名"])
    except Exception:
        return pd.DataFrame(columns=["品名"] + list(fields))


def fill_spec_and_wafer_info(main_plan_df: pd.DataFrame, dataframes: dict, additional_sheets: dict, field_mappings: dict) -> pd.DataFrame:
    """
    为主计划 DataFrame 补全 规格 和 晶圆品名 字段，按优先级从多个数据源中逐步填充。

    参数：
        main_plan_df: 主计划表，含 '品名' 列
        dataframes: 主文件字典，来自 classify_files 后的 self.dataframes
        additional_sheets: 辅助表字典，如预测、新旧料号等
        field_mappings: 各表字段映射配置（FIELD_MAPPINGS）

    返回：
        已补全规格和晶圆品名的主计划表
    """
    sources = [
        ("赛卓-未交订单", ("规格", "晶圆品名")),
        ("赛卓-安全库存", ("规格", "晶圆品名")),
        ("赛卓-新旧料号", ("规格", "晶圆品名")),
        ("赛卓-成品在制", ("规格", "晶圆品名")),
        ("赛卓-成品库存", ("规格", "晶圆品名")),
        ("赛卓-预测", ("规格",))  # ❗预测中无晶圆品名
    ]

    for sheet, fields in sources:
        source_df = (
            dataframes.get(sheet)
            if sheet in dataframes
            else additional_sheets.get(sheet)
        )
        if source_df is None or source_df.empty:
            continue

        if sheet not in field_mappings:
            continue

        mapping = field_mappings[sheet]
        if "品名" not in mapping or not all(f in mapping for f in fields):
            continue

        # 构建映射列
        try:
            extracted = source_df.copy()
            extracted = extracted[[mapping["品名"]] + [mapping[f] for f in fields]]
            extracted.columns = ["品名"] + list(fields)
            extracted["品名"] = extracted["品名"].astype(str).str.strip()
            extracted = extracted.drop_duplicates(subset=["品名"])
        except Exception:
            continue

        # 合并并优先填入主列
        main_plan_df = main_plan_df.merge(extracted, on="品名", how="left", suffixes=("", f"_{sheet}"))
        for f in fields:
            alt_col = f"{f}_{sheet}"
            if alt_col in main_plan_df.columns:
                main_plan_df[f] = main_plan_df[f].combine_first(main_plan_df[alt_col])
                main_plan_df.drop(columns=[alt_col], inplace=True)

    return main_plan_df


def fill_packaging_info(main_plan_df, product_df, mapping_df, order_df, pc_df):
    """
    为主计划表填入封装厂、封装形式、PC 信息。
    """
    def strip_suffix(s):
        return str(s).split("-")[0].strip() if isinstance(s, str) else s

    name_col = "品名"
    pkg_col = "封装形式"
    vendor_col = "封装厂"

    # ✅ 从成品在制填封装厂与封装形式
    if product_df is not None and not product_df.empty:
        product_df = product_df.copy()
        product_df[name_col] = product_df["产品品名"].astype(str).str.strip()
        product_df[vendor_col] = product_df["工作中心"].astype(str).apply(strip_suffix)
        product_df[pkg_col] = product_df["封装形式"].astype(str).str.strip()

        merged = main_plan_df.merge(
            product_df[[name_col, vendor_col, pkg_col]].drop_duplicates(),
            on=name_col, how="left", suffixes=("", "_prod")
        )

        # 如果 merge 后列被加后缀，优先使用合并后的列
        main_plan_df[vendor_col] = main_plan_df.get(vendor_col, pd.Series(index=main_plan_df.index)).fillna(
            merged.get(vendor_col) or merged.get(f"{vendor_col}_prod")
        )
        main_plan_df[pkg_col] = main_plan_df.get(pkg_col, pd.Series(index=main_plan_df.index)).fillna(
            merged.get(pkg_col) or merged.get(f"{pkg_col}_prod")
        )

    # ✅ 从新旧料号补充封装厂
    if mapping_df is not None and not mapping_df.empty:
        mapping_df = mapping_df.copy()
        mapping_df["新品名"] = mapping_df["新品名"].astype(str).str.strip()
        mapping_df["封装厂"] = mapping_df["封装厂"].astype(str).apply(strip_suffix)

        merged = main_plan_df.merge(
            mapping_df[["新品名", "封装厂"]].drop_duplicates(),
            left_on=name_col, right_on="新品名", how="left", suffixes=("", "_map")
        )

        col_fallback = "封装厂_map" if "封装厂_map" in merged.columns else "封装厂"
        main_plan_df[vendor_col] = main_plan_df.get(vendor_col, pd.Series(index=main_plan_df.index)).fillna(
            merged.get(col_fallback)
        )

    # ✅ 从下单明细补充封装厂
    if order_df is not None and not order_df.empty:
        order_df = order_df.copy()
        order_df[name_col] = order_df["回货明细_回货品名"].astype(str).str.strip()
        order_df["封装厂"] = order_df["供应商名称"].astype(str).apply(strip_suffix)

        merged = main_plan_df.merge(
            order_df[[name_col, "封装厂"]].drop_duplicates(),
            on=name_col, how="left", suffixes=("", "_order")
        )

        col_fallback = "封装厂_order" if "封装厂_order" in merged.columns else "封装厂"
        main_plan_df[vendor_col] = main_plan_df.get(vendor_col, pd.Series(index=main_plan_df.index)).fillna(
            merged.get(col_fallback)
        )

    # ✅ 通过封装厂找到 PC
    if pc_df is not None and not pc_df.empty:
        pc_df = pc_df.copy()
        pc_df["封装厂"] = pc_df["封装厂"].astype(str).apply(strip_suffix)
        pc_df["PC"] = pc_df["PC"].astype(str).str.strip()

        merged = main_plan_df.merge(
            pc_df.drop_duplicates(subset=["封装厂"]),
            on="封装厂", how="left"
        )

        main_plan_df["PC"] = merged.get("PC", pd.Series(index=main_plan_df.index))

    return main_plan_df
