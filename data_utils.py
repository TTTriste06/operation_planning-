import pandas as pd
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


def fill_packaging_info(main_plan_df, product_df, mapping_df, order_df, pc_df):
    """
    为主计划表填入封装厂、封装形式、PC 信息。

    参数：
        main_plan_df: 主计划 DataFrame，需含 "品名" 列
        product_df: 成品在制（赛卓-成品在制）
        mapping_df: 新旧料号（赛卓-新旧料号）
        order_df: 下单明细（赛卓-下单明细）
        pc_df: 封装厂-PC 映射（赛卓-供应商-PC），应含 "封装厂" 和 "PC"
    """
    def strip_suffix(s):
        return str(s).split("-")[0].strip() if isinstance(s, str) else s

    name_col = "品名"
    pkg_col = "封装形式"
    vendor_col = "封装厂"

    # ✅ 封装厂 & 封装形式（从成品在制）
    if product_df is not None and not product_df.empty:
        product_df = product_df.copy()
        product_df[name_col] = product_df["产品品名"].astype(str).str.strip()
        product_df[vendor_col] = product_df["封装厂"].astype(str).apply(strip_suffix)
        product_df[pkg_col] = product_df["封装形式"].astype(str).str.strip()

        merged = main_plan_df.merge(
            product_df[[name_col, vendor_col, pkg_col]].drop_duplicates(),
            on=name_col, how="left", suffixes=("", "_prod")
        )
        main_plan_df[vendor_col] = merged[vendor_col].fillna(main_plan_df.get(vendor_col))
        main_plan_df[pkg_col] = merged[pkg_col].fillna(main_plan_df.get(pkg_col))

    # ✅ 封装厂补充（从新旧料号）
    if mapping_df is not None and not mapping_df.empty:
        mapping_df = mapping_df.copy()
        mapping_df["新品名"] = mapping_df["新品名"].astype(str).str.strip()
        mapping_df["封装厂"] = mapping_df["封装厂"].astype(str).apply(strip_suffix)

        merged = main_plan_df.merge(
            mapping_df[["新品名", "封装厂"]].drop_duplicates(),
            left_on=name_col, right_on="新品名", how="left"
        )
        main_plan_df[vendor_col] = main_plan_df[vendor_col].fillna(merged["封装厂"])

    # ✅ 封装厂补充（从下单明细）
    if order_df is not None and not order_df.empty:
        order_df = order_df.copy()
        order_df[name_col] = order_df["回货明细_回货品名"].astype(str).str.strip()
        order_df["封装厂"] = order_df["封装厂"].astype(str).apply(strip_suffix)

        merged = main_plan_df.merge(
            order_df[[name_col, "封装厂"]].drop_duplicates(),
            on=name_col, how="left", suffixes=("", "_order")
        )
        main_plan_df[vendor_col] = main_plan_df[vendor_col].fillna(merged["封装厂"])

    # ✅ PC（通过封装厂）
    if pc_df is not None and not pc_df.empty:
        pc_df = pc_df.copy()
        pc_df["封装厂"] = pc_df["封装厂"].astype(str).apply(strip_suffix)
        pc_df["PC"] = pc_df["PC"].astype(str).str.strip()

        merged = main_plan_df.merge(
            pc_df.drop_duplicates(subset=["封装厂"]),
            on="封装厂", how="left"
        )
        main_plan_df["PC"] = merged["PC"]

    return main_plan_df
