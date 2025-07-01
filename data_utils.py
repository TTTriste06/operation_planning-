import pandas as pd
import streamlit as st
import re

def fill_spec_and_wafer_info(main_plan_df: pd.DataFrame,
                              dataframes: dict,
                              additional_sheets: dict,
                              source_nj: pd.DataFrame,
                              field_mappings: dict) -> pd.DataFrame:
    """
    为主计划 DataFrame 补全 规格 和 晶圆品名 字段，按优先级从多个数据源中逐步填充。
    并且如果主计划中的“品名”正好匹配“赛卓-新旧料号”表里的“半成品”，
    就用对应行的“新规格”和“新晶圆品名”来覆盖主计划中的值。

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
        main_plan_df = main_plan_df.merge(
            extracted,
            on="品名",
            how="left",
            suffixes=("", f"_{sheet}")
        )
        for f in fields:
            alt_col = f"{f}_{sheet}"
            if alt_col in main_plan_df.columns:
                main_plan_df[f] = main_plan_df[f].combine_first(main_plan_df[alt_col])
                main_plan_df.drop(columns=[alt_col], inplace=True)

    # 额外处理：“赛卓-新旧料号”表里，如果主计划中的“品名”匹配到“半成品”，
    # 就用对应行的“新规格”和“新晶圆品名”来覆盖
    if source_nj is not None and not source_nj.empty:
        # 取出“半成品”“新规格”“新晶圆品名”“旧规格”“旧晶圆品名”五列
        tmp = source_nj[[
            "半成品","新规格","新晶圆品名"
        ]].copy()
    
        # 重命名为统一列名
        tmp.columns = ["半成品", "新规格", "新晶圆品名"]
        tmp["半成品"] = tmp["半成品"].astype(str).str.strip()
    
        # 如果同一个“半成品”多行，只保留第一行
        tmp = tmp.drop_duplicates(subset=["半成品"])
    
        # 构造映射：如果“新规格”非空则用“新规格”，否则用“旧规格”
        spec_map = {}
        wafer_map = {}
        for _, row in tmp.iterrows():
            key = row["半成品"]
            # 检查“新规格”是否为空或 NaN
            new_spec = row["新规格"]
            spec_map[key] = new_spec if pd.notna(new_spec) and str(new_spec).strip() != "" else old_spec
    
            # 检查“新晶圆品名”是否为空或 NaN
            new_wafer = row["新晶圆品名"]
            wafer_map[key] = new_wafer if pd.notna(new_wafer) and str(new_wafer).strip() != "" else old_wafer
    
        # 找出 main_plan_df 中，“品名”正好等于某个“半成品”的行
        mask = main_plan_df["品名"].astype(str).str.strip().isin(tmp["半成品"])
        if mask.any():
            # 用映射值覆盖“规格”和“晶圆品名”
            main_plan_df.loc[mask, "规格"] = main_plan_df.loc[mask, "品名"].map(spec_map)
            main_plan_df.loc[mask, "晶圆品名"] = main_plan_df.loc[mask, "品名"].map(wafer_map)

    return main_plan_df

def fill_packaging_info(main_plan_df, dataframes: dict, additional_sheets: dict) -> pd.DataFrame:
    """
    根据多个数据源填入封装厂、封装形式、PC。
    执行顺序：
    1. 从“成品在制”和“下单明细”补充封装厂和封装形式（补空）
    2. 从“供应商-PC”表按封装厂补 PC（补空）
    3. 从“新旧料号”强制覆盖封装厂、封装形式、PC
    """

    VENDOR_ALIAS = {
        "绍兴千欣电子技术有限公司": "绍兴千欣",
        "南通宁芯": "南通宁芯微电子"
    }

    def normalize_vendor_name(name: str) -> str:
        name = str(name).strip()
        name = name.split("-")[0]
        return VENDOR_ALIAS.get(name, name)

    name_col = "品名"
    vendor_col = "封装厂"
    pkg_col = "封装形式"

    # 保证必要字段存在
    for col in [vendor_col, pkg_col, "PC"]:
        if col not in main_plan_df.columns:
            main_plan_df[col] = ""

    # ========== 1️⃣ 从“成品在制”、“下单明细”补封装厂/封装形式（补空） ==========
    sources = [
        ("赛卓-成品在制", {"品名": "产品品名", "封装厂": "工作中心", "封装形式": "封装形式"}),
        ("赛卓-下单明细", {"品名": "回货明细_回货品名", "封装厂": "供应商名称"})  # 无封装形式
    ]

    for sheet, field_map in sources:
        df = dataframes.get(sheet) if sheet in dataframes else additional_sheets.get(sheet)
        if df is None or df.empty:
            continue

        df = df.copy()
        for col in field_map.values():
            if col not in df.columns:
                continue

        df[field_map["品名"]] = df[field_map["品名"]].astype(str).str.strip()
        df[field_map["封装厂"]] = df[field_map["封装厂"]].astype(str).apply(normalize_vendor_name)
        if "封装形式" in field_map and field_map.get("封装形式") in df.columns:
            df[field_map["封装形式"]] = df[field_map["封装形式"]].astype(str).str.strip()

        for idx, row in main_plan_df.iterrows():
            pname = str(row[name_col]).strip()
            matched = df[df[field_map["品名"]] == pname]
            if matched.empty:
                continue

            if pd.isna(row[vendor_col]) or row[vendor_col] == "":
                main_plan_df.at[idx, vendor_col] = matched.iloc[0][field_map["封装厂"]]
            if "封装形式" in field_map and (pd.isna(row[pkg_col]) or row[pkg_col] == ""):
                main_plan_df.at[idx, pkg_col] = matched.iloc[0][field_map["封装形式"]]

    # ========== 2️⃣ 从“供应商-PC”补 PC（补空） ==========
    pc_df = additional_sheets.get("赛卓-供应商-PC")
    if pc_df is not None and not pc_df.empty:
        pc_df = pc_df.copy()
        pc_df.columns = pc_df.columns.str.strip()
        if "封装厂" not in pc_df.columns or "PC" not in pc_df.columns:
            raise ValueError("❌ ‘赛卓-供应商-PC’ 缺少必要字段 ‘封装厂’ 或 ‘PC’")

        pc_df["封装厂"] = pc_df["封装厂"].astype(str).apply(normalize_vendor_name)
        pc_df["PC"] = pc_df["PC"].astype(str).str.strip()

        # 标准化主表封装厂
        main_plan_df["封装厂"] = main_plan_df["封装厂"].astype(str).apply(normalize_vendor_name)

        # 补 PC
        mask_empty_pc = main_plan_df["PC"].isna() | (main_plan_df["PC"] == "")
        df_needs_pc = main_plan_df[mask_empty_pc].copy()

        # 在合并前删除原有 PC 列，避免 _x/_y 出现
        if "PC" in df_needs_pc.columns:
            df_needs_pc.drop(columns=["PC"], inplace=True)
        
        merged = df_needs_pc.merge(
            pc_df[["封装厂", "PC"]].drop_duplicates(),
            on="封装厂",
            how="left"
        )
        
        if "PC" not in merged.columns:
            merged["PC"] = ""

        main_plan_df.loc[mask_empty_pc, "PC"] = merged["PC"].reset_index(drop=True)

    # ========== 3️⃣ 从“新旧料号”强制覆盖 封装厂 / 封装形式 / PC ==========
    df_map = additional_sheets.get("赛卓-新旧料号")
    if df_map is not None and not df_map.empty:
        df_map = df_map.copy()
        df_map.columns = df_map.columns.str.strip()
        df_map["旧品名"] = df_map["旧品名"].astype(str).str.strip()
        df_map["新品名"] = df_map["新品名"].astype(str).str.strip()
        df_map["封装厂"] = df_map["封装厂"].astype(str).apply(normalize_vendor_name)
        df_map["封装形式"] = df_map["封装形式"].astype(str).str.strip()
        df_map["PC"] = df_map["PC"].astype(str).str.strip()
    
        for idx, row in main_plan_df.iterrows():
            pname = str(row[name_col]).strip()
    
            # 先用旧品名匹配，若未命中再用新品名匹配
            matched = df_map[df_map["旧品名"] == pname]
            if matched.empty:
                matched = df_map[df_map["新品名"] == pname]
            if matched.empty:
                continue
    
            # ✅ 强制覆盖（优先级最高）
            main_plan_df.at[idx, vendor_col] = matched.iloc[0]["封装厂"]
            main_plan_df.at[idx, pkg_col] = matched.iloc[0]["封装形式"]
            main_plan_df.at[idx, "PC"] = matched.iloc[0]["PC"]

    return main_plan_df
