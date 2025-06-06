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


def fill_spec_and_wafer_info(main_plan_df: pd.DataFrame,
                              dataframes: dict,
                              additional_sheets: dict,
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
    nj_sheet_name = "赛卓-新旧料号"
    source_nj = (
        additional_sheets.get(nj_sheet_name)
    )
    if source_nj is not None and not source_nj.empty:
        mapping_nj = field_mappings[nj_sheet_name]
        st.write(mapping_nj)
        st.write(mapping_nj["半成品"])
        tmp = source_nj[[mapping_nj["半成品"],
                         mapping_nj["新规格"],
                         mapping_nj["新晶圆品名"]]].copy()
        st.write(tmp)
        tmp.columns = ["半成品", "新规格", "新晶圆品名"]
        tmp["半成品"] = tmp["半成品"].astype(str).str.strip()
        tmp = tmp.drop_duplicates(subset=["半成品"])

        st.write(tmp)

        # 构造从“半成品”到“新规格”和“新晶圆品名”的映射字典
        spec_map = dict(zip(tmp["半成品"], tmp["新规格"]))
        wafer_map = dict(zip(tmp["半成品"], tmp["新晶圆品名"]))

        # 找出 main_plan_df 中，品名正好等于某个“半成品”的行，进行覆盖
        mask = main_plan_df["品名"].astype(str).str.strip().isin(tmp["半成品"])
        if mask.any():
            # 将“新晶圆品名”覆盖到主表的“晶圆品名”列
            main_plan_df.loc[mask, "晶圆品名"] = main_plan_df.loc[mask, "品名"].map(wafer_map)

    return main_plan_df



def fill_packaging_info(main_plan_df, dataframes: dict, additional_sheets: dict) -> pd.DataFrame:
    """
    根据多个数据源填入封装厂、封装形式、PC。

    参数：
        main_plan_df: 主计划 DataFrame，含“品名”列
        dataframes: 所有主文件表格（如“赛卓-成品在制”等）
        additional_sheets: 所有辅助文件表格（如“赛卓-新旧料号”、“赛卓-供应商-PC”等）
    返回：
        填入字段后的主计划 DataFrame
    """
    # ✅ 封装厂别名映射
    VENDOR_ALIAS = {
        "绍兴千欣电子技术有限公司": "绍兴千欣",
        "南通宁芯": "南通宁芯微电子"
    }
    
    def normalize_vendor_name(name: str) -> str:
        name = str(name).strip()
        name = name.split("-")[0]  # 先去除 -CP 之类后缀
        return VENDOR_ALIAS.get(name, name)


    name_col = "品名"
    vendor_col = "封装厂"
    pkg_col = "封装形式"

    # ========== 1️⃣ 封装厂、封装形式 来源顺序 ==========
    sources = [
        ("赛卓-成品在制", {"品名": "产品品名", "封装厂": "工作中心", "封装形式": "封装形式"}),
        ("赛卓-新旧料号", {"品名": "新品名", "封装厂": "封装厂"}),  # 无封装形式
        ("赛卓-下单明细", {"品名": "回货明细_回货品名", "封装厂": "供应商名称"})  # 无封装形式
    ]

    for sheet, field_map in sources:
        df = dataframes.get(sheet) if sheet in dataframes else additional_sheets.get(sheet)
        if df is None or df.empty:
            continue

        df = df.copy()
        df[field_map["品名"]] = df[field_map["品名"]].astype(str).str.strip()
        df[field_map["封装厂"]] = df[field_map["封装厂"]].astype(str).apply(normalize_vendor_name)

        extract_cols = {
            name_col: df[field_map["品名"]],
            vendor_col: df[field_map["封装厂"]]
        }

        if "封装形式" in field_map:
            df[field_map["封装形式"]] = df[field_map["封装形式"]].astype(str).str.strip()
            extract_cols[pkg_col] = df[field_map["封装形式"]]

        extracted = pd.DataFrame(extract_cols).drop_duplicates()

        # 合并
        merged = main_plan_df.merge(extracted, on=name_col, how="left", suffixes=("", f"_{sheet}"))
        for col in [vendor_col, pkg_col]:
            alt_col = f"{col}_{sheet}"
            if alt_col in merged.columns:
                main_plan_df[col] = main_plan_df.get(col, pd.Series(index=main_plan_df.index)).combine_first(
                    merged[alt_col]
                )
                if alt_col in main_plan_df.columns:
                    main_plan_df.drop(columns=[alt_col], inplace=True)

    # ========== 2️⃣ 通过封装厂填入 PC ==========
    pc_df = additional_sheets.get("赛卓-供应商-PC")
    
    if pc_df is not None and not pc_df.empty:
        pc_df = pc_df.copy()
        pc_df["封装厂"] = pc_df["封装厂"].astype(str).apply(normalize_vendor_name)
        pc_df["PC"] = pc_df["PC"].astype(str).str.strip()
    
        # 删除 main_plan_df 中可能已有的 PC 列
        if "PC" in main_plan_df.columns:
            main_plan_df.drop(columns=["PC"], inplace=True)
    
        # 合并
        merged_pc = main_plan_df.merge(
            pc_df[["封装厂", "PC"]].drop_duplicates(),
            on="封装厂",
            how="left"
        )
    
        # 填回 PC 列
        main_plan_df["PC"] = merged_pc["PC"]
        
    return main_plan_df
