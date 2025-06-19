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
            "半成品","新规格","新晶圆品名","旧规格","旧晶圆品名"
        ]].copy()
    
        # 重命名为统一列名
        tmp.columns = ["半成品", "新规格", "新晶圆品名", "旧规格", "旧晶圆品名"]
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
            old_spec = row["旧规格"]
            spec_map[key] = new_spec if pd.notna(new_spec) and str(new_spec).strip() != "" else old_spec
    
            # 检查“新晶圆品名”是否为空或 NaN
            new_wafer = row["新晶圆品名"]
            old_wafer = row["旧晶圆品名"]
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

    优先从“赛卓-新旧料号”获取 PC；若无，再通过“封装厂”匹配“赛卓-供应商-PC”。
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

    # ========== 1️⃣ 封装厂、封装形式、PC（第一优先） ==========
    df_map = additional_sheets.get("赛卓-新旧料号")
    if df_map is not None and not df_map.empty:
        df_map = df_map.copy()
        df_map["新品名"] = df_map["新品名"].astype(str).str.strip()
        df_map["封装厂"] = df_map["封装厂"].astype(str).apply(normalize_vendor_name)
        df_map["封装形式"] = df_map["封装形式"].astype(str).str.strip()
        df_map["PC"] = df_map["PC"].astype(str).str.strip()

        for idx, row in main_plan_df.iterrows():
            pname = str(row[name_col]).strip()
            matched = df_map[df_map["新品名"] == pname]
            if matched.empty:
                continue

            if pd.isna(row[vendor_col]) and matched.iloc[0]["封装厂"]:
                main_plan_df.at[idx, vendor_col] = matched.iloc[0]["封装厂"]

            if pd.isna(row.get(pkg_col)) and matched.iloc[0]["封装形式"]:
                main_plan_df.at[idx, pkg_col] = matched.iloc[0]["封装形式"]

            # ✅ 优先填入 PC
            if "PC" not in main_plan_df.columns:
                main_plan_df["PC"] = ""
            if pd.isna(row.get("PC")) or row["PC"] == "":
                pc_value = matched.iloc[0]["PC"]
                if pc_value:
                    main_plan_df.at[idx, "PC"] = pc_value

    # ========== 2️⃣ 封装厂、封装形式补充（其他来源） ==========
    sources = [
        ("赛卓-成品在制", {"品名": "产品品名", "封装厂": "工作中心", "封装形式": "封装形式"}),
        ("赛卓-下单明细", {"品名": "回货明细_回货品名", "封装厂": "供应商名称"})
    ]

    for sheet, field_map in sources:
        df = dataframes.get(sheet) if sheet in dataframes else additional_sheets.get(sheet)
        if df is None or df.empty:
            continue

        df = df.copy()
        if field_map["品名"] not in df.columns or field_map["封装厂"] not in df.columns:
            continue

        df[field_map["品名"]] = df[field_map["品名"]].astype(str).str.strip()
        df[field_map["封装厂"]] = df[field_map["封装厂"]].astype(str).apply(normalize_vendor_name)
        if "封装形式" in field_map and field_map["封装形式"] in df.columns:
            df[field_map["封装形式"]] = df[field_map["封装形式"]].astype(str).str.strip()

        for idx, row in main_plan_df.iterrows():
            pname = str(row[name_col]).strip()
            matched = df[df[field_map["品名"]] == pname]
            if matched.empty:
                continue

            if pd.isna(row[vendor_col]):
                main_plan_df.at[idx, vendor_col] = matched.iloc[0][field_map["封装厂"]]

            if "封装形式" in field_map and pd.isna(row.get(pkg_col)):
                main_plan_df.at[idx, pkg_col] = matched.iloc[0][field_map["封装形式"]]

    # ========== 3️⃣ PC 补充：通过封装厂匹配 ==========
    pc_df = additional_sheets.get("赛卓-供应商-PC")
    
    if pc_df is not None and not pc_df.empty:
        pc_df = pc_df.copy()
        pc_df.columns = pc_df.columns.str.strip()  # 防止列名中有空格
        if "封装厂" not in pc_df.columns or "PC" not in pc_df.columns:
            raise ValueError("❌ ‘赛卓-供应商-PC’ 缺少必要字段：‘封装厂’ 或 ‘PC’")
    
        pc_df["封装厂"] = pc_df["封装厂"].astype(str).apply(normalize_vendor_name)
        pc_df["PC"] = pc_df["PC"].astype(str).str.strip()
    
        # 主表封装厂也标准化
        main_plan_df["封装厂"] = main_plan_df["封装厂"].astype(str).apply(normalize_vendor_name)
    
        if "PC" not in main_plan_df.columns:
            main_plan_df["PC"] = ""
    
        # 只填补空值
        mask_empty_pc = main_plan_df["PC"].isna() | (main_plan_df["PC"] == "")
        df_needs_pc = main_plan_df[mask_empty_pc].copy()
    
        # 执行 merge
        if "PC" in main_plan_df.columns:
            main_plan_df.drop(columns=["PC"], inplace=True)
        
        # 合并后只有一个 PC 列
        main_plan_df = main_plan_df.merge(
            pc_df[["封装厂", "PC"]].drop_duplicates(),
            on="封装厂",
            how="left"
        )


        st.write(pc_df)
        st.write(main_plan_df)
        
        # 🔒 检查 merge 后是否含 PC 列
        if "PC" not in main_plan_df.columns:
            raise ValueError("❌ 合并后没有生成 PC 列，可能‘供应商-PC’表格式错误或无匹配")
    
        # ✅ 回填 PC
        main_plan_df.loc[mask_empty_pc, "PC"] = main_plan_df["PC"].values
    
        # 可选调试
        filled_count = merged["PC"].notna().sum()
        st.write(f"✅ 通过封装厂补充 PC：成功填入 {filled_count} 条")
        st.write(main_plan_df)


    return main_plan_df
