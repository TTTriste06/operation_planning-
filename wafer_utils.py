
def extract_wafer_with_grossdie(main_plan_df: pd.DataFrame, df_grossdie: pd.DataFrame) -> pd.DataFrame:
    """
    从主计划中提取唯一“晶圆品名”，并结合 grossdie 表通过“规格”匹配“单片数量”。
    
    参数：
        main_plan_df: 包含“晶圆品名”和“规格”的主计划 DataFrame
        df_grossdie: 包含“规格”和“GROSS DIE”的 grossdie DataFrame
    
    返回：
        包含“晶圆品名”和“单片数量”的 DataFrame
    """
    # 清洗 grossdie 表
    df_grossdie_clean = df_grossdie[["规格", "GROSS DIE"]].dropna()
    df_grossdie_clean["规格"] = df_grossdie_clean["规格"].astype(str).str.strip()
    df_grossdie_clean["GROSS DIE"] = pd.to_numeric(df_grossdie_clean["GROSS DIE"], errors="coerce")

    # 提取唯一晶圆品名
    unique_wafer = (
        main_plan_df["晶圆品名"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .reset_index(drop=True)
    )
    df_unique_wafer = pd.DataFrame({"晶圆品名": unique_wafer})

    # 构造 晶圆品名 - 规格 映射
    wafer_spec_map = (
        main_plan_df[["晶圆品名", "规格"]]
        .dropna()
        .drop_duplicates()
        .astype(str)
        .apply(lambda x: x.str.strip())
    )

    # 合并获取单片数量
    df_merged = pd.merge(
        df_unique_wafer.merge(wafer_spec_map, on="晶圆品名", how="left"),
        df_grossdie_clean.rename(columns={"GROSS DIE": "单片数量"}),
        on="规格",
        how="left"
    )

    # 返回结果：晶圆品名 + 单片数量
    df_result = df_merged[["晶圆品名", "单片数量"]].drop_duplicates().reset_index(drop=True)

    return df_result
