import pandas as pd

def extract_wafer_with_grossdie(main_plan_df: pd.DataFrame, df_grossdie: pd.DataFrame) -> pd.DataFrame:
    """
    从主计划中提取唯一“晶圆品名”，通过规格在 grossdie 表中查找对应的“单片数量”。

    返回包含“晶圆品名”和“单片数量”的 DataFrame。
    """
    # 提取唯一晶圆品名
    wafer_names = (
        main_plan_df["晶圆品名"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .reset_index(drop=True)
    )
    df_unique_wafer = pd.DataFrame({"晶圆品名": wafer_names})

    # 晶圆品名 -> 规格（可能一对多，先全部保留）
    wafer_spec_map = (
        main_plan_df[["晶圆品名", "规格"]]
        .dropna()
        .astype(str)
        .apply(lambda x: x.str.strip())
        .drop_duplicates()
    )

    # 规格 -> 单片数量映射
    df_grossdie_clean = df_grossdie[["规格", "GROSS DIE"]].dropna()
    df_grossdie_clean["规格"] = df_grossdie_clean["规格"].astype(str).str.strip()
    df_grossdie_clean["GROSS DIE"] = pd.to_numeric(df_grossdie_clean["GROSS DIE"], errors="coerce")
    grossdie_map = df_grossdie_clean.set_index("规格")["GROSS DIE"].to_dict()

    # 构造晶圆品名 → 规格 映射字典（只取第一个规格）
    wafer_to_spec = wafer_spec_map.groupby("晶圆品名")["规格"].first().to_dict()

    # 匹配单片数量
    def get_grossdie(wafer_name):
        spec = wafer_to_spec.get(wafer_name)
        return grossdie_map.get(spec, None)

    df_unique_wafer["单片数量"] = df_unique_wafer["晶圆品名"].apply(get_grossdie)

    return df_unique_wafer
