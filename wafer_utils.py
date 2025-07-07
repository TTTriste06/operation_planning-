import pandas as pd

def extract_wafer_with_grossdie(main_plan_df: pd.DataFrame, df_grossdie: pd.DataFrame) -> pd.DataFrame:
    """
    从主计划中提取唯一“晶圆品名”，通过规格在 grossdie 表中查找对应的“单片数量”。

    参数：
        main_plan_df: 包含“晶圆品名”和“规格”的主计划 DataFrame
        df_grossdie: 包含“规格”和“GROSS DIE”的 grossdie DataFrame

    返回：
        df_unique_wafer: 包含“晶圆品名”和“单片数量”的 DataFrame
    """
    # 第一步：提取所有晶圆品名并去重
    wafer_names = (
        main_plan_df["晶圆品名"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .reset_index(drop=True)
    )
    df_unique_wafer = pd.DataFrame({"晶圆品名": wafer_names})

    # 第二步：构建晶圆品名 → 规格 的映射（取主计划中第一个匹配的规格）
    wafer_spec_map = (
        main_plan_df[["晶圆品名", "规格"]]
        .dropna()
        .astype(str)
        .apply(lambda x: x.str.strip())
        .drop_duplicates()
        .set_index("晶圆品名")
    )

    # 第三步：构建规格 → GROSS DIE 的映射
    df_grossdie_clean = df_grossdie[["规格", "GROSS DIE"]].dropna()
    df_grossdie_clean["规格"] = df_grossdie_clean["规格"].astype(str).str.strip()
    df_grossdie_clean["GROSS DIE"] = pd.to_numeric(df_grossdie_clean["GROSS DIE"], errors="coerce")
    grossdie_map = df_grossdie_clean.set_index("规格")["GROSS DIE"].to_dict()

    # 第四步：匹配规格 → GROSS DIE
    def get_grossdie(wafer_name):
        spec = wafer_spec_map.loc[wafer_name]["规格"] if wafer_name in wafer_spec_map.index else None
        return grossdie_map.get(spec, None)

    df_unique_wafer["单片数量"] = df_unique_wafer["晶圆品名"].apply(get_grossdie)

    return df_unique_wafer
