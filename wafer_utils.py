import pandas as pd

def extract_wafer_with_grossdie_raw(main_plan_df: pd.DataFrame, df_grossdie: pd.DataFrame) -> pd.DataFrame:
    """
    从主计划提取唯一晶圆品名，通过“规格”去 df_grossdie 中直接匹配 GROSS DIE，不清理 df_grossdie。

    参数：
        main_plan_df: 包含“晶圆品名”和“规格”的主计划表
        df_grossdie: 原始 grossdie 表，必须包含“规格”和“GROSS DIE”

    返回：
        DataFrame: 包含“晶圆品名”和“单片数量”的 DataFrame
    """
    # 获取唯一晶圆品名
    wafer_names = (
        main_plan_df["晶圆品名"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .reset_index(drop=True)
    )
    df_unique_wafer = pd.DataFrame({"晶圆品名": wafer_names})

    # 获取 晶圆品名 → 所有规格
    wafer_spec_map = (
        main_plan_df[["晶圆品名", "规格"]]
        .dropna()
        .astype(str)
        .apply(lambda x: x.str.strip())
        .drop_duplicates()
        .groupby("晶圆品名")["规格"]
        .apply(list)
        .to_dict()
    )

    # 不处理 df_grossdie，直接遍历查找 GROSS DIE
    def find_grossdie(wafer_name: str):
        specs = wafer_spec_map.get(wafer_name, [])
        for spec in specs:
            # 在 df_grossdie 中查找第一个匹配的规格
            match_row = df_grossdie[df_grossdie["规格"] == spec]
            if not match_row.empty:
                return match_row.iloc[0]["GROSS DIE"]
        return None

    # 匹配
    df_unique_wafer["单片数量"] = df_unique_wafer["晶圆品名"].apply(find_grossdie)

    return df_unique_wafer
