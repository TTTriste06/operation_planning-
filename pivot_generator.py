import pandas as pd

def generate_pivot(df: pd.DataFrame, value_col: str = "数量") -> pd.DataFrame:
    """
    对给定的 DataFrame 按 '品名' 透视汇总 value_col 列
    """
    if "品名" not in df.columns or value_col not in df.columns:
        return pd.DataFrame()  # 必要字段不存在则返回空表

    pivot = df.groupby("品名", as_index=False)[value_col].sum()
    pivot = pivot.rename(columns={value_col: f"{value_col}汇总"})
    return pivot


def generate_all_pivots(dataframes: dict) -> dict:
    """
    针对指定的五个表生成透视表，返回一个 sheet_name -> df 的 dict
    """
    target_sheets = [
        "赛卓-成品在制",
        "赛卓-CP在制",
        "赛卓-成品库存",
        "赛卓-晶圆库存",
        "赛卓-未交订单"
    ]

    pivot_tables = {}
    for name in target_sheets:
        df = dataframes.get(name)
        if df is not None and isinstance(df, pd.DataFrame) and not df.empty:
            value_col = "数量"
            if name == "赛卓-未交订单":
                value_col = "未交订单数量"
            pivot_df = generate_pivot(df, value_col)
            if not pivot_df.empty:
                pivot_tables[f"{name}-汇总"] = pivot_df

    return pivot_tables
