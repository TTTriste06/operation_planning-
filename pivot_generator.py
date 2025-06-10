import pandas as pd

def generate_pivot_table(df: pd.DataFrame, group_cols: list, value_col: str = "数量") -> pd.DataFrame:
    """
    生成按指定字段分组的透视汇总表
    """
    missing_cols = [col for col in group_cols + [value_col] if col not in df.columns]
    if missing_cols:
        print(f"❌ 缺少字段: {missing_cols}")
        return pd.DataFrame()
    
    pivot = df.groupby(group_cols, as_index=False)[value_col].sum()
    pivot = pivot.rename(columns={value_col: f"{value_col}汇总"})
    return pivot


def generate_all_pivots(dataframes: dict) -> dict:
    """
    为指定表生成透视表，返回 sheet_name -> DataFrame 的字典
    """
    field_mappings = {
        "赛卓-未交订单": {
            "规格": "规格",
            "品名": "品名",
            "晶圆品名": "晶圆品名"
        },
        "赛卓-成品在制": {
            "规格": "产品规格",
            "品名": "产品品名",
            "晶圆品名": "晶圆型号",
            "封装形式": "封装形式",
            "供应商": "工作中心",
            "PC": "PC"
        },
        "赛卓-CP在制": {
            "规格": "产品规格",
            "品名": "产品品名",
            "晶圆品名": "晶圆型号",
            "供应商": "工作中心",
            "PC": "生管人员"
        },
        "赛卓-成品库存": {
            "规格": "规格",
            "品名": "品名",
            "晶圆品名": "WAFER品名"
        },
        "赛卓-晶圆库存": {
            "规格": "规格",
            "品名": "品名",
            "晶圆品名": "WAFER品名"
        }
    }

    value_cols_by_sheet = {
        "赛卓-未交订单": "未交订单数量",
        "赛卓-成品在制": "数量",
        "赛卓-CP在制": "数量",
        "赛卓-成品库存": "数量",
        "赛卓-晶圆库存": "数量"
    }

    pivot_tables = {}

    for sheet_name, mapping in field_mappings.items():
        df = dataframes.get(sheet_name)
        if df is not None and not df.empty:
            group_cols = [mapping[k] for k in mapping if mapping[k] in df.columns]
            value_col = value_cols_by_sheet.get(sheet_name, "数量")
            pivot_df = generate_pivot_table(df, group_cols, value_col)
            if not pivot_df.empty:
                pivot_tables[f"{sheet_name}-汇总"] = pivot_df

    return pivot_tables
