import pandas as pd
from datetime import datetime, timedelta

# 配置透视表字段
PIVOT_CONFIG = {
    "赛卓-未交订单": {
        "index": ["晶圆品名", "规格", "品名"],
        "columns": "预交货日",
        "values": ["订单数量", "未交订单数量"],
        "aggfunc": "sum",
        "date_format": "%Y-%m"
    },
    "赛卓-成品在制": {
        "index": ["工作中心", "封装形式", "晶圆型号", "产品规格", "产品品名"],
        "columns": "预计完工日期",
        "values": ["未交"],
        "aggfunc": "sum",
        "date_format": "%Y-%m"
    },
    "赛卓-成品库存": {
        "index": ["WAFER品名", "规格", "品名"],
        "columns": "仓库名称",
        "values": ["数量"],
        "aggfunc": "sum"
    },
    "赛卓-CP在制": {
        "index": ["晶圆型号", "产品品名"],
        "columns": "预计完工日期",
        "values": ["未交"],
        "aggfunc": "sum",
        "date_format": "%Y-%m"
    },
    "赛卓-晶圆库存": {
        "index": ["WAFER品名", "规格"],
        "columns": "仓库名称",
        "values": ["数量"],
        "aggfunc": "sum"
    }
}

# Excel 序列号转日期
def excel_serial_to_date(serial):
    try:
        base_date = datetime(1899, 12, 30)
        return base_date + timedelta(days=float(serial))
    except:
        return pd.NaT

# 创建透视表
def create_pivot_table(df, config):
    df = df.copy()
    if "date_format" in config:
        col = config["columns"]
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].apply(excel_serial_to_date)
        else:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        new_col = f"{col}_年月"
        df[new_col] = df[col].dt.strftime(config["date_format"])
        df[new_col] = df[new_col].fillna("未知日期")
        config = config.copy()
        config["columns"] = new_col

    pivot = pd.pivot_table(
        df,
        index=config["index"],
        columns=config["columns"],
        values=config["values"],
        aggfunc=config["aggfunc"],
        fill_value=0
    )

    pivot.columns = [
        f"{col[0]}_{col[1]}" if isinstance(col, tuple) else col
        for col in pivot.columns
    ]
    return pivot.reset_index()

# 批量透视

def generate_all_pivots(source_dataframes: dict) -> dict:
    pivot_tables = {}
    for sheet_name, config in PIVOT_CONFIG.items():
        if sheet_name in source_dataframes:
            try:
                df = source_dataframes[sheet_name]
                pivot_df = create_pivot_table(df, config)
                pivot_tables[f"{sheet_name}-汇总"] = pivot_df
            except Exception as e:
                pivot_tables[f"{sheet_name}-汇总"] = pd.DataFrame([{"错误": str(e)}])
    return pivot_tables
