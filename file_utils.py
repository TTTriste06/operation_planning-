import pandas as pd
from collections import defaultdict

def merge_cp_files_by_keyword(cp_dataframes: dict) -> dict:
    grouped = defaultdict(list)

    # 将 DB, DB2, DB3... 聚合成同一组
    for key, df in cp_dataframes.items():
        for kw in ["华虹", "先进", "DB", "上华"]:
            if key.startswith(kw):
                grouped[kw].append(df)
                break

    # 合并
    merged_cp_dataframes = {}
    for kw, df_list in grouped.items():
        # 排除空 DataFrame
        df_list = [df for df in df_list if df is not None and not df.empty]
        if df_list:
            merged_cp_dataframes[kw] = pd.concat(df_list, ignore_index=True)
        else:
            merged_cp_dataframes[kw] = pd.DataFrame()

    return merged_cp_dataframes

