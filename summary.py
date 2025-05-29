import pandas as pd
import re
import streamlit as st
from openpyxl.styles import PatternFill


def merge_safety_inventory(summary_df, safety_df):
    """
    将安全库存表中 InvWaf 和 InvPart 信息按 '品名' 合并到汇总表中，仅根据 '品名' 匹配。

    参数:
    - summary_df: 汇总后的 DataFrame，含 '品名'
    - safety_df: 安全库存表，含 'ProductionNO.'、' InvWaf'、' InvPart'

    返回:
    - merged: 合并后的 DataFrame
    - unmatched_keys: list of 品名 未被匹配的记录
    """

    # 统一列名
    safety_df = safety_df.rename(columns={
        'ProductionNO.': '品名'
    }).copy()

    # 去重，避免品名重复导致合并爆炸
    safety_df = safety_df[['品名', ' InvWaf', ' InvPart']].drop_duplicates()

    # 获取所有品名键
    all_keys = set(safety_df['品名'].dropna().astype(str).str.strip())

    # 合并
    merged = summary_df.merge(
        safety_df,
        on='品名',
        how='left'
    )

    # 实际用到的品名
    used_keys = set(
        merged[~merged[[' InvWaf', ' InvPart']].isna().all(axis=1)]['品名']
        .dropna().astype(str).str.strip()
    )

    # 找出未被用到的品名
    unmatched_keys = list(all_keys - used_keys)

    return merged, unmatched_keys


def append_unfulfilled_summary_columns(summary_df, pivoted_df):
    """
    提取历史未交订单 + 各未来月份未交订单列，仅根据 '品名' 合并到 summary_df 中。
    返回合并后的 summary_df 和未匹配的品名列表。
    """

    # 匹配所有未交订单列
    unfulfilled_cols = [col for col in pivoted_df.columns if "未交订单数量" in col]
    unfulfilled_df = pivoted_df[["品名"] + unfulfilled_cols].copy()

    # 计算总未交订单
    unfulfilled_df["总未交订单"] = unfulfilled_df[unfulfilled_cols].sum(axis=1)

    # 整理列顺序
    ordered_cols = ["品名", "总未交订单"]
    if "历史未交订单数量" in pivoted_df.columns:
        ordered_cols.append("历史未交订单数量")
    ordered_cols += [col for col in unfulfilled_cols if col != "历史未交订单数量"]
    unfulfilled_df = unfulfilled_df[ordered_cols]

    # 查找未匹配的品名
    summary_keys = set(summary_df["品名"].dropna().astype(str).str.strip())
    unmatched_keys = [
        str(row["品名"]).strip()
        for _, row in unfulfilled_df.iterrows()
        if str(row["品名"]).strip() not in summary_keys
    ]

    # 合并
    merged = summary_df.merge(unfulfilled_df, on="品名", how="left")

    return merged, unmatched_keys



def append_forecast_to_summary(summary_df, forecast_df):
    """
    从预测表中提取与 summary_df 匹配的预测记录（仅按品名匹配），并返回未匹配的品名列表。

    参数:
    - summary_df: 汇总表（含“品名”列）
    - forecast_df: 原始预测表

    返回:
    - merged: 合并后的 summary_df
    - unmatched_keys: list of 品名 未被匹配的条目
    """

    # 重命名以统一列名
    forecast_df = forecast_df.rename(columns={
        "生产料号": "品名"
    })

    # 使用的唯一主键
    key_col = ["品名"]

    # 提取预测月份列
    month_cols = [col for col in forecast_df.columns if isinstance(col, str) and "预测" in col]
    if not month_cols:
        st.warning("⚠️ 没有识别到任何预测列，请检查列名是否包含'预测'")
        return summary_df, []

    # 去重，保留每个品名第一条记录
    # forecast_df = forecast_df[key_col + month_cols].drop_duplicates(subset=key_col)

    # 查找未匹配的品名
    summary_keys = set(summary_df["品名"].dropna().astype(str).str.strip())
    # st.write(forecast_df["品名"])
    forecast_keys = forecast_df["品名"].dropna().astype(str).str.strip()
    # st.write(forecast_keys)
    unmatched_keys = [key for key in forecast_keys if key not in summary_keys]

    # 合并
    merged = summary_df.merge(forecast_df, on="品名", how="left")
    return merged, unmatched_keys
    


def merge_finished_inventory(summary_df, finished_df):
    """
    仅按“品名”将成品库存数据合并进 summary_df，并返回未匹配的品名列表。

    参数:
    - summary_df: 汇总数据，包含“品名”
    - finished_df: 成品库存表，包含“品名”与库存列

    返回:
    - merged: 合并后的 DataFrame
    - unmatched_keys: list of 品名 未被匹配的条目
    """

    # 清理列名和统一主键
    finished_df.columns = finished_df.columns.str.strip()
    finished_df = finished_df.rename(columns={"WAFER品名": "晶圆品名"})  # 虽然不用了，为安全保留

    key_col = "品名"
    value_cols = ["数量_HOLD仓", "数量_成品仓", "数量_半成品仓"]

    for col in [key_col] + value_cols:
        if col not in finished_df.columns:
            st.error(f"❌ 缺失列：{col}")
            return summary_df, []

    # 去重，避免爆炸式合并
    finished_df = finished_df[[key_col] + value_cols].drop_duplicates()

    # 查找未匹配的品名
    summary_keys = set(summary_df[key_col].dropna().astype(str).str.strip())
    finished_keys = set(finished_df[key_col].dropna().astype(str).str.strip())
    unmatched_keys = list(finished_keys - summary_keys)

    # 合并库存信息
    merged = summary_df.merge(finished_df, on=key_col, how="left")

    return merged, unmatched_keys



def append_product_in_progress(summary_df, product_in_progress_df, mapping_df):
    """
    仅根据“品名”将“成品在制”和“半成品在制”数据合并进 summary_df，返回未匹配的品名列表。

    参数：
    - summary_df: 汇总表（含“品名”）
    - product_in_progress_df: 透视后的成品在制表，含“产品品名”及数值列
    - mapping_df: 新旧料号映射表，含“半成品”列

    返回：
    - summary_df: 合并了“成品在制”和“半成品在制”的 DataFrame
    - unmatched_keys: list of 未匹配的品名
    """

    summary_df = summary_df.copy()
    summary_df["成品在制"] = 0
    summary_df["半成品在制"] = 0

    numeric_cols = product_in_progress_df.select_dtypes(include='number').columns.tolist()

    used_keys = set()
    unmatched_keys = []

    # === 成品在制（按品名匹配）===
    for idx, row in product_in_progress_df.iterrows():
        part = str(row["产品品名"]).strip()
        mask = summary_df["品名"].astype(str).str.strip() == part
        if mask.any():
            summary_df.loc[mask, "成品在制"] = row[numeric_cols].sum()
            used_keys.add(part)
        else:
            unmatched_keys.append(part)

    # === 半成品在制 ===
    semi_rows = mapping_df[mapping_df["半成品"].notna() & (mapping_df["半成品"] != "")]
    semi_info_table = semi_rows[["新品名", "旧品名", "半成品"]].copy()
    semi_info_table["在制数量"] = 0

    check_log = []

    for idx, row in semi_info_table.iterrows():
        part_name = row["半成品"]
        matched = product_in_progress_df[product_in_progress_df["产品品名"] == part_name]

        if not matched.empty:
            value = matched[numeric_cols].sum().sum()
            source = "半成品匹配成功"
        else:
            value = 0
            source = "未匹配"

        semi_info_table.at[idx, "在制数量"] = value
        check_log.append({
            "半成品": part_name,
            "匹配值": value,
            "匹配来源": source
        })

    # 将半成品在制合并回新品名（品名）
    for idx, row in semi_info_table.iterrows():
        target_part = row["新品名"]
        value = row["在制数量"]
        mask = summary_df["品名"].astype(str).str.strip() == str(target_part).strip()
        if mask.any():
            summary_df.loc[mask, "半成品在制"] = value
            used_keys.add(target_part)
        else:
            unmatched_keys.append(target_part)

    # st.write("【半成品匹配日志】")
    # st.dataframe(pd.DataFrame(check_log))

    return summary_df, list(set(unmatched_keys) - used_keys)
