import re

def process_history_columns(pivoted, config, selected_month):
    """
    将 <= selected_month 的订单/未交订单列合并为“历史订单数量”和“历史未交订单数量”

    参数:
    - pivoted: 透视后的 DataFrame（如包含“订单数量_2025-03”）
    - config: 当前 pivot 配置（需含 index）
    - selected_month: 截止月，如 "2025-03"
    """
    if not selected_month:
        return pivoted

    # 找出所有符合 "_YYYY-MM" 结尾的列
    month_pattern = re.compile(r"_(\d{4}-\d{2})$")
    history_order_cols = []
    history_pending_cols = []

    for col in pivoted.columns:
        match = month_pattern.search(col)
        if match:
            col_month = match.group(1)
            if col_month <= selected_month:
                if "订单数量" in col and "未交订单数量" not in col:
                    history_order_cols.append(col)
                elif "未交订单数量" in col:
                    history_pending_cols.append(col)

    # 添加新列合并值
    if history_order_cols:
        pivoted["历史订单数量"] = pivoted[history_order_cols].sum(axis=1)
    if history_pending_cols:
        pivoted["历史未交订单数量"] = pivoted[history_pending_cols].sum(axis=1)

    # 删除原始月份列
    pivoted.drop(columns=history_order_cols + history_pending_cols, inplace=True)

    # 插入合并列到合理位置
    fixed_cols = [col for col in pivoted.columns if col not in ['历史订单数量', '历史未交订单数量']]
    if '历史订单数量' in pivoted.columns:
        fixed_cols.insert(len(config['index']), '历史订单数量')
    if '历史未交订单数量' in pivoted.columns:
        fixed_cols.insert(len(config['index']) + 1, '历史未交订单数量')

    return pivoted[fixed_cols]
