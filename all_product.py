import pandas as pd
from io import BytesIO

def export_distinct_new_products(mapping_df: pd.DataFrame, output_io: BytesIO = None) -> BytesIO:
    """
    从 mapping_df 中提取所有不同的 新规格、新品名、新晶圆品名，并导出为 Excel 文件。

    参数:
    - mapping_df: 包含新旧料号的 DataFrame。
    - output_io: 可选的 BytesIO 对象，用于写入 Excel。

    返回:
    - 写有 Excel 文件的 BytesIO 对象。
    """
    if output_io is None:
        output_io = BytesIO()

    # 抓取所需列并去重
    unique_products = mapping_df[["新规格", "新品名", "新晶圆品名"]].drop_duplicates()

    # 写入 Excel
    with pd.ExcelWriter(output_io, engine="openpyxl") as writer:
        unique_products.to_excel(writer, sheet_name="替换后产品列表", index=False)

    output_io.seek(0)
    return output_io
