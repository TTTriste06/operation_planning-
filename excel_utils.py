from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill

def adjust_column_width(ws, max_width=70):
    """
    自动调整工作表中每列的宽度，适配内容长度。
    忽略第一行（通常用于填充非数据内容），以第二行起为基准。
    """
    for col_cells in ws.iter_cols(min_row=2):  # 从第二行开始
        max_length = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_length + 8, max_width)

def highlight_replaced_names_in_main_sheet(ws, replaced_names: list[str], name_col_header: str = "品名"):
    """
    在 Excel 主计划 sheet 中，将所有替换过的新品名所在行标红。

    参数:
        workbook: openpyxl 加载后的 Workbook 对象
        sheet_name: 主计划 sheet 的名称
        replaced_names: 所有被替换的新名字列表
        name_col_header: 品名列的列名（默认为 "品名"）
    """
    red_fill = PatternFill(start_color="FFFF6666", end_color="FFFF6666", fill_type="solid")

    # 找到“品名”列的列索引
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    try:
        name_col_idx = header_row.index(name_col_header) + 1  # openpyxl 是从 1 开始
    except ValueError:
        raise ValueError(f"❌ 未找到“{name_col_header}”列，无法标红替换新品名")

    # 遍历数据行，标记匹配的品名行
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cell_value = str(row[name_col_idx - 1].value).strip()
        if cell_value in replaced_names:
            for cell in row:
                cell.fill = red_fill
