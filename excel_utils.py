from openpyxl.utils import get_column_letter

def adjust_column_width(ws, max_width=100):
    """
    自动调整 openpyxl 工作表中每列宽度，默认最大宽度为 max_width。
    """
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                cell_value = str(cell.value)
                if cell_value:
                    max_length = max(max_length, len(cell_value)) + 10
            except:
                pass
        adjusted_width = min(max_length + 10, max_width)
        ws.column_dimensions[col_letter].width = adjusted_width
