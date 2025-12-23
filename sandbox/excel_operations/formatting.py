"""Formatting utilities for Excel operations."""


def get_copy_formatting_code() -> str:
    """Trả về code Python để sao chép định dạng từ cell nguồn sang cell đích.
    
    Returns:
        Python code dạng string để sao chép định dạng
    """
    return '''
def copy_cell_formatting(source_cell, target_cell):
    """Sao chép định dạng từ source_cell sang target_cell một cách triệt để.
    """
    from copy import copy
    
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)
    
    # Sao chép độ rộng cột nếu là cell ở các dòng đầu tiên
    if source_cell.row == 1:
        from openpyxl.utils import get_column_letter
        source_col_letter = get_column_letter(source_cell.column)
        target_col_letter = get_column_letter(target_cell.column)
        if source_col_letter in source_cell.parent.column_dimensions:
            source_dim = source_cell.parent.column_dimensions[source_col_letter]
            target_cell.parent.column_dimensions[target_col_letter].width = source_dim.width
            target_cell.parent.column_dimensions[target_col_letter].hidden = source_dim.hidden

    # Sao chép độ cao dòng nếu là cell ở cột đầu tiên
    if source_cell.column == 1:
        source_row = source_cell.row
        target_row = target_cell.row
        if source_row in source_cell.parent.row_dimensions:
            source_dim = source_cell.parent.row_dimensions[source_row]
            target_cell.parent.row_dimensions[target_row].height = source_dim.height
            target_cell.parent.row_dimensions[target_row].hidden = source_dim.hidden
'''

