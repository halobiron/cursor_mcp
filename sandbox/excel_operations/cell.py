"""Cell operations for Excel documents."""

from .formatting import get_copy_formatting_code


def update_cell_operation(op: dict) -> str:
    """Tạo code để cập nhật cell trong Excel document với định dạng giữ nguyên hoặc tối ưu.
    
    Args:
        op: Dictionary chứa 'row', 'column', và 'value'
        
    Returns:
        Python code để thực hiện cập nhật cell
    """
    row = op.get('row')
    col_name = op.get('column')
    value = op.get('value')
    
    code = f'''
# Update cell operation with formatting
{get_copy_formatting_code()}

row = {row}
col_name = {repr(col_name)}
value = {repr(value)}

# Tìm index cột nếu là tên
if isinstance(col_name, str) and col_name in df.columns:
    col_idx = list(df.columns).index(col_name) + 1
else:
    # Thử parse chữ cái cột (A, B, C...)
    try:
        from openpyxl.utils import column_index_from_string
        col_idx = column_index_from_string(str(col_name))
    except:
        col_idx = None

if col_idx:
    target_cell = ws.cell(row=row, column=col_idx)
    
    # Nếu cell chưa có style, thử sao chép từ cell cùng cột ở dòng khác
    if not target_cell.has_style:
        source_row = 2 if row != 2 else 1
        source_cell = ws.cell(row=source_row, column=col_idx)
        copy_cell_formatting(source_cell, target_cell)
    
    # Cập nhật giá trị
    target_cell.value = value
    
    # Tinh chỉnh định dạng số nếu là General và giá trị là số
    if isinstance(value, (int, float)) and (not target_cell.number_format or target_cell.number_format == 'General'):
        if isinstance(value, float):
            target_cell.number_format = '#,##0.00' if abs(value) >= 1 else '0.00'
        else:
            target_cell.number_format = '#,##0'

    print(f"- Đã cập nhật ô tại dòng {{row}}, cột {{col_name}} thành '{{value}}' với định dạng tối ưu")
else:
    print(f"- Không tìm thấy cột '{{col_name}}' để cập nhật")
'''
    return code
