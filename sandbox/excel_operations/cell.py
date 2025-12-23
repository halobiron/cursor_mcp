"""Cell operations for Excel documents."""


def update_cell_operation(op: dict) -> str:
    """Tạo code để cập nhật cell trong Excel document.
    
    Args:
        op: Dictionary chứa 'row', 'column', và 'value'
        
    Returns:
        Python code để thực hiện cập nhật cell
    """
    row = op.get('row')
    col_name = op.get('column')
    value = op.get('value')
    
    code = f'''
# Update cell operation
row = {row}
col_name = {repr(col_name)}
value = {repr(value)}

# Tìm index cột nếu là tên
if isinstance(col_name, str) and col_name in df.columns:
    col_idx = list(df.columns).index(col_name) + 1
else:
    # Giả sử là chữ cái cột (A, B, C...)
    from openpyxl.utils import column_index_from_string
    col_idx = column_index_from_string(col_name)
    
ws.cell(row=row, column=col_idx, value=value)
print(f"- Đã cập nhật ô tại dòng {{row}}, cột {{col_name}} thành '{{value}}'")
'''
    return code
