"""Row operations for Excel documents."""


def delete_rows_operation(op: dict) -> str:
    """Tạo code để xóa rows trong Excel document.
    
    Args:
        op: Dictionary chứa 'column' và 'value' để tìm rows cần xóa
        
    Returns:
        Python code để thực hiện xóa rows
    """
    col = op.get('column')
    val = op.get('value')
    
    code = f'''
# Delete rows operation
col = {repr(col)}
val = {repr(val)}
rows_to_delete = df[df[col].astype(str) == str(val)].index.tolist()
for idx in sorted(rows_to_delete, reverse=True):
    ws.delete_rows(idx + 2)
print(f"- Đã xóa {{len(rows_to_delete)}} dòng có {{col}} = '{{val}}'")
'''
    return code
