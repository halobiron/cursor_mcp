"""Column operations for Excel documents."""


def add_column_operation(op: dict) -> str:
    """Tạo code để thêm cột vào Excel document.
    
    Args:
        op: Dictionary chứa 'name' và 'formula' của cột mới
        
    Returns:
        Python code để thực hiện thêm cột
    """
    col_name = op.get('name')
    formula = op.get('formula')
    
    code = f'''
# Add column operation
col_name = {repr(col_name)}
formula = {repr(formula)}
try:
    df[col_name] = df.eval(formula)
    new_col_idx = ws.max_column + 1
    ws.cell(row=1, column=new_col_idx, value=col_name)
    for i, val in enumerate(df[col_name], start=2):
        ws.cell(row=i, column=new_col_idx, value=val)
    print(f"- Đã thêm cột '{{col_name}}' bằng công thức '{{formula}}'")
except Exception as e:
    print(f"- Lỗi thêm cột '{{col_name}}': {{e}}")
'''
    return code
