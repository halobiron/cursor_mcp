"""Filter operations for Excel documents."""


def filter_operation(op: dict) -> str:
    """Tạo code để lọc dữ liệu trong Excel document.
    
    Args:
        op: Dictionary chứa 'column', 'operator', và 'value' để lọc
        
    Returns:
        Python code để thực hiện lọc dữ liệu
    """
    col = op.get('column')
    op_sign = op.get('operator')
    val = op.get('value')
    
    code = f'''
# Filter operation
col = {repr(col)}
op_sign = {repr(op_sign)}
val = {repr(val)}

if op_sign == ">": df_filtered = df[df[col] > val]
elif op_sign == "<": df_filtered = df[df[col] < val]
elif op_sign == "==" or op_sign == "=": df_filtered = df[df[col] == val]

new_sheet_name = f"Filtered_{{col}}"
if new_sheet_name in wb.sheetnames: del wb[new_sheet_name]
ws_new = wb.create_sheet(new_sheet_name)

# Ghi header
for j, c in enumerate(df_filtered.columns, start=1):
    ws_new.cell(row=1, column=j, value=c)
# Ghi data
for i, row in enumerate(df_filtered.values, start=2):
    for j, v in enumerate(row, start=1):
        ws_new.cell(row=i, column=j, value=v)
print(f"- Đã lọc dữ liệu {{col}} {{op_sign}} {{val}} vào sheet '{{new_sheet_name}}'")
'''
    return code
