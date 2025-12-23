"""Summary operations for Excel documents."""


def create_summary_operation(op: dict) -> str:
    """Tạo code để tạo summary sheet trong Excel document.
    
    Args:
        op: Dictionary chứa 'sheet_name' cho summary sheet
        
    Returns:
        Python code để thực hiện tạo summary
    """
    new_sheet_name = op.get('sheet_name', 'Summary')
    
    code = f'''
# Create summary operation
new_sheet_name = {repr(new_sheet_name)}
summary = df.describe()
if new_sheet_name in wb.sheetnames: del wb[new_sheet_name]
ws_sum = wb.create_sheet(new_sheet_name)

ws_sum.cell(row=1, column=1, value="Statistic")
for j, c in enumerate(summary.columns, start=2):
    ws_sum.cell(row=1, column=j, value=c)
for i, (idx, row) in enumerate(summary.iterrows(), start=2):
    ws_sum.cell(row=i, column=1, value=idx)
    for j, v in enumerate(row, start=2):
        ws_sum.cell(row=i, column=j, value=v)
print(f"- Đã tạo sheet tổng hợp '{{new_sheet_name}}'")
'''
    return code
