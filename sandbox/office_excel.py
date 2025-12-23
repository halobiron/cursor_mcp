"""Excel document operations module.

This module provides functions to read and edit Excel documents using
a modular approach with separate operation handlers.
"""

import json
from .executor import execute_python_code
from .excel_operations import (
    add_column_operation,
    filter_operation,
    update_cell_operation,
    delete_rows_operation,
    create_summary_operation,
)


def read_excel_content(session_id: str, sheet_name: str = None, max_rows: int = 10) -> str:
    """Đọc nội dung Excel.
    
    Args:
        session_id: ID của session để thực thi code
        sheet_name: Tên sheet cần đọc (None = đọc sheet đầu tiên)
        max_rows: Số dòng tối đa hiển thị
        
    Returns:
        Kết quả đọc nội dung Excel
    """
    code = f'''
import pandas as pd
import os
excel_files = [f for f in os.listdir('/app/data') if f.endswith(('.xlsx', '.xls'))]
if not excel_files:
    print("Không tìm thấy file Excel!")
    exit(1)
df = pd.read_excel(f'/app/data/{{excel_files[0]}}', sheet_name={repr(sheet_name)})
if isinstance(df, dict):
    # If multiple sheets, take the first one or print info
    first_sheet = list(df.keys())[0]
    print(f"Báo cáo: File có nhiều sheet. Đang hiển thị sheet đầu tiên: '{{first_sheet}}'")
    df = df[first_sheet]
print(df.head({max_rows}).to_string())
'''
    return execute_python_code(code, session_id)


def edit_excel_document(session_id: str, operations: list) -> str:
    """Chỉnh sửa file Excel (.xlsx) đã được tải lên bằng các lệnh cấu trúc.
    
    Args:
        session_id: ID của session để thực thi code
        operations: Danh sách các operations cần thực hiện
        
    Returns:
        Kết quả thực thi các operations
    """
    # Map operation types to their handler functions
    operation_handlers = {
        'add_column': add_column_operation,
        'filter': filter_operation,
        'update_cell': update_cell_operation,
        'delete_rows': delete_rows_operation,
        'create_summary': create_summary_operation,
    }
    
    # Build the operation code
    operations_code = []
    for op in operations:
        op_type = op.get('type')
        
        if op_type == 'custom_code':
            # Handle custom code directly
            custom_code = op.get('code')
            operations_code.append(f'''
# Custom code operation
custom_code = {repr(custom_code)}
try:
    exec(custom_code, globals())
    print(f"- Đã thực thi custom code thành công.")
except Exception as e:
    print(f"- Lỗi thực thi custom code: {{e}}")
''')
        elif op_type in operation_handlers:
            # Use the appropriate handler
            operations_code.append(operation_handlers[op_type](op))
        else:
            operations_code.append(f'''
print(f"- Cảnh báo: Operation type '{op_type}' không được hỗ trợ")
''')
    
    # Combine all operations into final code
    code = f'''
import pandas as pd
import openpyxl
from openpyxl.styles import Font
import os
import json

# Tìm file Excel
excel_files = [f for f in os.listdir('/app/data') if f.endswith(('.xlsx', '.xls')) and not f.endswith('_edited.xlsx')]
if not excel_files:
    print("Lỗi: Không tìm thấy file Excel nào!")
    exit(1)

filename = excel_files[0]
file_path = f'/app/data/{{filename}}'

print(f"Đang xử lý file: {{filename}}")

# Đọc file Excel
wb = openpyxl.load_workbook(file_path)
sheet_name = wb.sheetnames[0]
ws = wb[sheet_name]

# Đọc bằng pandas để xử lý dữ liệu dễ hơn
df = pd.read_excel(file_path, sheet_name=sheet_name)
df.columns = [c.strip() for c in df.columns]

{chr(10).join(operations_code)}

# Lưu file đã chỉnh sửa
base, ext = os.path.splitext(filename)
output_filename = f"{{base}}_edited{{ext}}"
wb.save(f'/app/data/{{output_filename}}')
print(f"\\nĐã lưu file chỉnh sửa: {{output_filename}}")
'''
    return execute_python_code(code, session_id)

