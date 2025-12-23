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
    add_row_operation,
    delete_rows_operation,
    create_summary_operation,
)

def read_excel_content(session_id: str, filename: str = None, sheet_name: str = None, max_rows: int = 10) -> str:
    """Đọc nội dung Excel.
    
    Args:
        session_id: ID của session để thực thi code
        filename: Tên file cụ thể (tùy chọn)
        sheet_name: Tên sheet cần đọc (None = đọc sheet đầu tiên)
        max_rows: Số dòng tối đa hiển thị
        
    Returns:
        Kết quả đọc nội dung Excel
    """
    code = f'''
import pandas as pd
import os

def find_excel_file(target_name=None):
    files = os.listdir('/app/data')
    if target_name:
        if target_name in files: return target_name
        for ext in ['.xlsx', '.xls']:
            if f"{{target_name}}{{ext}}" in files: return f"{{target_name}}{{ext}}"
    
    excel_files = [f for f in files if f.endswith(('.xlsx', '.xls'))]
    edited_files = [f for f in excel_files if f.endswith('_edited.xlsx')]
    if edited_files: return edited_files[0]
    if excel_files: return excel_files[0]
    if target_name and target_name in files: return target_name
    return None

filename = find_excel_file({repr(filename)})
if not filename:
    print("Không tìm thấy file Excel!")
    exit(1)

try:
    df = pd.read_excel(f'/app/data/{{filename}}', sheet_name={repr(sheet_name)})
    if isinstance(df, dict):
        first_sheet = list(df.keys())[0]
        print(f"Báo cáo: File có nhiều sheet. Đang hiển thị sheet đầu tiên: '{{first_sheet}}'")
        df = df[first_sheet]
    print(df.head({max_rows}).to_string())
except Exception as e:
    print(f"Lỗi: {{e}}")
    exit(1)
'''
    return execute_python_code(code, session_id)

def edit_excel_document(session_id: str, operations: list, filename: str = None, sheet_name: str = None) -> str:
    """Edit Excel document with structured operations.
    You should read the file first to get the content by using read_excel_content operation.
    Then you can edit the file by using other operations.
    
    Args:
        session_id: ID of the session to execute code
        operations: List of operations to perform
        filename: Name of the file to edit (optional)
        sheet_name: Name of the sheet to edit (optional, default to first sheet)
        
    Returns:
        Result of executing the operations
    """
    # Map operation types to their handler functions
    operation_handlers = {
        'add_column': add_column_operation,
        'filter': filter_operation,
        'update_cell': update_cell_operation,
        'add_row': add_row_operation,
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
            
            # VALIDATION: Phát hiện pattern sai sẽ làm mất định dạng
            dangerous_patterns = [
                ('pd.ExcelWriter', 'Sử dụng pd.ExcelWriter sẽ GHI ĐÈ và MẤT TẤT CẢ ĐỊNH DẠNG!'),
                ('ExcelWriter', 'Sử dụng ExcelWriter sẽ GHI ĐÈ và MẤT TẤT CẢ ĐỊNH DẠNG!'),
                ('writer.book', 'Pattern writer.book không hoạt động và sẽ gây lỗi!'),
                ('to_excel', 'Sử dụng df.to_excel() sẽ GHI ĐÈ và MẤT TẤT CẢ ĐỊNH DẠNG!'),
            ]
            
            warnings = []
            for pattern, msg in dangerous_patterns:
                if pattern in custom_code:
                    warnings.append(f"⚠️  CẢNH BÁO: {msg}")
            
            if warnings:
                warning_msg = '\\n'.join(warnings)
                operations_code.append(f'''
# Custom code operation - VỚI CẢNH BÁO
print("=" * 80)
print("⚠️  PHÁT HIỆN CODE NGUY HIỂM - CÓ THỂ LÀM MẤT ĐỊNH DẠNG!")
print("=" * 80)
print({repr(warning_msg)})
print()
print("ĐỂ GIỮ ĐỊNH DẠNG, CHỈ NÊN:")
print("  ✅ Sử dụng ws.cell() để thêm/sửa cell")
print("  ✅ Sử dụng wb.save() để lưu file")
print("  ❌ KHÔNG dùng pd.ExcelWriter hoặc df.to_excel()")
print()
print("Xem file EXCEL_FORMATTING_GUIDE.md để biết thêm chi tiết.")
print("=" * 80)
print()

custom_code = {repr(custom_code)}
try:
    exec(custom_code, globals())
    print(f"- Đã thực thi custom code (có cảnh báo).")
except Exception as e:
    print(f"- Lỗi thực thi custom code: {{e}}")
    import traceback
    traceback.print_exc()
    print()
    print("💡 GỢI Ý: Nếu lỗi liên quan đến 'writer.book' hoặc 'ExcelWriter',")
    print("   hãy xóa code đó và chỉ dùng wb.save() để lưu file.")
''')
            else:
                operations_code.append(f'''
# Custom code operation
custom_code = {repr(custom_code)}
try:
    exec(custom_code, globals())
    print(f"- Đã thực thi custom code thành công.")
except Exception as e:
    print(f"- Lỗi thực thi custom code: {{e}}")
    import traceback
    traceback.print_exc()
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

def find_excel_file(target_name=None):
    files = os.listdir('/app/data')
    if target_name:
        if target_name in files: return target_name
        for ext in ['.xlsx', '.xls']:
            if f"{{target_name}}{{ext}}" in files: return f"{{target_name}}{{ext}}"
    
    excel_files = [f for f in files if f.endswith(('.xlsx', '.xls'))]
    # Khi EDIT, ưu tiên file đã edit trước đó để có thể sửa tiếp (chaining)
    # Nhưng nếu edit lần đầu thì lấy file gốc
    edited_files = [f for f in excel_files if f.endswith('_edited.xlsx')]
    if edited_files: return edited_files[0]
    if excel_files: return excel_files[0]
    return None

filename = find_excel_file({repr(filename)})
if not filename:
    print("Lỗi: Không tìm thấy file Excel nào để chỉnh sửa!")
    exit(1)

file_path = f'/app/data/{{filename}}'
print(f"Đang xử lý file: {{filename}}")

# Đọc file Excel với tất cả định dạng
# data_only=False: giữ công thức thay vì chỉ giá trị
# keep_vba=True: giữ macro VBA nếu có
try:
    wb = openpyxl.load_workbook(file_path, data_only=False, keep_vba=True)
except Exception as e:
    # Nếu không thể load bằng openpyxl (có thể là .xls cũ), fallback sang thông báo lỗi cụ thể
    print(f"Lỗi: Không thể mở file bằng engine openpyxl (có thể file là định dạng .xls cũ hoặc bị lỗi). {{e}}")
    exit(1)

target_sheet = {repr(sheet_name)}
if target_sheet and target_sheet in wb.sheetnames:
    ws = wb[target_sheet]
else:
    target_sheet = wb.sheetnames[0]
    ws = wb[target_sheet]

print(f"Đang làm việc trên sheet: {{target_sheet}}")

# Đọc bằng pandas để xử lý dữ liệu dễ hơn
df = pd.read_excel(file_path, sheet_name=target_sheet)
df.columns = [c.strip() for c in df.columns]

{chr(10).join(operations_code)}

# Lưu file đã chỉnh sửa với tất cả định dạng
base, ext = os.path.splitext(filename)
# Đảm bảo ext luôn là .xlsx nếu file gốc không có ext hoặc là .xls (chuyển đổi sang .xlsx)
if not ext or ext.lower() == '.xls':
    ext = '.xlsx'

if base.endswith('_edited'):
    output_filename = f"{{base}}{{ext}}"
else:
    output_filename = f"{{base}}_edited{{ext}}"

save_path = f'/app/data/{{output_filename}}'

# Đảm bảo workbook properties được giữ nguyên
# Điều này giúp giữ các metadata như author, created date, etc.
try:
    # Save với tất cả các tùy chọn để giữ định dạng
    wb.save(save_path)
    print(f"\\nĐã lưu file chỉnh sửa: {{output_filename}}")
    print("Tất cả định dạng gốc đã được giữ nguyên.")
except Exception as e:
    print(f"Lỗi khi lưu file: {{e}}")
    # Thử lưu với cách khác nếu có lỗi
    try:
        wb.save(save_path)
        print(f"\\nĐã lưu file (fallback): {{output_filename}}")
    except Exception as e2:
        print(f"Không thể lưu file: {{e2}}")
        exit(1)
'''
    return execute_python_code(code, session_id)

