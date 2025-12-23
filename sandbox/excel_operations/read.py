"""Read operations for Excel documents."""

from ..executor import execute_python_code


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
