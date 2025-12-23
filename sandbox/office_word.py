from .executor import execute_python_code
from .word_operations import (
    replace_text_operation,
    replace_paragraph_operation,
    insert_text_operation,
    insert_heading_operation,
    delete_paragraph_operation,
)


# Common helper function code to be injected
FIND_WORD_FILE_FUNC = '''
def find_word_file(target_name=None):
    files = os.listdir('/app/data')
    if target_name:
        if target_name in files: return target_name
        if not target_name.endswith('.docx'):
            if f"{target_name}.docx" in files: return f"{target_name}.docx"
    
    docx_files = [f for f in files if f.endswith('.docx')]
    # Ưu tiên file đã sửa trước đó
    edited_files = [f for f in docx_files if f.endswith('_edited.docx')]
    if edited_files: return edited_files[0]
    if docx_files: return docx_files[0]
    return None
'''


def read_word_content(session_id: str, filename: str = None) -> str:
    """Read Word content.
    
    Args:
        session_id: ID of the session to execute code
        filename: Specific filename to read (optional)
        
    Returns:
        Result of reading Word content
    """
    code = f'''
from docx import Document
import os

{FIND_WORD_FILE_FUNC}
filename = find_word_file({repr(filename)})
if not filename:
    print("Không tìm thấy file Word!")
    exit(1)

doc = Document(f'/app/data/{{filename}}')
print(f"Nội dung file: {{filename}}\\n")
for i, p in enumerate(doc.paragraphs):
    if p.text.strip(): 
        print(f"[{{i}}] {{p.text}}")
'''
    return execute_python_code(code, session_id)


def edit_word_document(session_id: str, operations: list, filename: str = None) -> str:
    """Edit Word document with structured operations.
    
    Args:
        session_id: ID of the session to execute code
        operations: List of operations to perform
        filename: Name of the file to edit (optional)
        
    Returns:
        Result of executing operations
    """
    # Map operation types to their handler functions
    operation_handlers = {
        'replace': replace_text_operation,
        'replace_paragraph': replace_paragraph_operation,
        'insert_text': insert_text_operation,
        'insert_heading': insert_heading_operation,
        'delete_paragraph': delete_paragraph_operation,
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
print(f"- Cảnh báo: Operation type '{{op_type}}' không được hỗ trợ")
''')
    
    # Combine all operations into final code
    code = f'''
import json
from docx import Document
import os

{FIND_WORD_FILE_FUNC}

filename = find_word_file({repr(filename)})
if not filename:
    print("Lỗi: Không tìm thấy file Word nào để chỉnh sửa!")
    exit(1)

base_name = filename.replace('_edited.docx', '').replace('.docx', '')
doc = Document(f'/app/data/{{filename}}')
print(f"Đang xử lý file: {{filename}}")

{chr(10).join(operations_code)}

# Lưu file kết quả
output_filename = f"{{base_name}}_edited.docx"
doc.save(f'/app/data/{{output_filename}}')
print(f"\\nĐã lưu file chỉnh sửa: {{output_filename}}")
'''
    return execute_python_code(code, session_id)