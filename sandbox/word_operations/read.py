"""Read operations for Word documents."""

from ..executor import execute_python_code


def read_word_content(session_id: str) -> str:
    """Đọc nội dung Word.
    
    Args:
        session_id: ID của session để thực thi code
        
    Returns:
        Kết quả đọc nội dung Word
    """
    code = '''
from docx import Document
import os
docx_files = [f for f in os.listdir('/app/data') if f.endswith('.docx')]
if not docx_files:
    print("Không tìm thấy file Word!")
    exit(1)
doc = Document(f'/app/data/{docx_files[0]}')
for p in doc.paragraphs:
    if p.text.strip(): print(f"- {p.text}")
'''
    return execute_python_code(code, session_id)
