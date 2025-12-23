"""Delete operations for Word documents."""


def delete_paragraph_operation(op: dict) -> str:
    """Tạo code để xóa paragraph trong Word document.
    
    Args:
        op: Dictionary chứa 'keyword' để tìm paragraph cần xóa
        
    Returns:
        Python code để thực hiện xóa paragraph
    """
    keyword = op.get('keyword')
    
    code = f'''
# Delete paragraph operation
keyword = {repr(keyword)}
paragraphs_to_delete = []
for i, paragraph in enumerate(doc.paragraphs):
    if keyword in paragraph.text:
        paragraphs_to_delete.append(i)
for i in reversed(paragraphs_to_delete):
    p = doc.paragraphs[i]
    p._element.getparent().remove(p._element)
print(f"- Đã xóa {{len(paragraphs_to_delete)}} đoạn văn chứa '{{keyword}}'")
'''
    return code
