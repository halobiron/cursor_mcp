"""Replace text operation for Word documents."""


def replace_text_operation(op: dict) -> str:
    """Tạo code để thay thế text trong Word document.
    
    Args:
        op: Dictionary chứa 'old' và 'new' text
        
    Returns:
        Python code để thực hiện thay thế
    """
    old_text = op.get('old')
    new_text = op.get('new')
    
    code = f'''
# Replace text operation
old_text = {repr(old_text)}
new_text = {repr(new_text)}
count = 0
for paragraph in doc.paragraphs:
    if old_text in paragraph.text:
        paragraph.text = paragraph.text.replace(old_text, new_text)
        count += 1
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            if old_text in cell.text:
                cell.text = cell.text.replace(old_text, new_text)
                count += 1
print(f"- Đã thay thế '{{old_text}}' thành '{{new_text}}' tại {{count}} vị trí.")
'''
    return code
