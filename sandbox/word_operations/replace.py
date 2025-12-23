def replace_text_operation(op: dict) -> str:
    """Tạo code để thay thế text trong Word document.
    
    Args:
        op: Dictionary chứa 'old' và 'new' text
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
print(f"- Replaced '{{old_text}}' -> '{{new_text}}' (Found: {{count}})")
if count == 0:
    print(f"- Warning: '{{old_text}}' not found to replace!")
'''
    return code


def replace_paragraph_operation(op: dict) -> str:
    """Tạo code để thay thế toàn bộ nội dung một paragraph theo index.
    
    Args:
        op: Dictionary chứa 'index' và 'new_text'
    """
    index = op.get('index')
    new_text = op.get('new_text')
    
    code = f'''
# Replace paragraph by index operation
index = {index}
new_text = {repr(new_text)}
if 0 <= index < len(doc.paragraphs):
    doc.paragraphs[index].text = new_text
    print(f"- Replaced paragraph at index {{index}}")
else:
    print(f"- Error: Paragraph index {{index}} out of range (0-{{len(doc.paragraphs)-1}})")
'''
    return code
