def insert_text_operation(op: dict) -> str:
    """Tạo code để thêm text vào Word document.
    
    Args:
        op: Dictionary chứa 'text' cần thêm
    """
    text = op.get('text')
    
    code = f'''
# Insert text operation
text = {repr(text)}
doc.add_paragraph(text)
print(f"- Added paragraph to the end of the document.")
'''
    return code


def insert_heading_operation(op: dict) -> str:
    """Tạo code để thêm heading vào Word document.
    
    Args:
        op: Dictionary chứa 'text' và 'level' của heading
    """
    text = op.get('text')
    level = op.get('level', 1)
    
    code = f'''
# Insert heading operation
text = {repr(text)}
level = {level}
doc.add_heading(text, level=level)
print(f"- Added heading level {{level}}: {{text}}")
'''
    return code
