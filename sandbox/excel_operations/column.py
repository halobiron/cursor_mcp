"""Column operations for Excel documents."""

from .formatting import get_copy_formatting_code


def add_column_operation(op: dict) -> str:
    """Tạo code để thêm cột vào Excel document với định dạng giống các cột khác.
    
    Args:
        op: Dictionary chứa 'name' và 'formula' của cột mới
        
    Returns:
        Python code để thực hiện thêm cột với định dạng
    """
    col_name = op.get('name')
    formula = op.get('formula')
    
    code = f'''
# Add column operation with formatting
{get_copy_formatting_code()}

col_name = {repr(col_name)}
formula = {repr(formula)}
try:
    df[col_name] = df.eval(formula)
    new_col_idx = ws.max_column + 1
    
    # Tìm cột mẫu để sao chép định dạng (lấy cột cuối cùng trước cột mới)
    source_col_idx = ws.max_column
    
    # Thêm header
    new_header = ws.cell(row=1, column=new_col_idx, value=col_name)
    if source_col_idx > 0:
        source_header = ws.cell(row=1, column=source_col_idx)
        copy_cell_formatting(source_header, new_header)
    
    # Thêm dữ liệu và sao chép định dạng
    for i, val in enumerate(df[col_name], start=2):
        new_cell = ws.cell(row=i, column=new_col_idx, value=val)
        
        if source_col_idx > 0:
            source_cell = ws.cell(row=i, column=source_col_idx)
            copy_cell_formatting(source_cell, new_cell)
            
            # Tinh chỉnh định dạng số nếu cần
            if isinstance(val, (int, float)):
                # Nếu cột gốc là tiền tệ/số và chúng ta đang tính tỉ lệ (margin/ratio)
                is_ratio = any(keyword in col_name.lower() for keyword in ['margin', 'ratio', 'rate', '%'])
                
                if is_ratio:
                    new_cell.number_format = '0.00%'
                elif source_cell.number_format and source_cell.number_format != 'General':
                    # Giữ nguyên định dạng từ cột nguồn nếu nó đặc biệt
                    new_cell.number_format = source_cell.number_format
                else:
                    # Áp dụng định dạng số mặc định đẹp hơn
                    if isinstance(val, float):
                        if abs(val) < 1:
                            new_cell.number_format = '0.00'
                        else:
                            new_cell.number_format = '#,##0.00'
                    else:
                        new_cell.number_format = '#,##0'
    
    # Cập nhật AutoFilter để bao phủ cả cột mới
    from openpyxl.utils import get_column_letter
    full_range = f"A1:{{get_column_letter(new_col_idx)}}{{ws.max_row}}"
    ws.auto_filter.ref = full_range
        
    print(f"- Đã thêm cột '{{col_name}}' bằng công thức '{{formula}}' với định dạng chuyên nghiệp và cập nhật bộ lọc")
except Exception as e:
    print(f"- Lỗi thêm cột '{{col_name}}': {{e}}")
    import traceback
    traceback.print_exc()
'''
    return code
