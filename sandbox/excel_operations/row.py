from .formatting import get_copy_formatting_code


def add_row_operation(op: dict) -> str:
    """Tạo code để thêm row vào Excel document với định dạng giống các row khác.
    
    Args:
        op: Dictionary chứa 'data' (dict với key là tên cột và value là giá trị) 
            hoặc 'position' (vị trí chèn, mặc định là cuối)
    """
    data = op.get('data', {})
    position = op.get('position', None)  # None = thêm vào cuối
    
    code = f'''
# Add row operation with formatting
{get_copy_formatting_code()}

data = {repr(data)}
position = {repr(position)}

try:
    # Xác định vị trí row mới và row mẫu để sao chép định dạng
    if position is not None and position <= ws.max_row:
        # Chèn vào giữa, cần dịch chuyển các row xuống
        ws.insert_rows(position)
        new_row_idx = position
        # Lấy row ngay trước hoặc sau vị trí chèn làm mẫu
        source_row_idx = position - 1 if position > 1 else position + 1
    else:
        # Thêm vào cuối
        new_row_idx = ws.max_row + 1
        # Lấy row cuối cùng hiện có làm mẫu
        source_row_idx = ws.max_row
    
    # Thêm dữ liệu vào row mới
    for col_name, value in data.items():
        # Tìm index cột
        col_idx = None
        if col_name in df.columns:
            col_idx = list(df.columns).index(col_name) + 1
        else:
            # Thử tìm bằng tên cột trực tiếp (strip whitespace)
            for idx, col in enumerate(df.columns, start=1):
                if str(col).strip() == str(col_name).strip():
                    col_idx = idx
                    break
        
        if col_idx is None:
            print(f"- Cảnh báo: Không tìm thấy cột '{{col_name}}', bỏ qua")
            continue
        
        new_cell = ws.cell(row=new_row_idx, column=col_idx, value=value)
        
        # Sao chép định dạng
        source_cell = ws.cell(row=source_row_idx if source_row_idx > 0 else 1, column=col_idx)
        copy_cell_formatting(source_cell, new_cell)
        
        # Tinh chỉnh định dạng số nếu cần
        if isinstance(value, (int, float)):
            if new_cell.number_format == 'General' or not new_cell.number_format:
                is_ratio = any(keyword in str(col_name).lower() for keyword in ['margin', 'ratio', 'rate', '%'])
                if is_ratio:
                    new_cell.number_format = '0.00%'
                elif isinstance(value, float):
                    new_cell.number_format = '#,##0.00' if abs(value) >= 1 else '0.00'
                else:
                    new_cell.number_format = '#,##0'
    
    # Cập nhật dataframe logic (optional for simple displays)
    print(f"- Đã thêm dòng mới tại vị trí {{new_row_idx}} với định dạng chuyên nghiệp")
except Exception as e:
    print(f"- Lỗi thêm dòng: {{e}}")
    import traceback
    traceback.print_exc()
'''
    return code


def delete_rows_operation(op: dict) -> str:
    """Tạo code để xóa rows trong Excel document.
    
    Args:
        op: Dictionary chứa 'column' và 'value' để tìm rows cần xóa
        
    Returns:
        Python code để thực hiện xóa rows
    """
    col = op.get('column')
    val = op.get('value')
    
    code = f'''
# Delete rows operation
col = {repr(col)}
val = {repr(val)}
try:
    # Tìm cột
    col_found = None
    for c in df.columns:
        if str(c).strip() == str(col).strip():
            col_found = c
            break
    
    if col_found:
        rows_to_delete = df[df[col_found].astype(str) == str(val)].index.tolist()
        for idx in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(idx + 2)
        print(f"- Đã xóa {{len(rows_to_delete)}} dòng có {{col}} = '{{val}}'")
    else:
        print(f"- Không tìm thấy cột '{{col}}' để xóa dòng")
except Exception as e:
    print(f"- Lỗi xóa dòng: {{e}}")
'''
    return code
