from mcp.server.fastmcp import FastMCP
from sandbox import (
    execute_python_code,
    download_document,
    read_word_content,
    edit_word_document,
    read_excel_content,
    edit_excel_document,
    get_excel_sheets
)

# Khởi tạo MCP Server
mcp = FastMCP("Python Sandbox with Storage")

@mcp.tool()
def python_sandbox(
    action: str,
    code: str = None,
    document_url: str = None,
    filename: str = None,
    session_id: str = None,
    operations: list = None,
    sheet_name: str = None,
    max_rows: int = 10
) -> str:
    """
    Unified tool for Python execution and Office document manipulation (Word/Excel) in a sandboxed environment.
    
    This tool allows you to:
    1. Execute arbitrary Python code for data analysis, calculations, or file processing.
    2. Download documents from URLs into a persistent session workspace.
    3. Read and edit Word (.docx) and Excel (.xlsx) files using structured operations or custom code.
    
    ACTIONS:
    - 'execute': Runs Python code for general calculations or data processing. Requires 'code'. 
      IMPORTANT: Avoid using 'execute' to modify Word or Excel files. 
      Use 'edit_word' or 'edit_excel' with 'custom_code' instead to ensure proper file handling, 
      formatting preservation, and consistent session persistence.
    - 'download': Downloads a file from 'document_url'. Optional: 'filename' (defaults to 'document').
    - 'read_word': Extracts text and table info from a Word file in the session.
    - 'read_excel': Reads an Excel file. Optional: 'sheet_name', 'max_rows' (default 10).
      Tip: Call 'list_sheets' first to see available sheet names.
    - 'list_sheets': Lists all sheet names in an Excel file.
    - 'edit_word': Modifies a Word file. Requires 'operations' (list of dicts).
        Ops: {"type": "replace", "old": "text", "new": "text"}, 
             {"type": "replace_paragraph", "index": 0, "new_text": "..."},
             {"type": "insert_text", "text": "..."}, 
             {"type": "insert_heading", "text": "...", "level": 1}, 
             {"type": "delete_paragraph", "keyword": "..."},
             {"type": "custom_code", "code": "..."}
        Tip for 'custom_code': Use 'doc' (python-docx Document object) to modify content.
        The file is automatically saved after all operations.
    - 'edit_excel': Modifies an Excel file. Requires 'operations' (list of dicts).
        Ops: {"type": "add_column", "name": "ColumnName", "formula": "ColA * ColB"}, 
             {"type": "filter", "column": "Col", "operator": ">", "value": 10},
             {"type": "update_cell", "row": 2, "column": "A", "value": 100},
             {"type": "delete_rows", "column": "Status", "value": "Error"},
             {"type": "create_summary", "sheet_name": "Summary"},
             {"type": "custom_code", "code": "..."}
        Tip for 'custom_code': Use 'wb' (openpyxl workbook) and 'ws' (worksheet) for edits.
        Avoid 'ExcelWriter' or 'to_excel' as they lose formatting. Use 'ws.cell()' instead.
        Available helpers: 'copy_cell_formatting(src, dst)', 'apply_smart_format(cell, val)'.
        The workbook is automatically saved to an '_edited.xlsx' file.
             
    SESSION_ID:
    - ALWAYS reuse the 'session_id' returned from previous calls to maintain data persistence 
      between downloading, reading, and editing.

    FILENAME SELECTION:
    - If 'filename' is provided: The tool will use that specific file (e.g., 'document.docx').
    - If 'filename' is NOT provided: The tool automatically selects the LATEST edited version 
      (e.g., 'document_edited.docx') to support iterative editing.
    - IMPORTANT: To start a NEW edit from the ORIGINAL document after multiple previous edits, 
      you MUST explicitly specify the original filename (e.g., 'document.docx'). Otherwise, 
      it will keep editing the latest '_edited' version.
    """
    if action == "execute":
        return execute_python_code(code, session_id)
    elif action == "download":
        return download_document(document_url, filename or "document", session_id)
    elif action == "read_word":
        return read_word_content(session_id, filename)
    elif action == "read_excel":
        return read_excel_content(session_id, filename, sheet_name, max_rows)
    elif action == "list_sheets":
        return get_excel_sheets(session_id, filename)
    elif action == "edit_word":
        return edit_word_document(session_id, operations, filename)
    elif action == "edit_excel":
        return edit_excel_document(session_id, operations, filename, sheet_name)
    else:
        return f"Lỗi: Action '{action}' không hợp lệ."

if __name__ == "__main__":
    mcp.run()
