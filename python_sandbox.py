from mcp.server.fastmcp import FastMCP
from sandbox import (
    execute_python_code,
    download_document,
    read_word_content,
    edit_word_document,
    read_excel_content,
    edit_excel_document
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
    - 'execute': Runs Python code. Requires 'code'. Use '/app/data' for file paths.
    - 'download': Downloads a file from 'document_url'. Optional: 'filename' (defaults to 'document').
    - 'read_word': Extracts text and table info from a Word file in the session.
    - 'read_excel': Reads an Excel file. Optional: 'sheet_name', 'max_rows' (default 10).
    - 'edit_word': Modifies a Word file. Requires 'operations' (list of dicts).
        Ops: {"type": "replace", "old": "text", "new": "text"}, {"type": "insert_text", "text": "..."}, 
             {"type": "insert_heading", "text": "...", "level": 1}, {"type": "delete_paragraph", "keyword": "..."},
             {"type": "custom_code", "code": "..."}
    - 'edit_excel': Modifies an Excel file. Requires 'operations' (list of dicts).
        Ops: {"type": "add_column", "name": "ColumnName", "formula": "ColA * ColB"}, 
             {"type": "filter", "column": "Col", "operator": ">", "value": 10},
             {"type": "update_cell", "row": 2, "column": "A", "value": 100},
             {"type": "delete_rows", "column": "Status", "value": "Error"},
             {"type": "create_summary", "sheet_name": "Summary"},
             {"type": "custom_code", "code": "..."}
             
    SESSION_ID:
    - ALWAYS reuse the 'session_id' returned from previous calls to maintain data persistence 
      between downloading, reading, and editing.
    """
    if action == "execute":
        return execute_python_code(code, session_id)
    elif action == "download":
        return download_document(document_url, filename or "document", session_id)
    elif action == "read_word":
        return read_word_content(session_id)
    elif action == "read_excel":
        return read_excel_content(session_id, sheet_name, max_rows)
    elif action == "edit_word":
        return edit_word_document(session_id, operations)
    elif action == "edit_excel":
        return edit_excel_document(session_id, operations)
    else:
        return f"Lỗi: Action '{action}' không hợp lệ."

if __name__ == "__main__":
    mcp.run()
