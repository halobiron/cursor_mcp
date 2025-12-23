import asyncio
import os
import sys
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
import uuid

async def run_mcp_tool(action, arguments):
    """Helper function to run the python_sandbox tool with a specific action"""
    server_params = StdioServerParameters(
        command=sys.executable,
        args=[os.path.abspath("python_sandbox.py")],
        env=os.environ.copy()
    )

    full_args = {"action": action}
    full_args.update(arguments)

    async with stdio_client(server_params) as (read, write), ClientSession(read, write) as session:
        await session.initialize()
        print(f"\n>>> Running action: {action}")
        print(f"Arguments: {arguments}")
        result = await session.call_tool("python_sandbox", arguments=full_args)
        output = "\n".join(c.text for c in result.content if hasattr(c, 'text'))
        print(output)
        return output

async def main():
    session_id = str(uuid.uuid4())[:8]
    word_url = "https://github.com/python-openxml/python-docx/raw/master/tests/test_files/test.docx"
    excel_url = "https://go.microsoft.com/fwlink/?LinkID=521962"

    print("\n" + "="*80)
    print("BẮT ĐẦU KỊCH BẢN KIỂM TRA: TẢI VÀ CHỈNH SỬA WORD/EXCEL")
    print("="*80)

    # 1. Tải file Word
    print("\n--- BƯỚC 1: TẢI FILE WORD ---")
    await run_mcp_tool("download", {
        "document_url": word_url,
        "filename": "test.docx",
        "session_id": session_id
    })

    # 2. Tải file Excel
    print("\n--- BƯỚC 2: TẢI FILE EXCEL ---")
    await run_mcp_tool("download", {
        "document_url": excel_url,
        "filename": "financial.xlsx",
        "session_id": session_id
    })

    # 3. Đọc nội dung Word
    print("\n--- BƯỚC 3: ĐỌC NỘI DUNG WORD ---")
    await run_mcp_tool("python_sandbox", {
        "action": "read_word",
        "session_id": session_id
    })

    # 4. Chỉnh sửa Word
    print("\n--- BƯỚC 4: CHỈNH SỬA WORD ---")
    await run_mcp_tool("python_sandbox", {
        "action": "edit_word",
        "session_id": session_id,
        "operations": [
            {"type": "replace", "old": "python", "new": "MCP AI"},
            {"type": "insert_text", "text": "Đây là dòng được thêm tự động bởi MCP."}
        ]
    })

    # 5. Đọc nội dung Excel
    print("\n--- BƯỚC 5: ĐỌC NỘI DUNG EXCEL ---")
    await run_mcp_tool("python_sandbox", {
        "action": "read_excel",
        "session_id": session_id,
        "max_rows": 5
    })

    # 6. Chỉnh sửa Excel - Thêm cột tính toán
    print("\n--- BƯỚC 6: CHỈNH SỬA EXCEL (Thêm cột tính toán) ---")
    await run_mcp_tool("python_sandbox", {
        "action": "edit_excel",
        "session_id": session_id,
        "operations": [
            {"type": "add_column", "name": "Profit_Margin", "formula": "Profit / Sales"}
        ]
    })

    # 7. Kiểm tra lại kết quả Excel
    print("\n--- BƯỚC 7: KIỂM TRA LẠI KẾT QUẢ EXCEL ---")
    await run_mcp_tool("python_sandbox", {
        "action": "read_excel",
        "session_id": session_id,
        "max_rows": 5
    })

    print("\n" + "="*80)
    print("KẾT THÚC KỊCH BẢN KIỂM TRA")
    print("="*80)

if __name__ == "__main__":
    asyncio.run(main())
