import asyncio
import os
import sys
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

async def run_mcp_tool(tool_name, arguments):
    """Helper function to run an MCP tool"""
    server_params = StdioServerParameters(
        command=sys.executable,
        args=[os.path.abspath("python_sandbox.py")],
        env=os.environ.copy()
    )

    async with stdio_client(server_params) as (read, write), ClientSession(read, write) as session:
        await session.initialize()
        print(f"\nRunning tool: {tool_name}")
        result = await session.call_tool(tool_name, arguments=arguments)
        print(*(c.text for c in result.content if hasattr(c, 'text')), sep="\n")

async def main():
    shared_session_id = "test_session_3"
    
    # 1. Test Document Download
    print("\n" + "="*60 + "\nTEST 1: DOWNLOAD CSV\n" + "="*60)
    await run_mcp_tool("python_sandbox", {
        "action": "download",
        "document_url": "https://raw.githubusercontent.com/datasciencedojo/datasets/master/titanic.csv",
        "session_id": shared_session_id
    })

    # 1b. Test Processing (Tạo file titanic_summary.csv)
    print("\n" + "="*60 + "\nTEST 1B: PROCESS & CREATE CSV\n" + "="*60)
    await run_mcp_tool("python_sandbox", {
        "action": "execute",
        "code": """
import pandas as pd
df = pd.read_csv('/app/data/document')
print(f"Titanic dataset loaded. Saving to summary...")
df.head(10).to_csv('/app/data/titanic_summary.csv', index=False)
""",
        "session_id": shared_session_id
    })

    # 2. Test File Editing (Sửa file titanic_summary.csv từ Test 1)
    print("\n" + "="*60 + "\nTEST 2: EDIT FILE FROM PREVIOUS STEP\n" + "="*60)
    edit_code = """
import pandas as pd
import os

file_path = '/app/data/titanic_summary.csv'
if os.path.exists(file_path):
    df = pd.read_csv(file_path)
    # Thêm cột mới để đánh dấu đã sửa
    df['Status'] = 'Processed'
    df.to_csv(file_path, index=False)
    print(f"Successfully edited {file_path}")
    print(df.head(3))
else:
    print(f"File {file_path} not found!")
"""
    await run_mcp_tool("python_sandbox", {
        "action": "execute",
        "code": edit_code.strip(),
        "session_id": shared_session_id
    })

if __name__ == "__main__":
    asyncio.run(main())
