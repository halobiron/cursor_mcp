import asyncio
import os
import sys
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

async def run_mcp_test():
    server_params = StdioServerParameters(
        command=sys.executable,
        args=[os.path.abspath("python_sandbox.py")],
        env=os.environ.copy()
    )

    async with stdio_client(server_params) as (read, write), ClientSession(read, write) as session:
        await session.initialize()
        
        code = r"""
import pandas as pd
data = {'Sản phẩm': ['Táo', 'Cam', 'Chuối'], 'Giá': [25000, 18000, 12000], 'Số lượng': [10, 5, 8]}
df = pd.DataFrame(data)
df['Thành tiền'] = df['Giá'] * df['Số lượng']
print(df)
df.to_csv('ban_hang.csv', index=False)
"""
        result = await session.call_tool("execute_python_code", arguments={"code": code.strip()})
        print(*(c.text for c in result.content if hasattr(c, 'text')), sep="\n")

if __name__ == "__main__":
    asyncio.run(run_mcp_test())
