# coding=utf-8
"""本机 Agent 使用的 MCP stdio 入口。"""
from .mcp_server import mcp_server


if __name__ == "__main__":
    mcp_server.run("stdio")
