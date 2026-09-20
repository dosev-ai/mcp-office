from __future__ import annotations


def test_public_server_imports_without_outlook_running() -> None:
    from mailmcp.server import mcp
    assert mcp is not None
