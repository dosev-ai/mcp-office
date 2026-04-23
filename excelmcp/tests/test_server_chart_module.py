"""Structural test: verify excelmcp.server_chart imports cleanly and
registers the unified chart tool on the shared mcp instance.

This test exercises the _server_instance.py decomposition pattern:
  - server_chart.py imports mcp from _server_instance (not standalone)
  - @mcp.tool() decorators fire at import time as side-effects
  - The shared mcp registry in _server_instance holds the registered tools
"""
import pytest


@pytest.mark.unit
def test_server_chart_module_importable():
    """server_chart.py imports without error and registers the chart tool."""
    import excelmcp.server_chart  # noqa: F401 — side-effect import

    from excelmcp._server_instance import mcp

    chart_tool_names = {
        "chart",
    }

    # FastMCP 3.x: use the public async list_tools() API (returns Sequence[Tool]).
    import asyncio
    tool_infos = asyncio.run(mcp.list_tools())
    registered = {t.name for t in tool_infos}
    missing = chart_tool_names - registered
    assert not missing, (
        f"Missing chart tools after importing server_chart: {missing}\n"
        f"Registered tools: {registered}"
    )
