"""
FR-02 drift-prevention gate: capabilities() must cover all MCP tools in server.py.

Asserts that nothing registered in the FastMCP server is missing from
capabilities(), and that any tools claimed but not registered are only the
known platform-conditional COM tools (Windows + PPT_ENABLE_COM=true).
"""
from __future__ import annotations

# COM-conditional tools are only registered when _COM_AVAILABLE is True (Windows).
# On Ubuntu CI they are absent from the server but present in com_tools — that is
# expected and intentional, not drift.
_COM_CONDITIONAL: frozenset[str] = frozenset(
    {"export_slide_as_png", "export_deck_as_pdf", "run_slide_show", "recalculate_charts"}
)


def _get_server_tool_names() -> set[str]:
    """Return the set of tool names registered in the FastMCP server."""
    from pptmcp import server

    mcp = server.mcp
    # Try internal attribute paths first (avoids asyncio overhead)
    if hasattr(mcp, "_tool_manager") and hasattr(mcp._tool_manager, "_tools"):
        return set(mcp._tool_manager._tools)
    if hasattr(mcp, "_tools"):
        return set(mcp._tools)
    # FastMCP 3.x stores tools in _local_provider._components with keys "tool:<name>@"
    if hasattr(mcp, "_local_provider") and hasattr(mcp._local_provider, "_components"):
        return {
            key.split(":")[1].split("@")[0]
            for key in mcp._local_provider._components
            if key.startswith("tool:")
        }

    # H-009: No asyncio fallback — asyncio.run() in sync context on Windows requires
    # socket creation (ProactorEventLoop socketpair) which is blocked in unit tests.
    # If we reach here, the wrong FastMCP version is installed. Update the attribute
    # paths above to match the installed version (check fastmcp.__version__).
    raise RuntimeError(
        "Cannot discover MCP tool names: none of the known FastMCP internal "
        "attributes (_tool_manager._tools, _tools, _local_provider._components) "
        "are available. Check FastMCP version (tested against >=3.x)."
    )


def test_capabilities_drift() -> None:
    """capabilities() must cover every registered MCP tool (drift == 0).

    Failure messages identify:
    - Tools registered in server.py but absent from capabilities() — real drift.
    - Tools in capabilities() not registered AND not a known COM-conditional — real drift.
    """
    from pptmcp.presentation_pptx import capabilities

    caps = capabilities()

    # Collect all tool names declared in capabilities (tools + com_tools)
    def _names(entries: list) -> set[str]:
        result = set()
        for e in entries:
            if isinstance(e, dict):
                result.add(e["tool"])
            else:
                result.add(str(e))  # fallback: plain string entry
        return result

    caps_tools = _names(caps.get("tools", []))
    caps_com = _names(caps.get("com_tools", []))
    caps_all = caps_tools | caps_com

    server_tools = _get_server_tool_names()
    assert server_tools, "No tools discovered from MCP server — drift check is void"

    # 1. Everything registered must be declared in capabilities
    missing_from_caps = server_tools - caps_all
    assert not missing_from_caps, (
        f"DRIFT: tools registered in server.py but missing from capabilities():\n"
        f"  {sorted(missing_from_caps)}\n"
        f"  → Add these to presentation_pptx.capabilities() 'tools' list."
    )

    # 2. Everything in capabilities but not registered must be a known COM-conditional
    extra_in_caps = caps_all - server_tools
    undocumented_extra = extra_in_caps - _COM_CONDITIONAL
    assert not undocumented_extra, (
        f"DRIFT: capabilities() declares tools not registered in server.py:\n"
        f"  {sorted(undocumented_extra)}\n"
        f"  (COM-conditional extras allowed: {sorted(_COM_CONDITIONAL)})\n"
        f"  → Remove stale entries or register them in server.py."
    )
