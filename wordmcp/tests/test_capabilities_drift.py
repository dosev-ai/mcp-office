"""Capabilities drift gate for wordmcp.

Asserts that nothing registered in the FastMCP server is missing from
capabilities(), and that any tools claimed but not registered are only the
known platform-conditional COM tools.

wordmcp registers ALL tools — including COM-backed ones — unconditionally at
module level: the @mcp.tool() decorator runs at import time regardless of whether
pywin32 is installed.  _require_com() only raises ToolError at *call* time.
Therefore _COM_CONDITIONAL is empty: every announced tool is always registered.
"""
from __future__ import annotations

from typing import Any

# COM tools are ALWAYS registered in wordmcp (decorator runs unconditionally;
# _require_com() raises ToolError at call time, not registration time).
_COM_CONDITIONAL: frozenset[str] = frozenset()


def _get_server_tool_names() -> set[str]:
    """Return the set of tool names registered in the FastMCP server instance."""
    import importlib
    import wordmcp.server as _ws_mod

    # Reload to get fresh @mcp.tool() registrations, avoiding mock/patch
    # contamination from other test modules that modify server internals.
    _ws_mod = importlib.reload(_ws_mod)
    mcp_any: Any = _ws_mod.mcp

    # FastMCP 2.x path 1
    if hasattr(mcp_any, "_tool_manager") and hasattr(mcp_any._tool_manager, "_tools"):
        return set(mcp_any._tool_manager._tools)
    # FastMCP 2.x path 2
    if hasattr(mcp_any, "_tools"):
        return set(mcp_any._tools)
    # FastMCP 3.x: _local_provider._components with keys "tool:<name>@<suffix>"
    if hasattr(mcp_any, "_local_provider") and hasattr(
        mcp_any._local_provider, "_components"
    ):
        return {
            key.split(":")[1].split("@")[0]
            for key in mcp_any._local_provider._components
            if key.startswith("tool:")
        }
    # H-009: No asyncio fallback — asyncio.run() in sync context on Windows
    # requires ProactorEventLoop socketpair which is blocked in unit tests.
    raise RuntimeError(
        "Cannot discover MCP tool names: none of the known FastMCP internal "
        "attributes (_tool_manager._tools, _tools, _local_provider._components) "
        "are available.  Check FastMCP version (tested against >=3.x)."
    )


def test_capabilities_drift() -> None:
    """capabilities() in wordmcp.server must cover every registered MCP tool.

    Failure messages identify:
    - Tools registered in server.py but absent from capabilities() — real drift.
    - Tools in capabilities() not registered AND not a known COM-conditional — real drift.
    """
    from wordmcp.server import capabilities

    caps = capabilities()

    # Aggregate all announced tool names from the top-level list and com_tools.
    # com_tools entries are already appended to caps["tools"] by server.py, but
    # we union both defensively in case the structure changes.
    caps_all: set[str] = set(caps.get("tools", []))
    for name in caps.get("com_tools", {}).get("tools", []):
        caps_all.add(name)

    server_tools = _get_server_tool_names()
    assert server_tools, "No tools discovered from MCP server — drift check is void"

    # 1. Every registered tool must be declared in capabilities
    missing_from_caps = server_tools - caps_all - _COM_CONDITIONAL
    assert not missing_from_caps, (
        f"DRIFT: tools registered in server.py but missing from capabilities():\n"
        f"  {sorted(missing_from_caps)}\n"
        f"  -> Add these to wordmcp.server.capabilities() tools list."
    )

    # 2. Every tool in capabilities must either be registered or be a known COM-conditional
    extra_in_caps = caps_all - server_tools
    undocumented_extra = extra_in_caps - _COM_CONDITIONAL
    assert not undocumented_extra, (
        f"DRIFT: capabilities() declares tools not registered in server:\n"
        f"  {sorted(undocumented_extra)}\n"
        f"  -> Remove stale entries from capabilities() or register them in server.py."
    )


def test_version_fidelity():
    """wordmcp.__version__ must match the version in pyproject.toml."""
    import importlib.metadata
    import wordmcp
    try:
        pkg_version = importlib.metadata.version("wordmcp")
        assert wordmcp.__version__ == pkg_version, (
            f"__version__ {wordmcp.__version__!r} != installed package {pkg_version!r}"
        )
    except importlib.metadata.PackageNotFoundError:
        # Editable install in CI may not register metadata — just check __version__ is not "unknown"
        assert wordmcp.__version__ not in ("", "unknown", "0.0.0", "0.1.0"), (
            f"__version__ looks like a placeholder: {wordmcp.__version__!r}"
        )


def test_capabilities_exposes_server_version():
    """capabilities() return dict must include a 'version' key matching __version__."""
    from wordmcp.document_docx import capabilities
    import wordmcp
    result = capabilities()
    assert "version" in result, "'version' key missing from capabilities() return dict"
    assert result["version"] == wordmcp.__version__, (
        f"capabilities()['version'] {result['version']!r} != __version__ {wordmcp.__version__!r}"
    )
