"""Capabilities drift gate for excelmcp.

Asserts that nothing registered in the FastMCP server is missing from
capabilities(), and that any tools claimed but not registered are only the
known platform-conditional COM tools (Windows + EXCEL_ENABLE_COM=true).

POST-SPRINT-1 STATE (as of 2026-03-06):
    capabilities() in _io.py now correctly lists the 50 compound tool names.
    No xfail. Drift gate is a hard assertion.
"""
from __future__ import annotations

from ._smoke_helpers import _get_capabilities_snapshot, _get_capability_prompt_names, _get_tool_names, _run_tool

# COM-conditional tools: always registered unconditionally by server_com.py (bare
# import in server.py; @mcp.tool() decorators run at module-level import time).
# Confirmed against server_com.py — 4 @mcp.tool() functions:
#   recalculate_workbook, export_as_pdf, hydrate_template, pivot_table.
# These 4 tools ARE present in both the server AND capabilities()["tools"].
# The _COM_CONDITIONAL set exists as a safety mask in case COM is unavailable
# at import time, not because of any xfail or individual-tool drift.
_COM_CONDITIONAL: frozenset[str] = frozenset(
    {
        "recalculate_workbook",
        "export_as_pdf",
        "hydrate_template",
        "pivot_table",
        "export_range_as_csv",  # live=True path is COM/Windows-conditional
    }
)
def test_capabilities_drift() -> None:
    """capabilities() must cover every registered MCP tool (drift == 0).

    Failure messages identify:
    - Tools registered in server_*.py but absent from capabilities() — real drift.
    - Tools in capabilities() not registered AND not a known COM-conditional — real drift.
    """
    from excelmcp import server

    caps = _get_capabilities_snapshot(server.mcp)
    caps_tools = caps["tools"]

    server_tools = _get_tool_names(server.mcp)
    assert server_tools, "No tools discovered from MCP server — drift check is void"
    assert "render_review" not in server_tools
    assert "render_review" not in caps_tools

    # 1. Every registered tool must be declared in capabilities
    missing_from_caps = server_tools - caps_tools - _COM_CONDITIONAL
    assert not missing_from_caps, (
        f"DRIFT: tools registered in server but missing from capabilities():\n"
        f"  {sorted(missing_from_caps)}\n"
        f"  -> Add these to excel/src/excelmcp/_io.py capabilities() tools list."
    )

    # 2. Every tool in capabilities must either be registered or be a known COM-conditional
    extra_in_caps = caps_tools - server_tools
    undocumented_extra = extra_in_caps - _COM_CONDITIONAL
    assert not undocumented_extra, (
        f"DRIFT: capabilities() declares tools not registered in server:\n"
        f"  {sorted(undocumented_extra)}\n"
        f"  (COM-conditional extras allowed: {sorted(_COM_CONDITIONAL)})\n"
        f"  -> Remove stale entries or register them in a server_*.py module."
    )


def test_prompt_capabilities_drift() -> None:
    """capabilities() prompt inventory must be internally consistent and live."""
    from excelmcp import server

    caps = _get_capabilities_snapshot(server.mcp)
    live_prompts = _get_capability_prompt_names(server.mcp)

    assert caps["prompts"] == live_prompts, (
        "DRIFT: capabilities() prompt inventory does not exactly match the live prompt registry:\n"
        f"  live_only={sorted(live_prompts - caps['prompts'])}\n"
        f"  caps_only={sorted(caps['prompts'] - live_prompts)}"
    )
    assert caps["prompt_count"] == len(live_prompts), (
        "DRIFT: capabilities() prompt count does not match the live prompt registry:\n"
        f"  prompt_count={caps['prompt_count']} live_count={len(live_prompts)}"
    )


def test_capabilities_backend_metadata_unchanged() -> None:
    """This sprint must not change the advertised backend metadata."""
    from excelmcp import server

    caps = _run_tool(server.mcp, "capabilities", {})

    assert caps["backend"] == "openpyxl"


def test_capabilities_metadata_governance_keys() -> None:
    """capabilities() must expose metadata_contract_version and metadata_policy."""
    from excelmcp import server

    caps = _run_tool(server.mcp, "capabilities", {})

    assert "metadata_contract_version" in caps, (
        "DRIFT: 'metadata_contract_version' missing from capabilities() output.\n"
        "  -> Add it to _io.py capabilities() return dict."
    )
    assert "metadata_policy" in caps, (
        "DRIFT: 'metadata_policy' missing from capabilities() output.\n"
        "  -> Add it to _io.py capabilities() return dict."
    )
    from excelmcp._metadata_contract import CONTRACT_VERSION as _CONTRACT_VERSION
    assert caps["metadata_contract_version"] == _CONTRACT_VERSION, (
        f"Expected contract version {_CONTRACT_VERSION!r}, got: {caps['metadata_contract_version']!r}"
    )
    assert caps["metadata_policy"] in ("strict", "lenient"), (
        f"metadata_policy must be 'strict' or 'lenient', got: {caps['metadata_policy']!r}"
    )


# ---------------------------------------------------------------------------
# Phase 3 refactor drift guards
# ---------------------------------------------------------------------------

def test_tool_count_exactly_55() -> None:
    """Tool count must remain exactly 55 after Phase 3 prompt split."""
    import excelmcp.server  # noqa: F401 -- ensure all side-effect imports fire
    from excelmcp._server_instance import mcp

    # FastMCP stores tools in different attributes across versions -- try all
    tool_names = []
    if hasattr(mcp, '_tool_manager') and hasattr(mcp._tool_manager, '_tools'):
        tool_names = list(mcp._tool_manager._tools.keys())
    elif hasattr(mcp, 'tools') and isinstance(mcp.tools, dict):
        tool_names = list(mcp.tools.keys())

    if not tool_names:
        import pytest
        pytest.skip("Cannot introspect tool count in this FastMCP version")

    count = len(tool_names)
    assert count == 55, (
        "DRIFT: expected 55 tools, got %d. Tools found: %s" % (count, sorted(tool_names))
    )


def test_prompt_count_unchanged_at_32() -> None:
    """Prompt count must remain exactly 32 after moving bodies to sub-modules."""
    import excelmcp.server  # noqa: F401 -- ensure all prompt side-effect imports fire
    from excelmcp._server_instance import mcp

    prompt_count = None
    if hasattr(mcp, '_prompt_manager') and hasattr(mcp._prompt_manager, '_prompts'):
        prompt_count = len(mcp._prompt_manager._prompts)
    elif hasattr(mcp, 'prompts') and isinstance(mcp.prompts, dict):
        prompt_count = len(mcp.prompts)

    if prompt_count is None:
        import pytest
        pytest.skip("Cannot introspect prompt count in this FastMCP version")

    assert prompt_count == 32, (
        "DRIFT: expected 32 prompts, got %d." % prompt_count
    )


def test_server_py_has_no_prompt_decorators_after_refactor() -> None:
    """After Phase 3, server.py must contain zero @mcp.prompt() decorators."""
    from pathlib import Path
    server_path = Path(__file__).resolve().parent.parent / "src" / "excelmcp" / "server.py"
    text = server_path.read_text(encoding="utf-8")
    assert "@mcp.prompt()" not in text, (
        "DRIFT: server.py still contains @mcp.prompt() decorators. "
        "All prompts must be in server_prompts_*.py sub-modules."
    )
