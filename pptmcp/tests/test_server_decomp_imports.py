"""
Import smoke tests for the decomposed pptmcp handler modules.
Verifies that all handler files are importable without side effects,
and that the mcp instance has the expected tool count.
"""
from __future__ import annotations

import asyncio
import builtins
import importlib
import sys
import types

import pytest

pytestmark = pytest.mark.enable_socket


# ---------------------------------------------------------------------------
# Import tests
# ---------------------------------------------------------------------------

def test_read_handlers_module_importable():
    """_server_handlers_read imports without error."""
    import pptmcp._server_handlers_read  # noqa: F401


def test_write_handlers_module_importable():
    """_server_handlers_write imports without error."""
    import pptmcp._server_handlers_write  # noqa: F401


def test_shape_ops_handlers_module_importable():
    """_server_handlers_shape_ops imports without error."""
    import pptmcp._server_handlers_shape_ops  # noqa: F401


def test_export_handlers_module_importable():
    """_server_handlers_export imports without error."""
    import pptmcp._server_handlers_export  # noqa: F401


def test_review_handlers_module_importable():
    """_server_handlers_review imports without error."""
    import pptmcp._server_handlers_review  # noqa: F401


def test_com_handlers_module_importable():
    """_server_handlers_com imports without error (COM unavailable is acceptable)."""
    import pptmcp._server_handlers_com  # noqa: F401


# ---------------------------------------------------------------------------
# server.mcp identity tests
# ---------------------------------------------------------------------------

def test_server_mcp_is_fastmcp_instance():
    """server.mcp must be a FastMCP instance."""
    from fastmcp import FastMCP
    from pptmcp import server
    assert isinstance(server.mcp, FastMCP)


def test_server_mcp_has_minimum_tools():
    """mcp must register at least 37 tools (plus 2 conditional COM tools when available)."""
    from pptmcp import server

    mcp = server.mcp
    tool_manager = getattr(mcp, "_tool_manager", None)
    tools_dict = getattr(tool_manager, "_tools", None)
    if tools_dict is None:
        tools_dict = getattr(mcp, "_tools", None)

    if tools_dict is not None:
        count = len(tools_dict)
    else:
        async def _count():
            tools = await mcp.list_tools()
            return len(tools)
        count = asyncio.run(_count())

    # 37 always-registered; +2 when _COM_AVAILABLE
    assert count >= 37, f"Expected at least 37 always-registered tools (+2 when _COM_AVAILABLE), got {count}"


# ---------------------------------------------------------------------------
# No stdout pollution
# ---------------------------------------------------------------------------

def test_no_stdout_on_handler_imports(capsys):
    """Importing all handler modules must not write anything to stdout."""
    import importlib

    for mod_name in (
        "pptmcp._server_handlers_read",
        "pptmcp._server_handlers_write",
        "pptmcp._server_handlers_shape_ops",
        "pptmcp._server_handlers_export",
        "pptmcp._server_handlers_review",
        "pptmcp._server_handlers_com",
    ):
        importlib.import_module(mod_name)

    captured = capsys.readouterr()
    assert captured.out == "", f"Unexpected stdout on handler import: {captured.out!r}"


# ---------------------------------------------------------------------------
# Per-group tool presence (non-COM tools that are always registered)
# ---------------------------------------------------------------------------

def _get_tool_names():
    from pptmcp import server
    return _get_registered_tool_names(server.mcp)


def _get_registered_tool_names(mcp):
    tool_manager = getattr(mcp, "_tool_manager", None)
    if tool_manager is not None:
        tools = getattr(tool_manager, "_tools", None)
        if tools is not None:
            return set(tools)

    tools = getattr(mcp, "_tools", None)
    if tools is not None:
        return set(tools)

    async def _names():
        tools = await mcp.list_tools()
        return {t.name for t in tools}

    return asyncio.run(_names())


def _import_fresh_com_tools_module(monkeypatch, *, platform: str, block_win32com: bool):
    import pptmcp

    module_name = "pptmcp._server_handlers_com_tools"
    shim_name = "pptmcp._server_handlers_com"
    fake_pcom = types.ModuleType("pptmcp.presentation_com")
    fake_win32com = types.ModuleType("win32com")
    fake_win32com_client = types.ModuleType("win32com.client")
    setattr(fake_win32com, "client", fake_win32com_client)

    monkeypatch.setattr(sys, "platform", platform)
    if block_win32com:
        monkeypatch.delitem(sys.modules, "win32com", raising=False)
        monkeypatch.delitem(sys.modules, "win32com.client", raising=False)
    else:
        monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
        monkeypatch.setitem(sys.modules, "win32com.client", fake_win32com_client)

    monkeypatch.setitem(sys.modules, "pptmcp.presentation_com", fake_pcom)
    monkeypatch.setattr(pptmcp, "presentation_com", fake_pcom, raising=False)

    for name in (module_name, shim_name):
        if name in sys.modules:
            monkeypatch.delitem(sys.modules, name)

    real_import = builtins.__import__

    def fake_import(name, globals=None, locals=None, fromlist=(), level=0):
        if block_win32com and (name == "win32com" or name.startswith("win32com.")):
            raise ImportError("win32com intentionally unavailable for test")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    return importlib.import_module(module_name)


def _import_fresh_com_tools_module_with_backend_failure(monkeypatch, *, platform: str, backend_error: Exception):
    import pptmcp

    module_name = "pptmcp._server_handlers_com_tools"
    shim_name = "pptmcp._server_handlers_com"
    fake_win32com = types.ModuleType("win32com")
    fake_win32com_client = types.ModuleType("win32com.client")
    setattr(fake_win32com, "client", fake_win32com_client)
    real_import_module = importlib.import_module

    monkeypatch.setattr(sys, "platform", platform)
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_win32com_client)
    monkeypatch.delitem(sys.modules, "pptmcp.presentation_com", raising=False)
    monkeypatch.delattr(pptmcp, "presentation_com", raising=False)

    for name in (module_name, shim_name):
        if name in sys.modules:
            monkeypatch.delitem(sys.modules, name)

    def fake_import_module(name, package=None):
        if name == "pptmcp.presentation_com":
            raise backend_error
        return real_import_module(name, package)

    monkeypatch.setattr(importlib, "import_module", fake_import_module)
    return real_import_module(module_name)


_READ_TOOLS = {
    "capabilities", "read_presentation", "get_presentation_metadata", "list_layouts",
    "list_slides", "read_slide", "read_speaker_notes", "list_shapes", "get_shape",
    "extract_tables", "extract_images", "export_slide_as_text",
    "extract_presentation_text", "detect_overlapping_shapes",
}

_WRITE_TOOLS = {
    "add_slide", "edit_text_placeholder", "set_speaker_notes", "replace_slide_text",
    "delete_shape", "insert_image", "reorder_slides", "delete_slide", "save",
    "create_presentation", "manage_hyperlinks", "add_content", "set_format",
    "copy_slide", "add_hyperlink",
}

_REVIEW_TOOLS = {
    "validate_contract", "check_presentation_against_contract",
    "review_slide_export", "produce_evidence_bundle",
}

_COM_ALWAYS_TOOLS = {
    "export_slide", "export_deck", "export_slides_to_stamped_dir",
    "export_changed_slides_only",
}


@pytest.mark.parametrize("tool_name", sorted(_READ_TOOLS))
def test_read_tool_registered(tool_name):
    assert tool_name in _get_tool_names(), f"Read tool missing: {tool_name}"


@pytest.mark.parametrize("tool_name", sorted(_WRITE_TOOLS))
def test_write_tool_registered(tool_name):
    assert tool_name in _get_tool_names(), f"Write tool missing: {tool_name}"


@pytest.mark.parametrize("tool_name", sorted(_REVIEW_TOOLS))
def test_review_tool_registered(tool_name):
    assert tool_name in _get_tool_names(), f"Review tool missing: {tool_name}"


@pytest.mark.parametrize("tool_name", sorted(_COM_ALWAYS_TOOLS))
def test_com_tool_registered(tool_name):
    assert tool_name in _get_tool_names(), f"COM tool missing: {tool_name}"


@pytest.mark.skipif(sys.platform != "win32", reason="COM tools only registered on Windows")
def test_com_conditional_tools_registered_on_windows():
    """On Windows with COM available, run_slide_show and recalculate_charts must register."""
    from pptmcp._server_handlers_com import _COM_AVAILABLE
    if not _COM_AVAILABLE:
        pytest.skip("COM not available even on Windows (win32com not installed)")
    names = _get_tool_names()
    assert "run_slide_show" in names
    assert "recalculate_charts" in names


def test_com_conditional_tools_not_registered_without_pywin32(monkeypatch):
    """Windows imports must still omit COM-only tools when pywin32 is unavailable."""
    from fastmcp import FastMCP

    com_tools = _import_fresh_com_tools_module(
        monkeypatch,
        platform="win32",
        block_win32com=True,
    )

    assert com_tools._COM_AVAILABLE is False
    assert com_tools._pcom is None

    mcp = FastMCP("pptmcp-test")
    com_tools.register_com_tools(mcp)
    names = _get_registered_tool_names(mcp)

    assert "run_slide_show" not in names
    assert "recalculate_charts" not in names


def test_com_backend_import_failure_is_not_silently_downgraded(monkeypatch):
    """Backend import failures must surface instead of being treated as missing pywin32."""
    with pytest.raises(ImportError, match="presentation_com intentionally broken"):
        _import_fresh_com_tools_module_with_backend_failure(
            monkeypatch,
            platform="win32",
            backend_error=ImportError("presentation_com intentionally broken"),
        )


def test_server_export_changed_slides_only_delegates_via_shim_with_synced_state(monkeypatch):
    """server.export_changed_slides_only must keep delegating through the shim after state monkeypatches."""
    from pptmcp import server

    sentinel_pcom = object()
    seen: dict[str, object] = {}
    shim_globals = server.export_changed_slides_only.__globals__
    tools_module = shim_globals["_tools"]

    def fake_export_changed_slides_only(path, base_dir, previous_hashes, width_px, height_px, confirm):
        seen["com_available"] = tools_module._COM_AVAILABLE
        seen["pcom"] = tools_module._pcom
        seen["args"] = (path, base_dir, previous_hashes, width_px, height_px, confirm)
        return {"success": False, "error": "shim path exercised"}

    monkeypatch.setitem(shim_globals, "_COM_AVAILABLE", False)
    monkeypatch.setitem(shim_globals, "_pcom", sentinel_pcom)
    monkeypatch.setattr(tools_module, "_COM_AVAILABLE", True)
    monkeypatch.setattr(tools_module, "_pcom", None)
    monkeypatch.setattr(tools_module, "export_changed_slides_only", fake_export_changed_slides_only)

    result = server.export_changed_slides_only(
        path="deck.pptx",
        base_dir="exports",
        previous_hashes={"1": "abc"},
        width_px=1200,
        height_px=900,
        confirm=True,
    )

    assert result == {"success": False, "error": "shim path exercised"}
    assert seen["com_available"] is False
    assert seen["pcom"] is sentinel_pcom
    assert seen["args"] == ("deck.pptx", "exports", {"1": "abc"}, 1200, 900, True)


def test_shim_register_com_tools_syncs_monkeypatched_state_into_registration(monkeypatch):
    """Shim state must control conditional COM registration even if the tool module had stale state."""
    from fastmcp import FastMCP
    from pptmcp import _server_handlers_com as shim

    sentinel_pcom = object()
    shim_globals = shim.register_com_tools.__globals__
    tools_module = shim_globals["_tools"]

    monkeypatch.setitem(shim_globals, "_COM_AVAILABLE", False)
    monkeypatch.setitem(shim_globals, "_pcom", sentinel_pcom)
    monkeypatch.setattr(tools_module, "_COM_AVAILABLE", True)
    monkeypatch.setattr(tools_module, "_pcom", None)

    mcp = FastMCP("pptmcp-test-sync")
    shim.register_com_tools(mcp)
    names = _get_registered_tool_names(mcp)

    assert tools_module._COM_AVAILABLE is False
    assert tools_module._pcom is sentinel_pcom
    assert "export_changed_slides_only" in names
    assert "run_slide_show" not in names
    assert "recalculate_charts" not in names
