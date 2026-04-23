"""Tests for excelmcp.server_com_session — FastMCP active-workbook bridge tools.

9 tests covering: get_active_workbook_context, get_all_open_workbooks,
COM-disabled gates, allowlist filtering, com_error status path, and stdout discipline.
All tests run without a live Excel instance.
"""
from __future__ import annotations

import ast
import inspect
import sys
from contextlib import contextmanager
from unittest.mock import MagicMock, PropertyMock

import pytest



# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _make_wb_mock(full_name: str, name: str = "wb.xlsx", sheets: tuple = ("Sheet1",)) -> MagicMock:
    wb = MagicMock()
    wb.FullName = full_name
    wb.Name = name
    ws_list = [MagicMock() for _ in sheets]
    for m, s in zip(ws_list, sheets):
        m.Name = s
    wb.Worksheets.__iter__ = MagicMock(return_value=iter(ws_list))
    return wb


def _noop_com_gate():
    """Replacement for _ensure_com_gate that does nothing (COM is 'enabled')."""
    pass


# ---------------------------------------------------------------------------
# B-01: get_active_workbook_context — ok status
# ---------------------------------------------------------------------------

def test_get_active_workbook_context_returns_ok_when_com_enabled(monkeypatch, tmp_path):
    """B-01: Returns status='ok' with workbook data when COM is enabled and wb is in allowlist."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr("excelmcp.server_com_session._ensure_com_gate", _noop_com_gate)

    wb_full = str(tmp_path / "active.xlsx")
    mock_wb = _make_wb_mock(wb_full, "active.xlsx", ("Sheet1", "Sheet2"))
    mock_app = MagicMock()
    mock_app.ActiveWorkbook = mock_wb

    @contextmanager
    def _fake_ctx():
        yield mock_app

    monkeypatch.setattr("excelmcp.server_com_session._com_active_app", _fake_ctx)

    from excelmcp.server_com_session import get_active_workbook_context
    result = get_active_workbook_context()

    assert result["status"] == "ok"
    assert result["workbook"]["path"] == wb_full
    assert "Sheet1" in result["workbook"]["sheets"]


# ---------------------------------------------------------------------------
# B-02: get_active_workbook_context — COM disabled raises
# ---------------------------------------------------------------------------

def test_get_active_workbook_context_raises_when_com_disabled(monkeypatch):
    """B-02: Raises NotAllowedError when EXCEL_ENABLE_COM != true."""
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setenv("EXCEL_ENABLE_COM", "false")

    from excelmcp._core import ExcelMCPError
    from excelmcp.server_com_session import get_active_workbook_context

    with pytest.raises(ExcelMCPError):
        get_active_workbook_context()


# ---------------------------------------------------------------------------
# B-03: get_active_workbook_context — no active workbook
# ---------------------------------------------------------------------------

def test_get_active_workbook_context_returns_no_active_workbook_when_none(monkeypatch, tmp_path):
    """B-03: Returns status='no_active_workbook' when ActiveWorkbook is None."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setattr("excelmcp.server_com_session._ensure_com_gate", _noop_com_gate)

    mock_app = MagicMock()
    mock_app.ActiveWorkbook = None

    @contextmanager
    def _fake_ctx():
        yield mock_app

    monkeypatch.setattr("excelmcp.server_com_session._com_active_app", _fake_ctx)

    from excelmcp.server_com_session import get_active_workbook_context
    result = get_active_workbook_context()

    assert result["status"] == "no_active_workbook"
    assert result["workbook"] is None


# ---------------------------------------------------------------------------
# B-04: get_active_workbook_context — outside allowlist
# ---------------------------------------------------------------------------

def test_get_active_workbook_context_returns_outside_allowlist_when_not_in_roots(
    monkeypatch, tmp_path
):
    """B-04: Returns status='outside_allowlist' when active wb is outside EXCEL_ALLOWLIST_ROOTS."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setattr("excelmcp.server_com_session._ensure_com_gate", _noop_com_gate)

    outside_path = r"C:\Other\outside.xlsx"
    mock_wb = _make_wb_mock(outside_path, "outside.xlsx")
    mock_app = MagicMock()
    mock_app.ActiveWorkbook = mock_wb

    @contextmanager
    def _fake_ctx():
        yield mock_app

    monkeypatch.setattr("excelmcp.server_com_session._com_active_app", _fake_ctx)

    from excelmcp.server_com_session import get_active_workbook_context
    result = get_active_workbook_context()

    assert result["status"] == "outside_allowlist"
    assert result["workbook"] is None


# ---------------------------------------------------------------------------
# B-05: get_all_open_workbooks — filtered list
# ---------------------------------------------------------------------------

def test_get_all_open_workbooks_returns_filtered_list(monkeypatch, tmp_path):
    """B-05: Returns only workbooks within EXCEL_ALLOWLIST_ROOTS; outside silently excluded."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setattr("excelmcp.server_com_session._ensure_com_gate", _noop_com_gate)

    in_path1 = str(tmp_path / "wb1.xlsx")
    in_path2 = str(tmp_path / "wb2.xlsx")
    out_path = r"C:\Other\wb_outside.xlsx"

    wb1 = _make_wb_mock(in_path1, "wb1.xlsx")
    wb2 = _make_wb_mock(in_path2, "wb2.xlsx")
    wb_out = _make_wb_mock(out_path, "wb_outside.xlsx")

    mock_app = MagicMock()
    mock_app.Workbooks.__iter__ = MagicMock(return_value=iter([wb1, wb_out, wb2]))

    @contextmanager
    def _fake_ctx():
        yield mock_app

    monkeypatch.setattr("excelmcp.server_com_session._com_active_app", _fake_ctx)

    from excelmcp.server_com_session import get_all_open_workbooks
    result = get_all_open_workbooks()

    assert result["status"] == "ok"
    assert result["count"] == 2
    returned_paths = {wb["path"] for wb in result["workbooks"]}
    assert in_path1 in returned_paths
    assert in_path2 in returned_paths
    assert out_path not in returned_paths


# ---------------------------------------------------------------------------
# B-06: get_all_open_workbooks — empty workbooks list
# ---------------------------------------------------------------------------

def test_get_all_open_workbooks_returns_empty_when_no_workbooks(monkeypatch, tmp_path):
    """B-06: Returns status='ok', count=0 when Excel has no open workbooks."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setattr("excelmcp.server_com_session._ensure_com_gate", _noop_com_gate)

    mock_app = MagicMock()
    mock_app.Workbooks.__iter__ = MagicMock(return_value=iter([]))

    @contextmanager
    def _fake_ctx():
        yield mock_app

    monkeypatch.setattr("excelmcp.server_com_session._com_active_app", _fake_ctx)

    from excelmcp.server_com_session import get_all_open_workbooks
    result = get_all_open_workbooks()

    assert result["status"] == "ok"
    assert result["count"] == 0
    assert result["workbooks"] == []


# ---------------------------------------------------------------------------
# B-07: get_all_open_workbooks — COM disabled raises
# ---------------------------------------------------------------------------

def test_get_all_open_workbooks_raises_when_com_disabled(monkeypatch):
    """B-07: Raises ExcelMCPError (NotAllowedError) when EXCEL_ENABLE_COM != true."""
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setenv("EXCEL_ENABLE_COM", "false")

    from excelmcp._core import ExcelMCPError
    from excelmcp.server_com_session import get_all_open_workbooks

    with pytest.raises(ExcelMCPError):
        get_all_open_workbooks()


# ---------------------------------------------------------------------------
# B-08: No print() calls in server_com_session module
# ---------------------------------------------------------------------------

def test_server_com_session_no_print_calls():
    """B-08: server_com_session.py must contain no print() calls (stdout discipline)."""
    import excelmcp.server_com_session as _mod

    source = inspect.getsource(_mod)
    tree = ast.parse(source)

    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            func = node.func
            if isinstance(func, ast.Name) and func.id == "print":
                pytest.fail(
                    f"print() call found in server_com_session.py at line {node.lineno}. "
                    "All output must go to stderr via logging, never stdout."
                )
            if isinstance(func, ast.Attribute) and func.attr == "print":
                pytest.fail(
                    f"print() call (via attribute) found in server_com_session.py at line {node.lineno}."
                )


# ---------------------------------------------------------------------------
# B-09: get_active_workbook_context — com_error status (H-006 fix)
# ---------------------------------------------------------------------------

def test_get_active_workbook_context_returns_com_error_when_fullname_raises(
    monkeypatch, tmp_path
):
    """B-09: Returns status='com_error' when ActiveWorkbook is not None but
    wb.FullName raises a COM exception (H-006 fix — previously uncovered path).

    B-03 only covers ActiveWorkbook=None; this test covers the sibling branch
    where the workbook object exists but its FullName property is inaccessible.
    """
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setattr("excelmcp.server_com_session._ensure_com_gate", _noop_com_gate)

    # wb object is present but FullName property access raises
    mock_wb = MagicMock()
    type(mock_wb).FullName = PropertyMock(
        side_effect=Exception("COM error: workbook metadata unavailable")
    )
    mock_app = MagicMock()
    mock_app.ActiveWorkbook = mock_wb

    @contextmanager
    def _fake_ctx():
        yield mock_app

    monkeypatch.setattr("excelmcp.server_com_session._com_active_app", _fake_ctx)

    from excelmcp.server_com_session import get_active_workbook_context
    result = get_active_workbook_context()

    assert result["status"] == "com_error"
    assert result["workbook"] is None
    assert "error" in result
