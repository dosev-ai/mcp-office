"""Tests for excelmcp._active_wb — active-workbook COM helper.

18 tests across: get_active_workbook_com(), _com_active_app() security
constraints, _check_file_not_open_in_excel(), and _is_in_allowlist().
All tests run without a live Excel instance.
"""
from __future__ import annotations

import sys
from contextlib import contextmanager
from unittest.mock import MagicMock, PropertyMock

import pytest

from excelmcp._core import OfficeCOMError, ValidationError


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_mock_wb(full_name: str, name: str = "test.xlsx", sheets: tuple = ("Sheet1",)) -> MagicMock:
    """Minimal mock workbook COM proxy."""
    wb = MagicMock()
    wb.FullName = full_name
    wb.Name = name
    ws_mocks = [MagicMock() for _ in sheets]
    for m, s in zip(ws_mocks, sheets):
        m.Name = s
    wb.Worksheets.__iter__ = MagicMock(return_value=iter(ws_mocks))
    return wb


def _make_active_app_mocks() -> tuple[MagicMock, MagicMock, MagicMock]:
    """Return (mock_pythoncom, mock_win32c, mock_xl_app)."""
    mock_pythoncom = MagicMock()
    mock_win32c = MagicMock()
    mock_xl_app = MagicMock()
    mock_win32c.GetActiveObject.return_value = mock_xl_app
    return mock_pythoncom, mock_win32c, mock_xl_app


# ---------------------------------------------------------------------------
# A-01 / A-02 / A-03 / A-04: get_active_workbook_com()
# ---------------------------------------------------------------------------

def test_get_active_workbook_returns_dict_on_success(monkeypatch, tmp_path):
    """A-01: Returns dict with name/path/sheets when COM works."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    wb_full_path = str(tmp_path / "test.xlsx")
    mock_wb = _make_mock_wb(wb_full_path, "test.xlsx", ("Sheet1", "Sheet2"))
    mock_app = MagicMock()
    mock_app.ActiveWorkbook = mock_wb

    @contextmanager
    def _fake_ctx():
        yield mock_app

    monkeypatch.setattr("excelmcp._active_wb._com_active_app", _fake_ctx)

    from excelmcp._active_wb import get_active_workbook_com
    result = get_active_workbook_com()

    assert result is not None
    assert result["name"] == "test.xlsx"
    assert result["path"] == wb_full_path
    assert "Sheet1" in result["sheets"]
    assert "Sheet2" in result["sheets"]


def test_get_active_workbook_returns_none_when_no_active_workbook(monkeypatch):
    """A-02: Returns None when ActiveWorkbook is None."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    mock_app = MagicMock()
    mock_app.ActiveWorkbook = None

    @contextmanager
    def _fake_ctx():
        yield mock_app

    monkeypatch.setattr("excelmcp._active_wb._com_active_app", _fake_ctx)

    from excelmcp._active_wb import get_active_workbook_com
    result = get_active_workbook_com()
    assert result is None


def test_get_active_workbook_returns_none_when_com_disabled(monkeypatch):
    """A-03: Returns None without entering COM context when EXCEL_ENABLE_COM != true."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "false")
    entered = []

    @contextmanager
    def _fake_ctx():
        entered.append(True)
        yield MagicMock()

    monkeypatch.setattr("excelmcp._active_wb._com_active_app", _fake_ctx)

    from excelmcp._active_wb import get_active_workbook_com
    result = get_active_workbook_com()

    assert result is None
    assert entered == [], "COM context must NOT be entered when COM is disabled"


def test_get_active_workbook_returns_none_on_com_error(monkeypatch):
    """A-04: Returns None when _com_active_app raises OfficeCOMError."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")

    @contextmanager
    def _failing_ctx():
        raise OfficeCOMError("No running Excel instance found or COM connection failed")
        yield  # noqa: E501

    monkeypatch.setattr("excelmcp._active_wb._com_active_app", _failing_ctx)

    from excelmcp._active_wb import get_active_workbook_com
    result = get_active_workbook_com()
    assert result is None


# ---------------------------------------------------------------------------
# A-05 / A-06 / A-07 / A-08: _com_active_app() security constraints
# ---------------------------------------------------------------------------

def test_com_active_app_does_not_set_automation_security(monkeypatch):
    """A-05: _com_active_app must NEVER set AutomationSecurity on the app object."""
    mock_pythoncom, mock_win32c, mock_xl_app = _make_active_app_mocks()
    monkeypatch.setattr(
        "excelmcp._active_wb._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )

    from excelmcp._active_wb import _com_active_app
    with _com_active_app():
        pass

    # Verify no AutomationSecurity attribute was SET (written) on the xl_app mock
    for c in mock_xl_app.mock_calls:
        assert "AutomationSecurity" not in str(c), (
            f"AutomationSecurity must never be set in _com_active_app; found call: {c}"
        )


def test_com_active_app_does_not_call_workbooks_open(monkeypatch):
    """A-06: _com_active_app must NEVER call Workbooks.Open()."""
    mock_pythoncom, mock_win32c, mock_xl_app = _make_active_app_mocks()
    monkeypatch.setattr(
        "excelmcp._active_wb._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )

    from excelmcp._active_wb import _com_active_app
    with _com_active_app():
        pass

    mock_xl_app.Workbooks.Open.assert_not_called()


def test_com_active_app_does_not_call_quit(monkeypatch):
    """A-07: _com_active_app must NEVER call Quit()."""
    mock_pythoncom, mock_win32c, mock_xl_app = _make_active_app_mocks()
    monkeypatch.setattr(
        "excelmcp._active_wb._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )

    from excelmcp._active_wb import _com_active_app
    with _com_active_app():
        pass

    mock_xl_app.Quit.assert_not_called()


def test_com_active_app_coinitialize_called(monkeypatch):
    """A-08: CoInitialize() is called exactly once when entering _com_active_app."""
    mock_pythoncom, mock_win32c, mock_xl_app = _make_active_app_mocks()
    monkeypatch.setattr(
        "excelmcp._active_wb._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )

    from excelmcp._active_wb import _com_active_app
    with _com_active_app():
        pass

    mock_pythoncom.CoInitialize.assert_called_once()


# ---------------------------------------------------------------------------
# A-09 / A-10 / A-11 / A-12 / A-13: _check_file_not_open_in_excel()
# ---------------------------------------------------------------------------

def test_check_file_not_open_noop_when_com_disabled(monkeypatch, tmp_path):
    """A-09: Returns silently when EXCEL_ENABLE_COM != true (no exception raised)."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "false")

    from excelmcp._active_wb import _check_file_not_open_in_excel
    # Must not raise, regardless of file path
    _check_file_not_open_in_excel(tmp_path / "anything.xlsx")


def test_check_file_not_open_passes_when_not_in_excel(monkeypatch, tmp_path):
    """A-10: No exception when the target file is NOT open in Excel."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setattr(sys, "platform", "win32")

    target = tmp_path / "target.xlsx"
    other = tmp_path / "other.xlsx"

    mock_wb = MagicMock()
    mock_wb.FullName = str(other)
    mock_pythoncom = MagicMock()
    mock_win32c = MagicMock()
    mock_xl_app = MagicMock()
    mock_xl_app.Workbooks = [mock_wb]  # real list — reliable iteration
    mock_win32c.GetActiveObject.return_value = mock_xl_app
    monkeypatch.setattr(
        "excelmcp._active_wb._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )

    from excelmcp._active_wb import _check_file_not_open_in_excel
    # Must not raise — different file is open, not target
    _check_file_not_open_in_excel(target)


def test_check_file_not_open_raises_validation_error_when_open(monkeypatch, tmp_path):
    """A-11: Raises ValidationError when the target file IS open in Excel."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setattr(sys, "platform", "win32")

    target = tmp_path / "target.xlsx"
    target.touch()  # must exist so Path.resolve() is deterministic

    # Use the *resolved* str so the path comparison inside the function succeeds.
    # The function does: Path(wb.FullName).resolve().lower() == resolved.resolve().lower()
    mock_wb = MagicMock()
    mock_wb.FullName = str(target.resolve())
    mock_pythoncom = MagicMock()
    mock_win32c = MagicMock()
    mock_xl_app = MagicMock()
    mock_xl_app.Workbooks = [mock_wb]  # real list — reliable iteration
    mock_win32c.GetActiveObject.return_value = mock_xl_app
    monkeypatch.setattr(
        "excelmcp._active_wb._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )

    from excelmcp._active_wb import _check_file_not_open_in_excel
    with pytest.raises(ValidationError):
        _check_file_not_open_in_excel(target)


def test_check_file_not_open_raises_office_com_error_on_iteration_error(monkeypatch, tmp_path):
    """A-12: Raises OfficeCOMError when iterating Workbooks raises RuntimeError."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setattr(sys, "platform", "win32")

    class _FailIterable:
        """Iterable that raises on __iter__ to simulate COM iteration failure."""
        def __iter__(self):
            raise RuntimeError("COM iteration failure")

    mock_pythoncom = MagicMock()
    mock_win32c = MagicMock()
    mock_xl_app = MagicMock()
    mock_xl_app.Workbooks = _FailIterable()  # raises on for-loop entry
    mock_win32c.GetActiveObject.return_value = mock_xl_app
    monkeypatch.setattr(
        "excelmcp._active_wb._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )

    from excelmcp._active_wb import _check_file_not_open_in_excel
    with pytest.raises(OfficeCOMError):
        _check_file_not_open_in_excel(tmp_path / "any.xlsx")


def test_check_file_not_open_com_error_suppresses_detail(monkeypatch, tmp_path):
    """A-13: OfficeCOMError raised from iteration has __cause__ == None (from None suppression)."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setattr(sys, "platform", "win32")

    class _FailIterable:
        def __iter__(self):
            raise RuntimeError("sensitive detail that must not leak")

    mock_pythoncom = MagicMock()
    mock_win32c = MagicMock()
    mock_xl_app = MagicMock()
    mock_xl_app.Workbooks = _FailIterable()  # raises on for-loop entry
    mock_win32c.GetActiveObject.return_value = mock_xl_app
    monkeypatch.setattr(
        "excelmcp._active_wb._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )

    from excelmcp._active_wb import _check_file_not_open_in_excel
    with pytest.raises(OfficeCOMError) as exc_info:
        _check_file_not_open_in_excel(tmp_path / "any.xlsx")

    assert exc_info.value.__cause__ is None, (
        "OfficeCOMError from iteration must suppress cause via 'raise ... from None'"
    )


# ---------------------------------------------------------------------------
# A-14: get_active_workbook_com() — no allowlist filtering at helper level
# ---------------------------------------------------------------------------

def test_get_active_workbook_returns_data_regardless_of_allowlist(monkeypatch, tmp_path):
    """A-14: get_active_workbook_com() returns data even for out-of-allowlist workbooks.

    Filtering by allowlist is the server layer's responsibility (server_com_session.py),
    NOT the raw helper function's. This test documents that design boundary.
    """
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    # Allowlist is tmp_path, but the active workbook is outside it
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    outside_path = r"C:\Other\wb.xlsx"
    mock_app = MagicMock()
    outside_wb = _make_mock_wb(outside_path, "wb.xlsx", ("Sheet1",))
    mock_app.ActiveWorkbook = outside_wb

    @contextmanager
    def _fake_ctx():
        yield mock_app

    monkeypatch.setattr("excelmcp._active_wb._com_active_app", _fake_ctx)

    from excelmcp._active_wb import get_active_workbook_com
    result = get_active_workbook_com()
    # Helper returns the dict — allowlist check is the server layer's job
    assert result is not None
    assert result["path"] == outside_path


# ---------------------------------------------------------------------------
# A-15 / A-16 / A-17: _is_in_allowlist()
# ---------------------------------------------------------------------------

def test_is_in_allowlist_false_when_roots_empty(monkeypatch, tmp_path):
    """A-15: Returns False when EXCEL_ALLOWLIST_ROOTS is unset."""
    monkeypatch.delenv("EXCEL_ALLOWLIST_ROOTS", raising=False)

    from excelmcp._active_wb import _is_in_allowlist
    assert _is_in_allowlist(str(tmp_path / "test.xlsx")) is False


def test_is_in_allowlist_true_for_file_in_root(monkeypatch, tmp_path):
    """A-16: Returns True for a file that is inside the configured root."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))

    from excelmcp._active_wb import _is_in_allowlist
    assert _is_in_allowlist(str(tmp_path / "test.xlsx")) is True


def test_is_in_allowlist_false_for_file_outside_root(monkeypatch, tmp_path):
    """A-17: Returns False for a file that is outside the configured root."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))

    from excelmcp._active_wb import _is_in_allowlist
    assert _is_in_allowlist(r"C:\Other\wb.xlsx") is False


# ---------------------------------------------------------------------------
# A-18: _check_file_not_open_in_excel() — H-002 inner FullName raise
# ---------------------------------------------------------------------------

def test_check_file_not_open_raises_when_fullname_raises_inside_loop(monkeypatch, tmp_path):
    """A-18 (H-002): OfficeCOMError is raised (fail-closed) when wb.FullName raises
    for a specific workbook inside the Workbooks iteration loop.

    This is distinct from A-12 (where Workbooks.__iter__ itself raises).
    The inner 'except Exception as exc' must fail-closed with OfficeCOMError,
    NOT silently skip the problematic workbook and continue iteration.
    """
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setattr(sys, "platform", "win32")

    # A workbook whose FullName property raises on access (simulates COM data error)
    mock_bad_wb = MagicMock()
    type(mock_bad_wb).FullName = PropertyMock(
        side_effect=Exception("COM read error: FullName property unavailable")
    )

    mock_pythoncom = MagicMock()
    mock_win32c = MagicMock()
    mock_xl_app = MagicMock()
    mock_xl_app.Workbooks = [mock_bad_wb]  # real list — iteration works; FullName raises
    mock_win32c.GetActiveObject.return_value = mock_xl_app
    monkeypatch.setattr(
        "excelmcp._active_wb._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )

    from excelmcp._active_wb import _check_file_not_open_in_excel
    with pytest.raises(OfficeCOMError):
        _check_file_not_open_in_excel(tmp_path / "any.xlsx")
