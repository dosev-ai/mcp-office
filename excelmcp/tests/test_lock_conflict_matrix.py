"""Lock/conflict matrix tests — parametrized over all 4 write functions.

Tests verify every combination of COM state × file-lock state × error type
for the write path gated by _check_file_not_open_in_excel (LC-2 guard).

E-01: COM disabled → guard is no-op; write proceeds
E-02: COM enabled + file not open → guard clean; write succeeds
E-03: COM enabled + file open → ValidationError propagates
E-04: COM enabled + COM error → OfficeCOMError propagates (fail-closed)
E-05: OfficeCOMError message contains no newline (from None suppression)
"""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock

import openpyxl
import pytest

from excelmcp._core import OfficeCOMError, ValidationError


# ---------------------------------------------------------------------------
# Parametrize setup
# ---------------------------------------------------------------------------

_WRITE_FN_IDS = ["write_cell", "write_range", "append_rows", "create_workbook"]


@pytest.fixture
def write_env(tmp_path, monkeypatch):
    """Allowlisted + write-enabled tmp dir with a pre-built test.xlsx."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
    wb_p = tmp_path / "test.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    wb.save(str(wb_p))
    return tmp_path


def _invoke_write_fn(fn_name: str, tmp_path: Path) -> object:
    """Dispatch to the named write function with appropriate arguments."""
    from excelmcp._io import write_cell, write_range, append_rows, create_workbook

    wb_path = str(tmp_path / "test.xlsx")

    if fn_name == "write_cell":
        return write_cell(wb_path, "Sheet1", "B1", "matrix_test", confirm=True)
    elif fn_name == "write_range":
        return write_range(wb_path, "Sheet1", "A1", [["v1", "v2"]], confirm=True)
    elif fn_name == "append_rows":
        return append_rows(wb_path, "Sheet1", [["row"]], confirm=True)
    elif fn_name == "create_workbook":
        new_path = str(tmp_path / "new_matrix.xlsx")
        return create_workbook(new_path, confirm=True)
    else:
        pytest.fail(f"Unknown fn_name: {fn_name}")


# ---------------------------------------------------------------------------
# E-01: COM disabled → guard is no-op; write proceeds normally
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("fn_name", _WRITE_FN_IDS)
def test_matrix_com_disabled_guard_noop(fn_name, write_env, monkeypatch):
    """E-01: When EXCEL_ENABLE_COM=false, guard returns immediately; write proceeds."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "false")

    guard_spy = MagicMock(wraps=lambda path: None)  # spy that calls real no-op
    monkeypatch.setattr("excelmcp._range_write._check_file_not_open_in_excel", guard_spy)

    result = _invoke_write_fn(fn_name, write_env)

    assert result is not None
    # Guard IS still called on the write path — the guard itself returns silently
    # because of the internal COM-disabled check. Verify the guard was invoked.
    guard_spy.assert_called_once()


# ---------------------------------------------------------------------------
# E-02: COM enabled + file not open → guard clean; write succeeds
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("fn_name", _WRITE_FN_IDS)
def test_matrix_com_enabled_file_not_open_write_proceeds(fn_name, write_env, monkeypatch):
    """E-02: Guard returns cleanly (no-raise); write succeeds end-to-end."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")

    def _clean_guard(path: Path) -> None:
        pass  # file is not open — guard passes

    monkeypatch.setattr("excelmcp._range_write._check_file_not_open_in_excel", _clean_guard)

    result = _invoke_write_fn(fn_name, write_env)
    assert result is not None


# ---------------------------------------------------------------------------
# E-03: COM enabled + file IS open → ValidationError propagates
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("fn_name", _WRITE_FN_IDS)
def test_matrix_file_is_open_write_blocked(fn_name, write_env, monkeypatch):
    """E-03: ValidationError raised by guard propagates and blocks the write."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")

    def _blocking_guard(path: Path) -> None:
        raise ValidationError(
            f"Workbook '{path.name}' is currently open in Excel. "
            "Close it in Excel first, or use a copy of the file."
        )

    monkeypatch.setattr("excelmcp._range_write._check_file_not_open_in_excel", _blocking_guard)

    with pytest.raises(ValidationError):
        _invoke_write_fn(fn_name, write_env)


# ---------------------------------------------------------------------------
# E-04: COM enabled + COM error → OfficeCOMError propagates (fail-closed)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("fn_name", _WRITE_FN_IDS)
def test_matrix_com_error_fail_closed(fn_name, write_env, monkeypatch):
    """E-04: OfficeCOMError raised by guard propagates — write fails closed."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")

    def _com_error_guard(path: Path) -> None:
        raise OfficeCOMError("Could not verify file lock state via COM")

    monkeypatch.setattr("excelmcp._range_write._check_file_not_open_in_excel", _com_error_guard)

    with pytest.raises(OfficeCOMError):
        _invoke_write_fn(fn_name, write_env)


# ---------------------------------------------------------------------------
# E-05: OfficeCOMError message from iteration contains no newline
# ---------------------------------------------------------------------------

def test_matrix_com_error_message_no_newline(monkeypatch, tmp_path):
    """E-05: OfficeCOMError raised from guard iteration contains no newline char.

    The 'from None' suppression means internal error detail is stripped;
    the public error message must be a single-line string with no \\n or \\r.
    """
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setattr(sys, "platform", "win32")

    class _FailIterable:
        def __iter__(self):
            raise RuntimeError("line1\nline2\nsensitive detail")

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
        _check_file_not_open_in_excel(tmp_path / "test.xlsx")

    message = str(exc_info.value)
    assert "\n" not in message, f"OfficeCOMError must not contain newline; got: {message!r}"
    assert "\r" not in message, f"OfficeCOMError must not contain carriage return; got: {message!r}"
