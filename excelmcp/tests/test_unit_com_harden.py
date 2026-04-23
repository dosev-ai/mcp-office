"""Unit tests for COM-Harden-G changes.

Covers three bug fixes introduced in the uncommitted COM-Harden-G work:

Bug 2A — xl_app.Ready polling loop in _excel_startup
Bug 2B — _classify_com_error pure classification function
Bug 1  — wb.Activate() calls in _com_template.py helpers

All tests run without live Excel. COM is fully mocked.
"""
from __future__ import annotations

import sys
from unittest.mock import MagicMock, PropertyMock, call, patch

import pytest

from excelmcp._core import OfficeCOMError

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="COM requires Windows")


# ---------------------------------------------------------------------------
# Helpers shared by Bug 2A tests
# ---------------------------------------------------------------------------

def _make_com_mocks(monkeypatch):
    """Return (mock_pythoncom, mock_win32c) with EXCEL_ENABLE_COM patched."""
    mock_pythoncom = MagicMock()
    mock_win32c = MagicMock()
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setattr(
        "excelmcp._com._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )
    return mock_pythoncom, mock_win32c


# ===========================================================================
# Bug 2A — xl_app.Ready polling loop
# ===========================================================================

class TestExcelReadyPolling:
    """Tests for the xl_app.Ready polling loop added to _excel_startup (Bug 2A).

    `time` is imported as `_time` inside `_excel_startup`, so we patch the
    module-level `time.monotonic` and `time.sleep` builtins rather than an
    attribute on `excelmcp._com`.

    Note: `_com_excel_app` always calls `excel.Quit()` in its finally block on
    normal exit (teardown).  Tests therefore check Quit call COUNT rather than
    assert_not_called().
    """

    @pytest.mark.unit
    def test_excel_ready_immediately_no_retries(self, monkeypatch):
        """When excel.Ready is True on the first check the loop breaks immediately.

        Quit should be called exactly once — by the context manager teardown,
        not by the timeout path.
        """
        mock_pythoncom, mock_win32c = _make_com_mocks(monkeypatch)
        mock_excel = MagicMock()
        mock_excel.Ready = True
        mock_win32c.Dispatch.return_value = mock_excel

        from excelmcp._com import _com_excel_app

        with _com_excel_app():
            pass

        # Teardown calls Quit once; the timeout/failure path would call it
        # before raising OfficeCOMError, so exactly 1 call means no retry.
        assert mock_excel.Quit.call_count == 1

    @pytest.mark.unit
    def test_excel_ready_after_two_retries(self, monkeypatch):
        """Ready=False twice then True: startup succeeds without OfficeCOMError.

        time.sleep is patched to a no-op; time.monotonic returns values that
        stay well below the 15 s deadline so the loop keeps iterating.
        """
        mock_pythoncom, mock_win32c = _make_com_mocks(monkeypatch)
        mock_excel = MagicMock()
        # PropertyMock returns values in sequence: False, False, True
        type(mock_excel).Ready = PropertyMock(side_effect=[False, False, True])
        mock_win32c.Dispatch.return_value = mock_excel

        with patch("time.monotonic", side_effect=[0.0, 0.1, 0.2, 0.3, 0.4]):
            with patch("time.sleep"):
                from excelmcp._com import _com_excel_app

                with _com_excel_app():
                    pass  # must not raise

        # Normal teardown: exactly 1 Quit call from context manager exit
        assert mock_excel.Quit.call_count == 1

    @pytest.mark.unit
    def test_excel_never_ready_raises_office_com_error(self, monkeypatch):
        """When excel.Ready never becomes True the polling loop raises OfficeCOMError.

        time.monotonic is scripted to cross the 15 s deadline quickly.
        time.sleep is patched to a no-op so the test runs instantly.
        """
        mock_pythoncom, mock_win32c = _make_com_mocks(monkeypatch)
        mock_excel = MagicMock()
        mock_excel.Ready = False  # always False
        mock_win32c.Dispatch.return_value = mock_excel

        # First call returns 0.0 (sets _ready_deadline = 15.0),
        # second call returns 20.0 (>= deadline — triggers timeout).
        with patch("time.monotonic", side_effect=[0.0, 20.0]):
            with patch("time.sleep"):
                from excelmcp._com import _com_excel_app

                with pytest.raises(OfficeCOMError) as exc_info:
                    with _com_excel_app():
                        pass

        assert "not ready after 15 s" in str(exc_info.value)
        assert "xl_app.Ready" in str(exc_info.value)

    @pytest.mark.unit
    def test_excel_ready_attribute_error_treated_as_not_ready_then_succeeds(
        self, monkeypatch
    ):
        """AttributeError from excel.Ready is silenced; polling continues.

        After the error, when Ready becomes True, startup must succeed.
        """
        mock_pythoncom, mock_win32c = _make_com_mocks(monkeypatch)
        mock_excel = MagicMock()
        # First access raises AttributeError, second returns True
        type(mock_excel).Ready = PropertyMock(
            side_effect=[AttributeError("no Ready yet"), True]
        )
        mock_win32c.Dispatch.return_value = mock_excel

        with patch("time.monotonic", return_value=0.0):
            with patch("time.sleep"):
                from excelmcp._com import _com_excel_app

                with _com_excel_app():
                    pass  # must not raise


# ===========================================================================
# Bug 2B — _classify_com_error
# ===========================================================================

class TestClassifyComError:
    """Tests for _classify_com_error pure classification function (Bug 2B).

    This is a pure function — no COM mocking needed.
    """

    @pytest.mark.unit
    def test_classify_coinitialize_substring(self):
        """OfficeCOMError mentioning 'CoInitialize' maps to installation hint."""
        from excelmcp.server_com_recalc import _classify_com_error

        result = _classify_com_error(
            OfficeCOMError("Excel COM startup failed during CoInitialize()")
        )
        assert "is Excel installed" in result

    @pytest.mark.unit
    def test_classify_not_available_creating_substring(self):
        """OfficeCOMError mentioning 'creating Excel.Application' maps to dispatch hint."""
        from excelmcp.server_com_recalc import _classify_com_error

        result = _classify_com_error(
            OfficeCOMError("Excel COM startup failed creating Excel.Application")
        )
        # The message must mention dispatch or installation
        assert "dispatched" in result or "installed" in result

    @pytest.mark.unit
    def test_classify_not_ready_substring(self):
        """OfficeCOMError mentioning 'not ready after 15 s' maps to retry hint."""
        from excelmcp.server_com_recalc import _classify_com_error

        result = _classify_com_error(
            OfficeCOMError(
                "Excel COM startup: Excel.Application not ready after 15 s — timed out"
            )
        )
        assert "retry" in result.lower() or "ready" in result.lower()

    @pytest.mark.unit
    def test_classify_busy_initialising_substring(self):
        """OfficeCOMError mentioning 'Excel is busy' maps to busy/retry hint."""
        from excelmcp.server_com_recalc import _classify_com_error

        result = _classify_com_error(OfficeCOMError("Excel is busy doing something"))
        assert "busy" in result.lower() or "retry" in result.lower()

    @pytest.mark.unit
    def test_classify_unknown_returns_fallback(self):
        """Unknown OfficeCOMError message returns a generic fallback string."""
        from excelmcp.server_com_recalc import _classify_com_error

        result = _classify_com_error(OfficeCOMError("some completely unknown error xyz"))
        assert result.startswith("Excel COM error")


# ===========================================================================
# Bug 1 — wb.Activate() in _com_template.py
# ===========================================================================

class TestWbActivate:
    """Tests for wb.Activate() calls added to _com_template helpers (Bug 1).

    These tests use mock_com_dispatch (from conftest.py) so no real COM fires.
    """

    @pytest.mark.unit
    def test_wb_activate_called_on_write_open_create_pivot(
        self, valid_com_env, mock_com_dispatch
    ):
        """_create_pivot_table calls wb.Activate() when opening for write.

        Activate is called once and must not raise.
        """
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env

        mock_ws = MagicMock()
        mock_ws.Name = "Sheet1"
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: mock_ws
        mock_cache = MagicMock()
        mock_wb.PivotCaches.return_value.Create.return_value = mock_cache
        mock_pivot = MagicMock()
        mock_cache.CreatePivotTable.return_value = mock_pivot

        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import create_pivot_table
        result = create_pivot_table(
            path=str(wb_path),
            source_range="Sheet1!A1:D10",
            dest_sheet="Sheet1",
            dest_cell="F1",
            row_fields=["Category"],
            col_fields=[],
            value_fields=["Sales"],
            confirm=True,
        )

        assert result["ok"] is True
        mock_wb.Activate.assert_called_once()

    @pytest.mark.unit
    def test_wb_activate_failure_write_raises_tool_error(
        self, valid_com_env, mock_com_dispatch
    ):
        """When wb.Activate() raises in write-mode (_create_pivot_table), ToolError is raised.

        _com_template raises OfficeCOMError which is wrapped to ToolError at the server boundary.
        """
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env

        mock_wb.Activate.side_effect = Exception("COM: workbook not accessible")

        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from fastmcp.exceptions import ToolError
        from excelmcp.server_com import create_pivot_table
        with pytest.raises(ToolError):
            create_pivot_table(
                path=str(wb_path),
                source_range="Sheet1!A1:D10",
                dest_sheet="Sheet1",
                dest_cell="F1",
                row_fields=["Category"],
                col_fields=[],
                value_fields=["Sales"],
                confirm=True,
            )

    @pytest.mark.unit
    def test_wb_activate_called_on_refresh_pivot_tables(
        self, valid_com_env, mock_com_dispatch
    ):
        """_refresh_pivot_tables calls wb.Activate() once (write-mode open)."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env

        mock_ws = MagicMock()
        mock_ws.Name = "Sheet1"
        pt_proxy = MagicMock()
        pt_proxy.Count = 0
        mock_ws.PivotTables.side_effect = lambda *a: pt_proxy
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: mock_ws

        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import refresh_pivot_tables
        result = refresh_pivot_tables(path=str(wb_path), confirm=True)

        assert result["ok"] is True
        mock_wb.Activate.assert_called_once()

    @pytest.mark.unit
    def test_wb_activate_failure_on_refresh_raises_tool_error(
        self, valid_com_env, mock_com_dispatch
    ):
        """When wb.Activate() raises in _refresh_pivot_tables, ToolError is raised."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env

        mock_wb.Activate.side_effect = Exception("COM: activate failed")

        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from fastmcp.exceptions import ToolError
        from excelmcp.server_com import refresh_pivot_tables
        with pytest.raises(ToolError):
            refresh_pivot_tables(path=str(wb_path), confirm=True)

    @pytest.mark.unit
    def test_wb_activate_failure_read_silent_evaluate_formula_live(
        self, valid_com_env, mock_com_dispatch
    ):
        """When wb.Activate() raises in read-mode (_evaluate_formula_live), no exception is raised.

        The Activate failure is silenced for read-mode operations in a hidden Excel instance.
        """
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env

        mock_wb.Activate.side_effect = Exception("COM: can't activate hidden workbook")

        mock_ws = MagicMock()
        mock_ws.Name = "Sheet1"
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: mock_ws
        mock_range = MagicMock()
        mock_range.Count = 1
        mock_range.Formula = "=A1+1"
        mock_range.Value = 42
        mock_ws.Range.return_value = mock_range

        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import evaluate_formula_live
        # Must NOT raise despite Activate() failure
        result = evaluate_formula_live(
            path=str(wb_path), sheet="Sheet1", cell_address="B1"
        )
        assert result["ok"] is True
