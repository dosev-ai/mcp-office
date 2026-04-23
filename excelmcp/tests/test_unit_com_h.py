"""Unit tests for Sprint H COM tool: evaluate_formula_live (T8).

All tests run without live Excel. COM interaction is mocked via the
mock_com_dispatch and valid_com_env fixtures in conftest.py.
win32com is NEVER imported in this test file.

Test organisation (7 tests total):
    TestEvaluateFormulaLive   -- server_com.evaluate_formula_live (7 tests)

Note: Pivot table tests (TestCreatePivotTable, TestRefreshPivotTables,
TestReadPivotTable) were moved to test_unit_com_h2.py in Sprint P.

Exception contract:
    - (NotAllowedError, ValidationError) are all explicitly converted to ToolError
      in server_com because naked re-raise causes intermittent anyio.ClosedResourceError
      when FastMCP runs the tool in a thread on certain system states.
    - OfficeCOMError / generic Exception -> ToolError at server_com boundary.
"""
from __future__ import annotations

import sys

import pytest
from unittest.mock import MagicMock

from fastmcp.exceptions import ToolError
from excelmcp._core import OfficeCOMError

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="COM requires Windows")


# ===========================================================================
# T8: evaluate_formula_live
# ===========================================================================

class TestEvaluateFormulaLive:
    """Tests for server_com.evaluate_formula_live (T8)."""

    @pytest.mark.unit
    def test_com_gate_blocked(self, monkeypatch, tmp_path):
        """evaluate_formula_live propagates NotAllowedError when EXCEL_ENABLE_COM is absent."""
        monkeypatch.delenv("EXCEL_ENABLE_COM", raising=False)
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import evaluate_formula_live
        with pytest.raises(ToolError):
            evaluate_formula_live(path=str(wb_path), sheet="Sheet1", cell_address="A1")

    @pytest.mark.unit
    def test_invalid_cell_address_raises_validation_error(self, monkeypatch, tmp_path):
        """evaluate_formula_live raises ValidationError for non-A1 cell_address."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import evaluate_formula_live
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            evaluate_formula_live(
                path=str(wb_path), sheet="Sheet1", cell_address="NOTACELL"
            )

    @pytest.mark.unit
    def test_range_address_rejected(self, monkeypatch, tmp_path):
        """evaluate_formula_live raises ValidationError when cell_address is a range A1:B2."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import evaluate_formula_live
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            evaluate_formula_live(
                path=str(wb_path), sheet="Sheet1", cell_address="A1:B2"
            )

    @pytest.mark.unit
    def test_sheet_not_found_raises_validation_error(
        self, valid_com_env, mock_com_dispatch
    ):
        """evaluate_formula_live raises ValidationError when sheet is not in workbook."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        sheet1 = MagicMock()
        sheet1.Name = "Sheet1"
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: sheet1
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import evaluate_formula_live
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            evaluate_formula_live(
                path=str(wb_path), sheet="Ghost", cell_address="A1"
            )

    @pytest.mark.unit
    def test_success_returns_formula_and_value(
        self, valid_com_env, mock_com_dispatch
    ):
        """evaluate_formula_live returns {ok, formula, value, note='live_com_evaluated'}."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        sheet1 = MagicMock()
        sheet1.Name = "Sheet1"
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: sheet1
        mock_range = MagicMock()
        mock_range.Count = 1
        mock_range.Formula = "=SUM(A1:A10)"
        mock_range.Value = 100
        sheet1.Range.return_value = mock_range
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import evaluate_formula_live
        result = evaluate_formula_live(
            path=str(wb_path), sheet="Sheet1", cell_address="b5"
        )

        assert result["ok"] is True
        assert result["formula"] == "=SUM(A1:A10)"
        assert result["value"] == 100
        assert result["cell_address"] == "B5"
        assert result["note"] == "live_com_evaluated"

    @pytest.mark.unit
    def test_range_size_guard_rc6(
        self, valid_com_env, mock_com_dispatch, monkeypatch
    ):
        """evaluate_formula_live raises ValidationError when data_range.Count exceeds limit."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        monkeypatch.setenv("EXCEL_MAX_RANGE_CELLS", "5000")
        sheet1 = MagicMock()
        sheet1.Name = "Sheet1"
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: sheet1
        mock_range = MagicMock()
        mock_range.Count = 5001
        sheet1.Range.return_value = mock_range
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import evaluate_formula_live
        with pytest.raises(ToolError):
            evaluate_formula_live(
                path=str(wb_path), sheet="Sheet1", cell_address="A1"
            )

    @pytest.mark.unit
    def test_office_com_error_wrapped_as_tool_error(
        self, valid_com_env, mock_com_dispatch
    ):
        """evaluate_formula_live wraps OfficeCOMError ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ ToolError."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        mock_excel.Workbooks.Open.side_effect = OfficeCOMError("injected COM failure")
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import evaluate_formula_live
        with pytest.raises(ToolError):
            evaluate_formula_live(
                path=str(wb_path), sheet="Sheet1", cell_address="A1"
            )


# ===========================================================================
# T2: create_pivot_table
# ===========================================================================

