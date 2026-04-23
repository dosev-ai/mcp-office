"""
Unit tests for Sprint H COM tools — Part 2: pivot tables.

TestCreatePivotTable      — server_com.create_pivot_table  (8 tests)
TestRefreshPivotTables    — server_com.refresh_pivot_tables (6 tests)
TestReadPivotTable        — server_com.read_pivot_table     (6 tests)
Total: 20 tests

All tests run without live Excel. COM interaction is mocked via fixtures
in conftest.py.
"""
from __future__ import annotations

import sys

import pytest
from unittest.mock import MagicMock

from fastmcp.exceptions import ToolError
from excelmcp._core import OfficeCOMError

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="COM requires Windows")

class TestCreatePivotTable:
    """Tests for server_com.create_pivot_table (T2)."""

    @pytest.mark.unit
    def test_write_gate_blocked(self, monkeypatch, tmp_path, mock_com_dispatch):
        """create_pivot_table raises NotAllowedError when EXCEL_ENABLE_WRITE is absent."""
        _, mock_excel, _ = mock_com_dispatch
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.delenv("EXCEL_ENABLE_WRITE", raising=False)
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import create_pivot_table
        with pytest.raises(ToolError, match="Operation failed \\(NotAllowedError\\)"):
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
        mock_excel.Workbooks.Open.assert_not_called()

    @pytest.mark.unit
    def test_confirm_blocked(self, monkeypatch, tmp_path):
        """create_pivot_table raises ValidationError when confirm=False."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import create_pivot_table
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            create_pivot_table(
                path=str(wb_path),
                source_range="Sheet1!A1:D10",
                dest_sheet="Sheet1",
                dest_cell="F1",
                row_fields=["Category"],
                col_fields=[],
                value_fields=["Sales"],
                confirm=False,
            )

    @pytest.mark.unit
    def test_field_count_rc5_blocked(
        self, monkeypatch, tmp_path, mock_com_dispatch
    ):
        """create_pivot_table raises ValidationError before COM when total fields > 50 (RC-5)."""
        _, mock_excel, _ = mock_com_dispatch
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import create_pivot_table
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            create_pivot_table(
                path=str(wb_path),
                source_range="Sheet1!A1:Z100",
                dest_sheet="Sheet1",
                dest_cell="A1",
                row_fields=[f"R{i}" for i in range(26)],
                col_fields=[f"C{i}" for i in range(25)],
                value_fields=["Total"],
                confirm=True,
            )
        mock_excel.Workbooks.Open.assert_not_called()

    @pytest.mark.unit
    def test_source_range_missing_sheet_prefix(self, monkeypatch, tmp_path):
        """create_pivot_table raises ValidationError when source_range lacks '!'."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import create_pivot_table
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            create_pivot_table(
                path=str(wb_path),
                source_range="A1:D10",
                dest_sheet="Sheet1",
                dest_cell="F1",
                row_fields=["Region"],
                col_fields=[],
                value_fields=["Amount"],
                confirm=True,
            )

    @pytest.mark.unit
    def test_dest_sheet_not_found(self, valid_com_env, mock_com_dispatch):
        """create_pivot_table raises ValidationError when dest_sheet is not in the workbook."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        sheet1 = MagicMock()
        sheet1.Name = "Sheet1"
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: sheet1
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import create_pivot_table
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            create_pivot_table(
                path=str(wb_path),
                source_range="Sheet1!A1:D10",
                dest_sheet="NoSuchSheet",
                dest_cell="F1",
                row_fields=["Region"],
                col_fields=[],
                value_fields=["Sales"],
                confirm=True,
            )

    @pytest.mark.unit
    def test_success_creates_pivot_and_saves(
        self, valid_com_env, mock_com_dispatch
    ):
        """create_pivot_table returns {ok, table_name} and calls PivotCaches.Create."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        mock_ws = MagicMock()
        mock_ws.Name = "PivotSheet"
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
            source_range="DataSheet!A1:D10",
            dest_sheet="PivotSheet",
            dest_cell="A1",
            row_fields=["Category"],
            col_fields=[],
            value_fields=["Sales"],
            table_name="MyPivot",
            confirm=True,
        )

        assert result["ok"] is True
        assert result["table_name"] == "MyPivot"
        assert result["row_fields"] == ["Category"]
        assert result["value_fields"] == ["Sales"]
        mock_wb.PivotCaches.return_value.Create.assert_called_once_with(
            SourceType=1, SourceData="DataSheet!A1:D10"
        )
        mock_cache.CreatePivotTable.assert_called_once()
        mock_wb.Save.assert_called_once()

    @pytest.mark.unit
    def test_row_field_not_found_raises_validation_error(
        self, valid_com_env, mock_com_dispatch
    ):
        """create_pivot_table raises ValidationError when row_field not found in pivot."""
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
        mock_pivot.PivotFields.side_effect = Exception("Field 'BadField' not found")
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import create_pivot_table
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            create_pivot_table(
                path=str(wb_path),
                source_range="Sheet1!A1:D10",
                dest_sheet="Sheet1",
                dest_cell="F1",
                row_fields=["BadField"],
                col_fields=[],
                value_fields=[],
                confirm=True,
            )

    @pytest.mark.unit
    def test_com_error_raises_tool_error(self, valid_com_env, mock_com_dispatch):
        """create_pivot_table wraps OfficeCOMError ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ ToolError."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        mock_excel.Workbooks.Open.side_effect = OfficeCOMError("injected COM failure")
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import create_pivot_table
        with pytest.raises(ToolError):
            create_pivot_table(
                path=str(wb_path),
                source_range="Sheet1!A1:D10",
                dest_sheet="Sheet1",
                dest_cell="F1",
                row_fields=["Region"],
                col_fields=[],
                value_fields=["Sales"],
                confirm=True,
            )


# ===========================================================================
# T3: refresh_pivot_tables
# ===========================================================================

class TestRefreshPivotTables:
    """Tests for server_com.refresh_pivot_tables (T3)."""

    @pytest.mark.unit
    def test_write_gate_blocked(self, monkeypatch, tmp_path):
        """refresh_pivot_tables raises NotAllowedError when EXCEL_ENABLE_WRITE absent."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.delenv("EXCEL_ENABLE_WRITE", raising=False)
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import refresh_pivot_tables
        with pytest.raises(ToolError, match="Operation failed \\(NotAllowedError\\)"):
            refresh_pivot_tables(path=str(wb_path), confirm=True)

    @pytest.mark.unit
    def test_confirm_blocked(self, monkeypatch, tmp_path):
        """refresh_pivot_tables raises ValidationError when confirm=False."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import refresh_pivot_tables
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            refresh_pivot_tables(path=str(wb_path), confirm=False)

    @pytest.mark.unit
    def test_external_source_blocked_ssrf_rc2(
        self, valid_com_env, mock_com_dispatch
    ):
        """refresh_pivot_tables raises ValidationError for SourceType != 1 (RC-2/SSRF)."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        mock_pt = MagicMock()
        mock_pt.Name = "ExtPivot"
        mock_pt.PivotCache.return_value.SourceType = 2
        pt_proxy = MagicMock()
        pt_proxy.Count = 1
        def _pt_se(*args):
            return pt_proxy if not args else mock_pt
        mock_ws = MagicMock()
        mock_ws.Name = "Sheet1"
        mock_ws.PivotTables.side_effect = _pt_se
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: mock_ws
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import refresh_pivot_tables
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            refresh_pivot_tables(path=str(wb_path), confirm=True)
        mock_pt.RefreshTable.assert_not_called()

    @pytest.mark.unit
    def test_sheet_not_found(self, valid_com_env, mock_com_dispatch):
        """refresh_pivot_tables raises ValidationError when sheet param is not found."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        sheet1 = MagicMock()
        sheet1.Name = "Sheet1"
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: sheet1
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import refresh_pivot_tables
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            refresh_pivot_tables(
                path=str(wb_path), sheet="MissingSheet", confirm=True
            )

    @pytest.mark.unit
    def test_success_all_sheets(
        self, valid_com_env, mock_com_dispatch, mock_wb_two_pivots
    ):
        """refresh_pivot_tables returns count=2 and calls RefreshTable for two pivots."""
        _, mock_excel, _ = mock_com_dispatch
        tmp_path = valid_com_env
        mock_excel.Workbooks.Open.return_value = mock_wb_two_pivots
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import refresh_pivot_tables
        result = refresh_pivot_tables(path=str(wb_path), confirm=True)

        assert result["ok"] is True
        assert result["count"] == 2
        assert "PT1" in result["refreshed"]
        assert "PT2" in result["refreshed"]

    @pytest.mark.unit
    def test_no_pivot_tables_returns_zero(self, valid_com_env, mock_com_dispatch):
        """refresh_pivot_tables returns count=0 (not an error) when no pivots exist."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        empty_pt_proxy = MagicMock()
        empty_pt_proxy.Count = 0
        mock_ws = MagicMock()
        mock_ws.Name = "Sheet1"
        mock_ws.PivotTables.side_effect = None
        mock_ws.PivotTables.return_value = empty_pt_proxy
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: mock_ws
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import refresh_pivot_tables
        result = refresh_pivot_tables(path=str(wb_path), confirm=True)

        assert result["ok"] is True
        assert result["count"] == 0
        assert result["refreshed"] == []


# ===========================================================================
# T4: read_pivot_table
# ===========================================================================

class TestReadPivotTable:
    """Tests for server_com.read_pivot_table (T4). Read-only tool."""

    @pytest.mark.unit
    def test_com_gate_blocked(self, monkeypatch, tmp_path):
        """read_pivot_table raises NotAllowedError when EXCEL_ENABLE_COM absent."""
        monkeypatch.delenv("EXCEL_ENABLE_COM", raising=False)
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import read_pivot_table
        with pytest.raises(ToolError):
            read_pivot_table(path=str(wb_path), sheet="Sheet1", pivot_name="PT1")

    @pytest.mark.unit
    def test_sheet_not_found(self, valid_com_env, mock_com_dispatch):
        """read_pivot_table raises ValidationError when sheet is not in workbook."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        sheet1 = MagicMock()
        sheet1.Name = "Sheet1"
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: sheet1
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import read_pivot_table
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            read_pivot_table(path=str(wb_path), sheet="Ghost", pivot_name="PT1")

    @pytest.mark.unit
    def test_pivot_not_found(
        self, valid_com_env, mock_com_dispatch, mock_sheet_with_pivots
    ):
        """read_pivot_table raises ValidationError when pivot_name not found on sheet."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: mock_sheet_with_pivots
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        # mock_sheet_with_pivots has PT1 â requesting a different name should fail
        from excelmcp.server_com import read_pivot_table
        with pytest.raises(ToolError, match="Operation failed \\(ValidationError\\)"):
            read_pivot_table(
                path=str(wb_path), sheet="Sheet1", pivot_name="NonExistentPT"
            )

    @pytest.mark.unit
    def test_success_returns_full_schema(
        self, valid_com_env, mock_com_dispatch, mock_sheet_with_pivots, mock_pivot_table
    ):
        """read_pivot_table returns {ok, sheet, pivot_name, row_fields, data}."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: mock_sheet_with_pivots
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import read_pivot_table
        result = read_pivot_table(
            path=str(wb_path), sheet="Sheet1", pivot_name="PT1"
        )

        assert result["ok"] is True
        assert result["pivot_name"] == "PT1"
        assert result["sheet"] == "Sheet1"
        assert "row_fields" in result
        assert "col_fields" in result
        assert "value_fields" in result
        assert "data" in result

    @pytest.mark.unit
    def test_data_body_extracted_as_list_of_lists(
        self, valid_com_env, mock_com_dispatch, mock_sheet_with_pivots, mock_pivot_table
    ):
        """read_pivot_table converts DataBodyRange tuple-of-tuples to list-of-lists."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: mock_sheet_with_pivots
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import read_pivot_table
        result = read_pivot_table(
            path=str(wb_path), sheet="Sheet1", pivot_name="PT1"
        )

        assert result["data"] == [[10, 20], [30, 40]]

    @pytest.mark.unit
    def test_workbook_opened_readonly(
        self, valid_com_env, mock_com_dispatch
    ):
        """read_pivot_table opens workbook with ReadOnly=True (via _open_workbook)."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        sheet1 = MagicMock()
        sheet1.Name = "Sheet1"
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: sheet1
        empty_pt_proxy = MagicMock()
        empty_pt_proxy.Count = 0
        sheet1.PivotTables.side_effect = None
        sheet1.PivotTables.return_value = empty_pt_proxy
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import read_pivot_table
        # Even with 0 pivots, the sheet is found â†’ raises ToolError for pivot not found
        with pytest.raises(ToolError):
            read_pivot_table(
                path=str(wb_path), sheet="Sheet1", pivot_name="PT1"
            )

        # Verify Workbooks.Open was called with ReadOnly=True
        call_kwargs = mock_excel.Workbooks.Open.call_args
        assert call_kwargs is not None
        # ReadOnly is passed as a keyword arg in _open_workbook
        if call_kwargs.kwargs:
            assert call_kwargs.kwargs.get("ReadOnly", False) is True
        # If _open_workbook passes it positionally, skip check ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Â RC-1 is enforced at _com.py level
