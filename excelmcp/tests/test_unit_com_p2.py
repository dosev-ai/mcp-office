"""
Unit tests for Sprint G COM Phase 2 tools — Part 2.

TestExportWorkbookAsPdf      — server_com.export_workbook_as_pdf (4 tests)
TestHydrateTemplate          — server_com.hydrate_template (8 tests)
TestOpenTemplateAndRecalculate — server_com.open_template_and_recalculate (3 tests)
Total: 15 tests

All tests run without live Excel. COM interaction is mocked via fixtures
in conftest.py.
"""
from __future__ import annotations

import sys

import pytest
from unittest.mock import MagicMock

from fastmcp.exceptions import ToolError
from excelmcp._core import OfficeCOMError, ValidationError

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="COM requires Windows")

class TestExportWorkbookAsPdf:
    """Tests for server_com.export_workbook_as_pdf (T6).

    EXCEL_ENABLE_WRITE is NOT required ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Â export opens read_only=True.
    """

    @pytest.mark.unit
    def test_confirm_blocked(self, monkeypatch, tmp_path):
        """export_workbook_as_pdf raises ValidationError when confirm=False."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import export_workbook_as_pdf
        with pytest.raises(ValidationError, match="confirm"):
            export_workbook_as_pdf(
                path=str(wb_path),
                output_path=str(tmp_path / "out.pdf"),
                confirm=False,
            )

    @pytest.mark.unit
    def test_success(self, monkeypatch, valid_com_env, mock_com_dispatch):
        """export_workbook_as_pdf returns {ok, output_path} and calls wb.ExportAsFixedFormat(Type=0)."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()
        out_pdf = tmp_path / "full.pdf"

        # Simulate Excel creating the PDF file (as the real COM call would)
        def _create_pdf(**kwargs):
            out_pdf.write_bytes(b"%PDF-1.4 mock")
        mock_wb.ExportAsFixedFormat.side_effect = _create_pdf

        from excelmcp.server_com import export_workbook_as_pdf
        result = export_workbook_as_pdf(
            path=str(wb_path),
            output_path=str(out_pdf),
            confirm=True,
        )

        assert result["ok"] is True
        assert "output_path" in result
        mock_wb.ExportAsFixedFormat.assert_called_once()
        _, kwargs = mock_wb.ExportAsFixedFormat.call_args
        assert kwargs.get("Type") == 0

    @pytest.mark.unit
    def test_outside_export_roots(self, monkeypatch, tmp_path):
        """export_workbook_as_pdf raises ValidationError when output path is outside EXCEL_EXPORT_ROOTS."""
        exports_only = tmp_path / "exports_only"
        exports_only.mkdir()
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(exports_only))
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from excelmcp.server_com import export_workbook_as_pdf
        with pytest.raises(ValidationError):
            export_workbook_as_pdf(
                path=str(wb_path),
                output_path=str(tmp_path / "bad_out.pdf"),  # outside exports_only
                confirm=True,
            )

    @pytest.mark.unit
    def test_com_error_raises_tool_error(self, monkeypatch, valid_com_env, mock_com_dispatch):
        """export_workbook_as_pdf wraps OfficeCOMError ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ ToolError."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        mock_excel.Workbooks.Open.side_effect = OfficeCOMError("injected COM failure")
        wb_path = tmp_path / "test.xlsx"
        wb_path.touch()

        from fastmcp.exceptions import ToolError
        from excelmcp.server_com import export_workbook_as_pdf
        with pytest.raises(ToolError):
            export_workbook_as_pdf(
                path=str(wb_path),
                output_path=str(tmp_path / "out.pdf"),
                confirm=True,
            )


# ---------------------------------------------------------------------------
# Class 5: TestHydrateTemplate (5 tests)
# Tests server_com.hydrate_template (T7a) and
#       server_com.open_template_and_recalculate (T7b).
# ---------------------------------------------------------------------------

class TestHydrateTemplate:
    """Tests for hydrate_template (T7a) and open_template_and_recalculate (T7b).

    Write gate (C2a/C2b): NotAllowedError passes through unmodified.
    Formula injection guard (H1): ValidationError passes through unmodified.
    Gate tests use mock_com_dispatch to verify COM is never reached.
    """

    @pytest.mark.unit
    def test_write_gate_blocked(self, monkeypatch, tmp_path, mock_com_dispatch):
        """hydrate_template raises ToolError wrapping NotAllowedError when EXCEL_ENABLE_WRITE is absent (C2a)."""
        _, mock_excel, _ = mock_com_dispatch
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.delenv("EXCEL_ENABLE_WRITE", raising=False)
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()

        from excelmcp.server_com import hydrate_template
        with pytest.raises(ToolError, match="NotAllowedError"):
            hydrate_template(
                template_path=str(tmpl),
                cell_writes=[{"ref": "A1", "value": "x"}],
                output_path=str(tmp_path / "out.xlsx"),
                confirm=True,
            )
        mock_excel.Workbooks.Open.assert_not_called()

    @pytest.mark.unit
    def test_formula_injection_eq_sum_blocked(self, monkeypatch, valid_com_env, mock_com_dispatch):
        """hydrate_template raises ToolError wrapping ValidationError for '=SUM(...)' formula injection (H1)."""
        _, mock_excel, _ = mock_com_dispatch
        tmp_path = valid_com_env
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()

        from excelmcp.server_com import hydrate_template
        with pytest.raises(ToolError, match="ValidationError"):
            hydrate_template(
                template_path=str(tmpl),
                cell_writes=[{"ref": "A1", "value": "=SUM(A1:A10)"}],
                output_path=str(tmp_path / "out.xlsx"),
                confirm=True,
            )
        mock_excel.Workbooks.Open.assert_not_called()

    @pytest.mark.unit
    def test_space_prefix_injection_blocked(self, monkeypatch, valid_com_env, mock_com_dispatch):
        """hydrate_template raises ToolError wrapping ValidationError when value starts with ' =X' (H1, lstrip guard)."""
        _, mock_excel, _ = mock_com_dispatch
        tmp_path = valid_com_env
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()

        from excelmcp.server_com import hydrate_template
        with pytest.raises(ToolError, match="ValidationError"):
            hydrate_template(
                template_path=str(tmpl),
                cell_writes=[{"ref": "A1", "value": " =X"}],
                output_path=str(tmp_path / "out.xlsx"),
                confirm=True,
            )
        mock_excel.Workbooks.Open.assert_not_called()

    @pytest.mark.unit
    def test_success(self, monkeypatch, valid_com_env, mock_com_dispatch):
        """hydrate_template returns {ok, output_path, writes_applied} and calls SaveAs + CalculateFull."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        sheet1 = MagicMock()
        sheet1.Name = "Sheet1"
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: sheet1
        mock_wb.Names.Count = 0
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()
        out = tmp_path / "output.xlsx"

        from excelmcp.server_com import hydrate_template
        result = hydrate_template(
            template_path=str(tmpl),
            cell_writes=[
                {"ref": "A1", "value": "ProjectX"},
                {"ref": "B2", "value": 42.0},
            ],
            output_path=str(out),
            recalculate=True,
            confirm=True,
        )

        assert result["ok"] is True
        assert result["writes_applied"] == 2
        assert result["recalculated"] is True
        assert "output_path" in result
        mock_wb.Application.CalculateFull.assert_called_once()
        mock_wb.SaveAs.assert_called_once()

    @pytest.mark.unit
    def test_open_template_recalculate_calls_calculate_full(
        self, monkeypatch, valid_com_env, mock_com_dispatch
    ):
        """open_template_and_recalculate returns {ok, output_path, template}
        and calls Application.CalculateFull() + SaveAs (C2b happy path)."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()
        out = tmp_path / "recalculated.xlsx"

        from excelmcp.server_com import open_template_and_recalculate
        result = open_template_and_recalculate(
            template_path=str(tmpl),
            output_path=str(out),
            confirm=True,
        )

        assert result["ok"] is True
        assert "output_path" in result
        assert "template" in result
        mock_wb.Application.CalculateFull.assert_called_once()
        mock_wb.SaveAs.assert_called_once()

    @pytest.mark.unit
    def test_invalid_ref_key_blocked(self, monkeypatch, valid_com_env, mock_com_dispatch):
        """hydrate_template raises ToolError wrapping ValidationError when ref does not match valid address pattern (US-17)."""
        _, mock_excel, _ = mock_com_dispatch
        tmp_path = valid_com_env
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()

        from excelmcp.server_com import hydrate_template
        with pytest.raises(ToolError, match="ValidationError"):
            hydrate_template(
                template_path=str(tmpl),
                cell_writes=[{"ref": "INVALID!!!", "value": "x"}],
                output_path=str(tmp_path / "out.xlsx"),
                confirm=True,
            )
        mock_excel.Workbooks.Open.assert_not_called()

    @pytest.mark.unit
    def test_confirm_blocked(self, monkeypatch, tmp_path):
        """hydrate_template raises ToolError wrapping ValidationError when confirm=False (US-03)."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()

        from excelmcp.server_com import hydrate_template
        with pytest.raises(ToolError, match="ValidationError"):
            hydrate_template(
                template_path=str(tmpl),
                cell_writes=[{"ref": "A1", "value": "x"}],
                output_path=str(tmp_path / "out.xlsx"),
                confirm=False,
            )

    @pytest.mark.unit
    def test_com_error_raises_tool_error(self, monkeypatch, valid_com_env, mock_com_dispatch):
        """hydrate_template wraps OfficeCOMError ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ ToolError (US-18)."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        mock_excel.Workbooks.Open.side_effect = OfficeCOMError("injected COM failure")
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()

        from fastmcp.exceptions import ToolError
        from excelmcp.server_com import hydrate_template
        with pytest.raises(ToolError):
            hydrate_template(
                template_path=str(tmpl),
                cell_writes=[{"ref": "A1", "value": "x"}],
                output_path=str(tmp_path / "out.xlsx"),
                confirm=True,
            )


# ---------------------------------------------------------------------------
# Class 6: TestOpenTemplateAndRecalculate (3 tests)
# ---------------------------------------------------------------------------

class TestOpenTemplateAndRecalculate:
    """Tests for open_template_and_recalculate (T7b).

    Split from TestHydrateTemplate for clarity. Tests cover the return schema,
    confirm gate, and COM error wrapping.
    """

    @pytest.mark.unit
    def test_return_schema_includes_recalculated(
        self, monkeypatch, valid_com_env, mock_com_dispatch
    ):
        """open_template_and_recalculate return dict must include recalculated=True (US-11)."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()
        out = tmp_path / "recalculated.xlsx"

        from excelmcp.server_com import open_template_and_recalculate
        result = open_template_and_recalculate(
            template_path=str(tmpl),
            output_path=str(out),
            confirm=True,
        )

        assert result["recalculated"] is True

    @pytest.mark.unit
    def test_confirm_blocked(self, monkeypatch, tmp_path):
        """open_template_and_recalculate raises ToolError wrapping ValidationError when confirm=False (US-03)."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()

        from excelmcp.server_com import open_template_and_recalculate
        with pytest.raises(ToolError, match="ValidationError"):
            open_template_and_recalculate(
                template_path=str(tmpl),
                output_path=str(tmp_path / "out.xlsx"),
                confirm=False,
            )

    @pytest.mark.unit
    def test_com_error_raises_tool_error(self, monkeypatch, valid_com_env, mock_com_dispatch):
        """open_template_and_recalculate wraps OfficeCOMError ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ ToolError (US-18)."""
        _, mock_excel, mock_wb = mock_com_dispatch
        tmp_path = valid_com_env
        mock_excel.Workbooks.Open.side_effect = OfficeCOMError("injected COM failure")
        tmpl = tmp_path / "template.xlsx"
        tmpl.touch()

        from fastmcp.exceptions import ToolError
        from excelmcp.server_com import open_template_and_recalculate
        with pytest.raises(ToolError):
            open_template_and_recalculate(
                template_path=str(tmpl),
                output_path=str(tmp_path / "out.xlsx"),
                confirm=True,
            )
