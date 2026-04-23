"""Sprint O Commit E — unified COM tool tests.

Covers the 4 consolidations reducing 55→50 tools:
  evaluate_formula()  adds live=True path (replaces evaluate_formula_live)
  hydrate_template()  cell_writes is now optional (absorbs open_template_and_recalculate)
  export_as_pdf()     replaces export_sheet_as_pdf + export_workbook_as_pdf
  pivot_table()       replaces create_pivot_table + refresh_pivot_tables + read_pivot_table

All tests are unit-level; no live COM required.
"""
from __future__ import annotations

import hashlib
import sys
from pathlib import Path

import pytest
import openpyxl
from unittest.mock import MagicMock

from fastmcp.exceptions import ToolError

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="COM requires Windows")

from excelmcp._core import NotAllowedError, ValidationError
from excelmcp.server import evaluate_formula
from excelmcp.server_com import (
    export_as_pdf,
    hydrate_template,
    pivot_table,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_xlsx(tmp_path, name="test.xlsx", cell_a1=42):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = cell_a1
    wb.save(str(p))
    return str(p)


def _mock_com(monkeypatch, excel=None, wb=None):
    """Patch _ensure_com_available so no pywin32 import occurs.

    Returns (mock_pythoncom, mock_excel, mock_wb).
    """
    mock_pythoncom = MagicMock()
    mock_excel = excel or MagicMock()
    mock_wb = wb or MagicMock()
    sheet1 = MagicMock()
    sheet1.Name = "Sheet1"
    if wb is None:
        mock_wb.Sheets.Count = 1
        mock_wb.Sheets.side_effect = lambda n: sheet1
        mock_wb.Application = mock_excel
        mock_wb.Names.Count = 0
    mock_win32c = MagicMock()
    mock_win32c.Dispatch.return_value = mock_excel
    mock_excel.Workbooks.Open.return_value = mock_wb

    monkeypatch.setattr(
        "excelmcp._com._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )
    return mock_pythoncom, mock_excel, mock_wb


def _write_fake_pdf(target: str | Path, marker: str) -> None:
    Path(target).write_bytes(f"%PDF-1.7\n{marker}\n".encode("utf-8"))


# ---------------------------------------------------------------------------
# evaluate_formula — live=False (openpyxl path)
# ---------------------------------------------------------------------------

@pytest.mark.unit
def test_evaluate_formula_live_false_reads_cached_value(tmp_path, monkeypatch):
    """live=False uses openpyxl and returns the cached cell value."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path, cell_a1=99)

    result = evaluate_formula(path=path, sheet="Sheet1", address="A1", live=False)
    assert result["cached_value"] == 99


@pytest.mark.unit
def test_evaluate_formula_live_false_default(tmp_path, monkeypatch):
    """live defaults to False — calling without live= uses openpyxl path."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path, cell_a1=7)

    result = evaluate_formula(path=path, sheet="Sheet1", address="A1")
    assert result["cached_value"] == 7


@pytest.mark.unit
def test_evaluate_formula_live_true_no_com_raises(tmp_path, monkeypatch):
    """live=True without EXCEL_ENABLE_COM raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.delenv("EXCEL_ENABLE_COM", raising=False)
    path = _make_xlsx(tmp_path)

    with pytest.raises(ToolError, match="EXCEL_ENABLE_COM"):
        evaluate_formula(path=path, sheet="Sheet1", address="A1", live=True)


@pytest.mark.unit
def test_evaluate_formula_live_true_calls_com(tmp_path, monkeypatch):
    """live=True dispatches to _evaluate_formula_live and returns its result."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    path = _make_xlsx(tmp_path)

    mock_pythoncom, mock_excel, mock_wb = _mock_com(monkeypatch)
    sheet1 = MagicMock()
    sheet1.Name = "Sheet1"
    mock_wb.Sheets.Count = 1
    mock_wb.Sheets.side_effect = lambda n: sheet1
    mock_cell = MagicMock()
    mock_cell.Count = 1
    mock_cell.Formula = "=SUM(A1:A3)"
    mock_cell.Value = 42
    sheet1.Range.return_value = mock_cell

    result = evaluate_formula(path=path, sheet="Sheet1", address="A1", live=True)
    assert result["ok"] is True
    assert result["note"] == "live_com_evaluated"
    assert result["value"] == 42


# ---------------------------------------------------------------------------
# hydrate_template — cell_writes optional
# ---------------------------------------------------------------------------

@pytest.mark.unit
def test_hydrate_template_cell_writes_none_gate_only(tmp_path, monkeypatch):
    """hydrate_template(cell_writes=None) is accepted (fails at COM gate, not param)."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
    monkeypatch.delenv("EXCEL_ENABLE_COM", raising=False)
    path = _make_xlsx(tmp_path, name="tmpl.xlsx")
    out = str(tmp_path / "out.xlsx")

    # cell_writes=None should NOT raise a param error; COM gate will raise
    with pytest.raises(ToolError, match="NotAllowedError"):
        hydrate_template(
            template_path=path, cell_writes=None, output_path=out, confirm=True
        )


@pytest.mark.unit
def test_hydrate_template_cell_writes_empty_list_gate_only(tmp_path, monkeypatch):
    """hydrate_template(cell_writes=[]) is accepted (fails at COM gate, not param)."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
    monkeypatch.delenv("EXCEL_ENABLE_COM", raising=False)
    path = _make_xlsx(tmp_path, name="tmpl2.xlsx")
    out = str(tmp_path / "out2.xlsx")

    with pytest.raises(ToolError, match="NotAllowedError"):
        hydrate_template(
            template_path=path, cell_writes=[], output_path=out, confirm=True
        )


@pytest.mark.unit
def test_hydrate_template_no_write_env_raises(tmp_path, monkeypatch):
    """hydrate_template without EXCEL_ENABLE_WRITE raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.delenv("EXCEL_ENABLE_WRITE", raising=False)
    path = _make_xlsx(tmp_path, name="tmpl3.xlsx")

    with pytest.raises(ToolError, match="NotAllowedError"):
        hydrate_template(
            template_path=path, cell_writes=None,
            output_path=str(tmp_path / "out3.xlsx"), confirm=True
        )


# ---------------------------------------------------------------------------
# export_as_pdf
# ---------------------------------------------------------------------------

@pytest.mark.unit
def test_export_as_pdf_sheet_scope_without_sheet_raises(tmp_path, monkeypatch):
    """scope='sheet' without sheet param raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)
    _mock_com(monkeypatch)

    with pytest.raises(ValidationError, match="sheet is required"):
        export_as_pdf(
            scope="sheet", path=path,
            output_path=str(tmp_path / "out.pdf"),
            sheet=None, confirm=True
        )


@pytest.mark.unit
def test_export_as_pdf_quality_invalid_raises(tmp_path, monkeypatch):
    """quality=2 raises ToolError (must be 0 or 1)."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)
    _mock_com(monkeypatch)

    with pytest.raises(ValidationError, match=r"quality must be 0 \(standard\) or 1 \(minimum\)"):
        export_as_pdf(
            scope="workbook", path=path,
            output_path=str(tmp_path / "out.pdf"),
            quality=2, confirm=True
        )


@pytest.mark.unit
def test_export_as_pdf_no_com_raises(tmp_path, monkeypatch):
    """EXCEL_ENABLE_COM not set raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.delenv("EXCEL_ENABLE_COM", raising=False)
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)

    with pytest.raises(NotAllowedError):
        export_as_pdf(
            scope="workbook", path=path,
            output_path=str(tmp_path / "out.pdf"),
            confirm=True
        )


@pytest.mark.unit
def test_export_as_pdf_not_confirmed_raises(tmp_path, monkeypatch):
    """confirm=False raises ToolError before any COM call."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)

    with pytest.raises(ValidationError):
        export_as_pdf(
            scope="workbook", path=path,
            output_path=str(tmp_path / "out.pdf"),
            confirm=False
        )


@pytest.mark.unit
def test_export_as_pdf_relative_output_path_rejected(tmp_path, monkeypatch):
    """Relative PDF output paths are rejected before COM export starts."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)
    _mock_com(monkeypatch)

    with pytest.raises(ValidationError, match="absolute path"):
        export_as_pdf(
            scope="workbook",
            path=path,
            output_path="relative-output.pdf",
            confirm=True,
        )


@pytest.mark.unit
def test_export_as_pdf_relative_source_path_rejected(tmp_path, monkeypatch):
    """Relative source workbook paths are rejected at the public tool boundary."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    _make_xlsx(tmp_path, name="relative-source.xlsx")
    _mock_com(monkeypatch)

    with pytest.raises(ValidationError, match="path must be an absolute path"):
        export_as_pdf(
            scope="workbook",
            path="relative-source.xlsx",
            output_path=str(tmp_path / "out.pdf"),
            confirm=True,
        )


@pytest.mark.unit
def test_export_as_pdf_invalid_scope_rejected_explicitly(tmp_path, monkeypatch):
    """Invalid scope values are rejected instead of falling through to workbook export."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)
    _mock_com(monkeypatch)

    with pytest.raises(ValidationError, match="scope must be 'sheet' or 'workbook'"):
        export_as_pdf(
            scope="invalid",  # type: ignore[arg-type]
            path=path,
            output_path=str(tmp_path / "out.pdf"),
            confirm=True,
        )


@pytest.mark.unit
def test_export_as_pdf_sheet_scope_returns_verified_payload(tmp_path, monkeypatch):
    """scope='sheet' returns verification evidence after PDF generation."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)
    _, _, mock_wb = _mock_com(monkeypatch)

    sheet1 = MagicMock()
    sheet1.Name = "Sheet1"
    sheet1.ExportAsFixedFormat.side_effect = lambda **kwargs: _write_fake_pdf(
        kwargs["Filename"], "sheet-scope"
    )
    mock_wb.Sheets.Count = 1
    mock_wb.Sheets.side_effect = lambda n: sheet1

    out_pdf = tmp_path / "sheet_scope.pdf"
    result = export_as_pdf(
        scope="sheet",
        path=path,
        sheet="Sheet1",
        output_path=str(out_pdf),
        confirm=True,
    )

    assert result == {
        "ok": True,
        "sheet": "Sheet1",
        "output_path": str(out_pdf.resolve()),
        "quality": 0,
        "sha256": hashlib.sha256(out_pdf.read_bytes()).hexdigest(),
        "size_bytes": out_pdf.stat().st_size,
        "exported_at": result["exported_at"],
        "elapsed_ms": result["elapsed_ms"],
    }
    assert result["exported_at"].endswith("Z")
    assert isinstance(result["elapsed_ms"], int)
    assert result["elapsed_ms"] >= 0


@pytest.mark.unit
def test_export_as_pdf_workbook_scope_returns_verified_payload(tmp_path, monkeypatch):
    """scope='workbook' returns verification evidence after PDF generation."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path, name="book.xlsx")
    _, _, mock_wb = _mock_com(monkeypatch)

    mock_wb.ExportAsFixedFormat.side_effect = lambda **kwargs: _write_fake_pdf(
        kwargs["Filename"], "workbook-scope"
    )

    out_pdf = tmp_path / "workbook_scope.pdf"
    result = export_as_pdf(
        scope="workbook",
        path=path,
        output_path=str(out_pdf),
        confirm=True,
    )

    assert result == {
        "ok": True,
        "output_path": str(out_pdf.resolve()),
        "quality": 0,
        "source": str(Path(path).resolve()),
        "sha256": hashlib.sha256(out_pdf.read_bytes()).hexdigest(),
        "size_bytes": out_pdf.stat().st_size,
        "exported_at": result["exported_at"],
        "elapsed_ms": result["elapsed_ms"],
    }
    assert result["exported_at"].endswith("Z")
    assert isinstance(result["elapsed_ms"], int)
    assert result["elapsed_ms"] >= 0


@pytest.mark.unit
def test_export_as_pdf_zero_byte_output_raises_tool_error(tmp_path, monkeypatch):
    """A zero-byte PDF is rejected even if COM returns successfully."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path, name="zero.xlsx")
    _, _, mock_wb = _mock_com(monkeypatch)

    mock_wb.ExportAsFixedFormat.side_effect = lambda **kwargs: Path(
        kwargs["Filename"]
    ).write_bytes(b"")

    with pytest.raises(ToolError, match="empty"):
        export_as_pdf(
            scope="workbook",
            path=path,
            output_path=str(tmp_path / "zero.pdf"),
            confirm=True,
        )


@pytest.mark.unit
def test_export_as_pdf_clears_preexisting_target_before_export(tmp_path, monkeypatch):
    """Pre-existing PDFs are removed so evidence always reflects a fresh export."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path, name="fresh.xlsx")
    _, _, mock_wb = _mock_com(monkeypatch)

    out_pdf = tmp_path / "fresh.pdf"
    out_pdf.write_bytes(b"%PDF-1.7\nstale\n")

    def _export_after_clear(**kwargs):
        target = Path(kwargs["Filename"])
        assert not target.exists(), "stale target should be removed before export"
        _write_fake_pdf(target, "fresh-bytes")

    mock_wb.ExportAsFixedFormat.side_effect = _export_after_clear

    result = export_as_pdf(
        scope="workbook",
        path=path,
        output_path=str(out_pdf),
        confirm=True,
    )

    assert result["ok"] is True
    assert result["sha256"] == hashlib.sha256(out_pdf.read_bytes()).hexdigest()


@pytest.mark.unit
def test_export_as_pdf_invalid_sheet_preserves_preexisting_target(tmp_path, monkeypatch):
    """Invalid calls must not delete a pre-existing target PDF before validation completes."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path, name="preserve.xlsx")
    _, _, mock_wb = _mock_com(monkeypatch)

    sheet1 = MagicMock()
    sheet1.Name = "Sheet1"
    mock_wb.Sheets.Count = 1
    mock_wb.Sheets.side_effect = lambda n: sheet1

    out_pdf = tmp_path / "preserve.pdf"
    stale_bytes = b"%PDF-1.7\nstale-before-invalid-call\n"
    out_pdf.write_bytes(stale_bytes)

    with pytest.raises(ValidationError, match="not found in workbook"):
        export_as_pdf(
            scope="sheet",
            path=path,
            sheet="MissingSheet",
            output_path=str(out_pdf),
            confirm=True,
        )

    assert out_pdf.read_bytes() == stale_bytes


@pytest.mark.unit
def test_export_as_pdf_workbook_scope_rejects_stray_sheet_preserves_preexisting_target(tmp_path, monkeypatch):
    """Workbook-scope validation failures must preserve a pre-existing target PDF."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path, name="preserve_workbook.xlsx")
    _mock_com(monkeypatch)

    out_pdf = tmp_path / "preserve_workbook.pdf"
    stale_bytes = b"%PDF-1.7\nstale-before-workbook-validation\n"
    out_pdf.write_bytes(stale_bytes)

    with pytest.raises(ValidationError, match="sheet is not allowed when scope='workbook'"):
        export_as_pdf(
            scope="workbook",
            path=path,
            sheet="Sheet1",
            output_path=str(out_pdf),
            confirm=True,
        )

    assert out_pdf.read_bytes() == stale_bytes


@pytest.mark.unit
def test_export_as_pdf_preexisting_target_cannot_be_falsely_attested(tmp_path, monkeypatch):
    """If COM produces no replacement file, a stale pre-existing PDF is not attested."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_EXPORT_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path, name="stale.xlsx")
    _, _, mock_wb = _mock_com(monkeypatch)

    out_pdf = tmp_path / "stale.pdf"
    out_pdf.write_bytes(b"%PDF-1.7\nold-bytes\n")
    mock_wb.ExportAsFixedFormat.side_effect = lambda **kwargs: None

    with pytest.raises(ToolError, match="was not created"):
        export_as_pdf(
            scope="workbook",
            path=path,
            output_path=str(out_pdf),
            confirm=True,
        )


# ---------------------------------------------------------------------------
# pivot_table
# ---------------------------------------------------------------------------

@pytest.mark.unit
def test_pivot_table_create_missing_source_range_raises(tmp_path, monkeypatch):
    """pivot_table(create) without source_address raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)

    with pytest.raises(ToolError, match="source_address"):
        pivot_table(
            operation="create", path=path,
            dest_sheet="Sheet1", dest_address="F1",
            row_fields=["A"], value_fields=["B"],
            confirm=True
        )


@pytest.mark.unit
def test_pivot_table_create_missing_value_fields_raises(tmp_path, monkeypatch):
    """pivot_table(create) without value_fields raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)

    with pytest.raises(ToolError, match="value_fields"):
        pivot_table(
            operation="create", path=path,
            source_address="Sheet1!A1:B5",
            dest_sheet="Sheet1", dest_address="F1",
            row_fields=["A"],
            confirm=True
        )


@pytest.mark.unit
def test_pivot_table_create_no_write_env_raises(tmp_path, monkeypatch):
    """pivot_table(create) without EXCEL_ENABLE_WRITE raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.delenv("EXCEL_ENABLE_WRITE", raising=False)
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    path = _make_xlsx(tmp_path)
    _mock_com(monkeypatch)

    with pytest.raises(ToolError):
        pivot_table(
            operation="create", path=path,
            source_address="Sheet1!A1:B5",
            dest_sheet="Sheet1", dest_address="F1",
            row_fields=["A"], value_fields=["B"],
            confirm=True
        )


@pytest.mark.unit
def test_pivot_table_read_missing_sheet_raises(tmp_path, monkeypatch):
    """pivot_table(read) without sheet raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)

    with pytest.raises(ToolError, match="sheet"):
        pivot_table(operation="read", path=path, pivot_name="PT1")


@pytest.mark.unit
def test_pivot_table_read_missing_pivot_name_raises(tmp_path, monkeypatch):
    """pivot_table(read) without pivot_name raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)

    with pytest.raises(ToolError, match="pivot_name"):
        pivot_table(operation="read", path=path, sheet="Sheet1")


@pytest.mark.unit
def test_pivot_table_refresh_no_confirm_raises(tmp_path, monkeypatch):
    """pivot_table(refresh) with confirm=False raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    path = _make_xlsx(tmp_path)
    _mock_com(monkeypatch)

    with pytest.raises(ToolError):
        pivot_table(operation="refresh", path=path, confirm=False)


@pytest.mark.unit
def test_pivot_table_invalid_operation_raises(tmp_path, monkeypatch):
    """pivot_table with an unknown operation raises ToolError."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    path = _make_xlsx(tmp_path)

    with pytest.raises(ToolError, match="operation must be"):
        pivot_table(operation="delete", path=path)
