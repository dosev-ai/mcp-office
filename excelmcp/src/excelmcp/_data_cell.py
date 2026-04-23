"""Cell inspection: metadata, formula evaluation, syntax validation, merged cells.

Internal sub-module. Import via excelmcp._data (public router).
All functions are read-only — no write gate required.
"""
from __future__ import annotations

from ._core import (
    ValidationError,
    _check_path,
    _load_wb,
    openpyxl,
)

__all__ = [
    "get_cell_metadata",
    "evaluate_formula",
    "validate_formula_syntax",
    "list_merged_cells",
]


def get_cell_metadata(path: str, sheet: str, address: str) -> dict:
    """Return formatting and metadata for a single cell.

    Loads without ``data_only`` to preserve all formatting attributes.
    """
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    if sheet not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb[sheet]
    cell = ws[address.upper()]

    font_info = None
    try:
        if cell.font:
            font_info = {
                "name": cell.font.name,
                "size": cell.font.size,
                "bold": cell.font.bold,
                "italic": cell.font.italic,
            }
    except AttributeError:
        font_info = None

    fill_type = None
    try:
        if cell.fill:
            fill_type = cell.fill.patternType
    except AttributeError:
        fill_type = None

    has_comment = cell.comment is not None
    comment_text = None
    try:
        if cell.comment:
            comment_text = cell.comment.text
    except AttributeError:
        comment_text = None

    hyperlink = None
    try:
        if cell.hyperlink:
            hyperlink = str(cell.hyperlink.target)
    except AttributeError:
        hyperlink = None

    return {
        "address": address.upper(),
        "value": cell.value,
        "data_type": cell.data_type,
        "number_format": cell.number_format,
        "font": font_info,
        "fill_type": fill_type,
        "has_comment": has_comment,
        "comment_text": comment_text,
        "hyperlink": hyperlink,
    }


def evaluate_formula(path: str, sheet: str, address: str) -> dict:
    """Return both the formula string and the cached value for *address*.

    Uses two separate ``openpyxl.load_workbook`` calls (not the shared cache) so
    both ``data_only=True`` and ``data_only=False`` modes can be used.
    Returns ``{"address", "formula", "cached_value", "note"}``.
    """
    resolved = _check_path(path)
    wb_check = _load_wb(resolved)
    if sheet not in wb_check.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    addr = address.upper()
    wb_formula = openpyxl.load_workbook(str(resolved), data_only=False)
    wb_data = openpyxl.load_workbook(str(resolved), data_only=True)
    try:
        return {
            "address": addr,
            "formula": wb_formula[sheet][addr].value,
            "cached_value": wb_data[sheet][addr].value,
            "note": "openpyxl_cached_only",
        }
    finally:
        wb_formula.close()
        wb_data.close()


def validate_formula_syntax(formula: str, path: str = "") -> dict:
    """Check whether a formula string *appears* valid for Excel.

    Pure in-memory — does NOT write or read any file on disk.
    If *path* is provided, validates it against EXCEL_ALLOWLIST_ROOTS first.
    Returns ``{"valid": True, "error": None}`` or ``{"valid": False, "error": "<msg>"}``.

    .. note::
        openpyxl stores formula strings without performing full Excel syntax parsing.
        Structurally broken formulas such as ``=SUMX(((`` may therefore return
        ``valid=True``. This is a known openpyxl limitation (peer review finding #1,
        non-critical). Do NOT rely on this function for strict syntax enforcement.
    """
    if path:
        _check_path(path)  # allowlist guard; no disk I/O for the formula itself

    if len(formula) > 8192:
        return {"valid": False, "error": "Formula exceeds maximum length of 8192 characters"}

    if not formula or not formula.startswith("=") or formula.strip() == "=":
        return {"valid": False, "error": "Formula must start with '=' and have content after it"}

    try:
        test_wb = openpyxl.Workbook()
        ws = test_wb.active
        ws["A1"] = formula
        return {"valid": True, "error": None}
    except Exception as exc:
        return {"valid": False, "error": str(exc)}


def list_merged_cells(path: str, sheet: str) -> list:
    """Return a list of merged range address strings (e.g. ['A1:B2', 'C3:D4']).

    Read-only — no write gate required.
    Returns an empty list if the sheet has no merged cell ranges.
    """
    resolved = _check_path(path)
    wb_obj = _load_wb(resolved)
    if sheet not in wb_obj.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb_obj[sheet]
    return [str(m) for m in ws.merged_cells.ranges]
