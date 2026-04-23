"""Live formula evaluation via Excel COM.

No MCP imports. No win32com imports at module scope.

Split from _com_template.py — pivot helpers live in _com_pivot.py.

Public API:
    _evaluate_formula_live()  -- return live COM-computed formula/value for a cell
"""
from __future__ import annotations

import logging
import re

from excelmcp._core import (
    ValidationError,
    _check_path,
    _check_range_size,
)

_log = logging.getLogger(__name__)

_SINGLE_CELL_RE = re.compile(r"^\$?[A-Z]{1,3}\$?\d+$", re.IGNORECASE)


def _evaluate_formula_live(path: str, sheet: str, cell_address: str) -> dict:
    """Return the live COM-computed formula and value for a single cell.

    Read-only — does NOT require EXCEL_ENABLE_WRITE.

    Args:
        path:         Absolute path to the workbook within EXCEL_ALLOWLIST_ROOTS.
        sheet:        Worksheet name.
        cell_address: Single cell in A1 notation (e.g. 'B5').

    Returns:
        {
            "ok": True,
            "path": str,
            "sheet": str,
            "cell_address": str,
            "formula": str,
            "value": Any,
            "note": "live_com_evaluated",
        }

    Raises:
        ValidationError:  cell_address invalid or sheet not found.
        OfficeCOMError:   COM failure during evaluation.
        NotAllowedError:  EXCEL_ENABLE_COM not set.
    """
    from excelmcp._com import _close_workbook, _com_excel_app, _open_workbook  # noqa: PLC0415

    if not _SINGLE_CELL_RE.match(cell_address):
        raise ValidationError(
            f"cell_address must be A1-notation single cell (e.g. 'B5'), "
            f"got: {cell_address!r}"
        )

    resolved = _check_path(path)

    with _com_excel_app() as excel:
        wb = _open_workbook(excel, resolved, read_only=True)
        try:
            wb.Activate()
        except Exception:
            pass  # non-fatal in hidden Excel instance
        try:
            sheet_names = [wb.Sheets(i + 1).Name for i in range(wb.Sheets.Count)]
            if sheet not in sheet_names:
                raise ValidationError(
                    f"Sheet {sheet!r} not found in workbook (available: {sheet_names})"
                )
            ws = wb.Sheets(sheet)
            data_range = ws.Range(cell_address.upper())
            _check_range_size(data_range.Count)
            formula = data_range.Formula
            value = data_range.Value
        finally:
            _close_workbook(wb, save=False)
            del wb

    return {
        "ok": True,
        "path": str(resolved),
        "sheet": sheet,
        "cell_address": cell_address.upper(),
        "formula": formula,
        "value": value,
        "note": "live_com_evaluated",
    }
