"""
Read-only range and sheet helpers: list_sheets, get_used_range, read_cell, read_range.

No MCP/FastMCP imports. Pure openpyxl + _core.
"""
from __future__ import annotations

from ._core import (
    ValidationError,
    _check_path,
    _check_range_size,
    _get_used_range,
    _load_wb,
    _parse_a1_range,
    _read_range_cells,
)


def list_sheets(path: str) -> list[str]:
    """Return the sheet names in the workbook."""
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    return wb.sheetnames


def get_used_range(path: str, sheet: str) -> dict:
    """Return the used range of a worksheet.

    Returns a dict with keys: min_row, max_row, min_col, max_col, address.
    address is in A1 notation e.g. "A1:C5".
    """
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    if sheet not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb[sheet]
    return _get_used_range(ws)


def read_cell(path: str, sheet: str, address: str) -> dict:
    """Read a single cell.

    Returns {address, value, data_type}.
    """
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    if sheet not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb[sheet]
    cell = ws[address.upper()]
    return {
        "address": address.upper(),
        "value": cell.value,
        "data_type": cell.data_type,
    }


def read_range(path: str, sheet: str, a1_range: str) -> list[list[dict]]:
    """Read a range of cells.

    Returns a 2-D list (rows x cols) of {address, value, data_type} dicts.
    Raises ValidationError if the range exceeds EXCEL_MAX_RANGE_CELLS.
    """
    resolved = _check_path(path)
    min_col, min_row, max_col, max_row = _parse_a1_range(a1_range)
    n_cells = (max_row - min_row + 1) * (max_col - min_col + 1)
    _check_range_size(n_cells)
    wb = _load_wb(resolved)
    if sheet not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb[sheet]
    return _read_range_cells(ws, min_row, min_col, max_row, max_col)
