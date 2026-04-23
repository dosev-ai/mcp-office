"""Batch-operation MCP tools for excelmcp.

Extracted from server.py. Tools cover bulk sheet copy, bulk row height setting,
bulk column width setting, bulk comment addition, and formula range fill.

Import side-effect: importing this module registers all tools on the shared
mcp instance from excelmcp._server_instance.
"""
from __future__ import annotations

from fastmcp.exceptions import ToolError

from excelmcp import workbook_openpyxl as wb
from excelmcp._server_instance import mcp
from excelmcp.workbook_openpyxl import ExcelMCPError

# ---------------------------------------------------------------------------
# Phase 2.7 Sprint C — bulk operations + security backport
# ---------------------------------------------------------------------------


@mcp.tool()
def bulk_copy_sheet(
    path: str,
    source_sheet: str,
    new_names: list[str],
    insert_after: str | None = None,
    confirm: bool = False,
) -> dict:
    """Copy *source_sheet* to multiple new sheets in a single call.

    new_names: list of new sheet names to create (no duplicates, max 31 chars each,
    no /\\*?[] chars).
    insert_after: if provided, new sheets are placed immediately after this sheet.
    Note: openpyxl copy_worksheet does NOT copy charts, images, or print settings.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.bulk_copy_sheet(
            path, source_sheet, new_names,
            insert_after=insert_after, confirm=confirm,
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def set_row_heights(
    path: str,
    sheet: str,
    rows: list[int] | None = None,
    row_range: str | None = None,
    height: float | None = None,
    row_specs: list[dict] | None = None,
    confirm: bool = False,
) -> dict:
    """Set heights for multiple rows in a single call.

    Mode 1 (uniform): provide rows or row_range (e.g. '1:20') plus height.
    Mode 2 (per-row): provide row_specs list of {"row": int, "height": float}.
    Valid height: 0.0–409.0.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.set_row_heights(
            path, sheet, rows=rows, row_range=row_range,
            height=height, row_specs=row_specs, confirm=confirm,
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def set_column_widths(
    path: str,
    sheet: str,
    columns: list[str] | None = None,
    col_range: str | None = None,
    width: float | None = None,
    autofit: bool = False,
    autofit_padding: float = 2.0,
    max_width: float = 60.0,
    confirm: bool = False,
) -> dict:
    """Set widths for multiple columns in a single call.

    Provide exactly one of columns (list of letters) or col_range (e.g. 'A:E').
    Provide exactly one of width (0-255) or autofit=True.
    autofit approximates width from cell content; max_width caps the result.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.set_column_widths(
            path, sheet, columns=columns, col_range=col_range,
            width=width, autofit=autofit, autofit_padding=autofit_padding,
            max_width=max_width, confirm=confirm,
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# Sprint D Batch 1 — T7: bulk_add_comments, T5: fill_formula_range
# ---------------------------------------------------------------------------

@mcp.tool()
def bulk_add_comments(
    path: str,
    sheet: str,
    comments: list[dict],
    overwrite: bool = True,
    confirm: bool = False,
) -> dict:
    """Add multiple cell comments in one call. Each dict: address, text, author (optional). Returns comments_added count."""
    try:
        return wb.bulk_add_comments(
            path, sheet, comments=comments, overwrite=overwrite, confirm=confirm
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def fill_formula_range(
    path: str,
    sheet: str,
    formula: str,
    start_address: str,
    end_address: str,
    confirm: bool = False,
) -> dict:
    """Fill start_address:end_address with formula, auto-relativizing references. Formula must start with '='. Returns filled_cells count.

    Requires EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    try:
        return wb.fill_formula_range(
            path, sheet, anchor_cell=start_address, end_cell=end_address,
            formula=formula, confirm=confirm,
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def bulk_range_write(
    path: str,
    sheet: str,
    writes: list[dict],
    confirm: bool = False,
) -> dict:
    """Write multiple non-contiguous cells in a single tool call.

    Writes a list of cell updates to the same sheet in one open/save cycle —
    more efficient than calling ``range_io`` or ``cell`` once per cell.

    Parameters
    ----------
    path : str
        Absolute path to the workbook. Must be within EXCEL_ALLOWLIST_ROOTS.
    sheet : str
        Sheet name.
    writes : list[dict]
        Each item: ``{"address": "A1", "value": <scalar>}``.
        Strings starting with ``=``, ``+``, ``-``, or ``@`` are rejected
        (formula injection guard).
    confirm : bool, optional
        Must be True to authorise the write. Requires ``EXCEL_ENABLE_WRITE=true``.

    Returns
    -------
    dict
        ``{"ok": true, "cells_written": N}``
    """
    try:
        return wb.bulk_write_cells(path, sheet, writes, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e
