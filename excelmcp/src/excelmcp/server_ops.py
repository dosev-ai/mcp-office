"""Structure-operations MCP tools for excelmcp.

Extracted from server.py. Tools cover merged cells, freeze panes, auto filter,
row/column insert-delete, sheet/workbook protection, print area, cell comments,
sheet properties, data validation, formula syntax checking, range copy,
table creation, column visibility, workbook calc mode, and hyperlinks.

Import side-effect: importing this module registers all tools on the shared
mcp instance from excelmcp._server_instance.
"""
from __future__ import annotations

from typing import Literal

from fastmcp.exceptions import ToolError

from excelmcp import workbook_openpyxl as wb
from excelmcp._server_instance import mcp
from excelmcp.workbook_openpyxl import ExcelMCPError

# ---------------------------------------------------------------------------
# Phase 2.2 tools  (require EXCEL_ENABLE_WRITE=true + confirm=True)
# ---------------------------------------------------------------------------

@mcp.tool()
def merge(
    operation: Literal["list", "merge", "unmerge"],
    path: str,
    sheet: str,
    address: str | None = None,
    confirm: bool = False,
) -> dict:
    """Manage merged cell ranges in a worksheet.

    operation='list': returns {"sheet": str, "merged_ranges": [str, ...], "count": int}.
    operation='merge': merges the address range. Requires EXCEL_ENABLE_WRITE=true and confirm=True.
    operation='unmerge': unmerges the address range. Requires EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "list":
        try:
            merged = wb.list_merged_cells(path, sheet)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
        return {"sheet": sheet, "merged_ranges": merged, "count": len(merged)}
    elif operation in ("merge", "unmerge"):
        if address is None:
            raise ToolError(f"address is required for operation='{operation}'")
        if operation == "merge":
            try:
                return wb.merge_cells(path, sheet=sheet, address=address, confirm=confirm)
            except ExcelMCPError as e:
                raise ToolError(str(e)) from e
        else:
            try:
                return wb.unmerge_cells(path, sheet=sheet, address=address, confirm=confirm)
            except ExcelMCPError as e:
                raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'list', 'merge', or 'unmerge'.")


@mcp.tool()
def freeze_panes(
    path: str,
    sheet: str,
    address: str | None = None,
    confirm: bool = False,
) -> dict:
    """Freeze rows/columns at a cell; pass address=None to unfreeze. Requires confirm=True and EXCEL_ENABLE_WRITE."""
    try:
        return wb.freeze_panes(path, sheet=sheet, freeze_at=address, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def auto_filter(
    path: str,
    sheet: str,
    address: str | None,
    confirm: bool = False,
) -> dict:
    """Apply Excel AutoFilter to a range; pass address=None to clear. Requires confirm=True and EXCEL_ENABLE_WRITE."""
    try:
        return wb.auto_filter(path, sheet=sheet, address=address, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# Phase 2.3 tools  (require EXCEL_ENABLE_WRITE=true + confirm=True)
# ---------------------------------------------------------------------------

@mcp.tool()
def rows(
    operation: Literal["insert", "delete"],
    path: str,
    sheet: str,
    start_row: int,
    count: int = 1,
    end_row: int | None = None,
    confirm: bool = False,
) -> dict:
    """Insert or delete rows in a worksheet.

    operation='insert': inserts `count` blank rows before `start_row`.
    operation='delete': deletes rows from `start_row` to `end_row` (inclusive; end_row defaults to start_row).
    Both require EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "insert":
        try:
            return wb.insert_rows(path, sheet, start_row, count=count, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "delete":
        effective_end = end_row if end_row is not None else start_row
        if effective_end < start_row:
            raise ToolError("end_row must be >= start_row")
        try:
            return wb.delete_rows(path, sheet, start_row, end_row=effective_end, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'insert' or 'delete'.")


@mcp.tool()
def cols(
    operation: Literal["insert", "delete"],
    path: str,
    sheet: str,
    start_col: int,
    count: int = 1,
    end_col: int | None = None,
    confirm: bool = False,
) -> dict:
    """Insert or delete columns in a worksheet.

    operation='insert': inserts `count` blank columns before `start_col` (1-based).
    operation='delete': deletes columns from `start_col` to `end_col` (inclusive; end_col defaults to start_col).
    Both require EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "insert":
        try:
            return wb.insert_cols(path, sheet, start_col, count=count, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "delete":
        effective_end = end_col if end_col is not None else start_col
        if effective_end < start_col:
            raise ToolError("end_col must be >= start_col")
        try:
            return wb.delete_cols(path, sheet, start_col, end_col=effective_end, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'insert' or 'delete'.")


@mcp.tool()
def protect(
    scope: Literal["sheet", "workbook"],
    path: str,
    sheet: str | None = None,
    password: str | None = None,
    lock_structure: bool = True,
    lock_windows: bool = False,
    confirm: bool = False,
) -> dict:
    """Protect or unprotect a sheet or the entire workbook structure.

    scope='sheet': protects/unprotects the named sheet. password=None clears protection.
    scope='workbook': protects/unprotects workbook structure. password=None clears protection.
    Both require EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if scope == "sheet":
        if sheet is None:
            raise ToolError("sheet is required when scope='sheet'")
        try:
            return wb.protect_sheet(path, sheet, password=password, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif scope == "workbook":
        try:
            return wb.protect_workbook(
                path, password=password,
                lock_structure=lock_structure, lock_windows=lock_windows,
                confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown scope: {scope}. Use 'sheet' or 'workbook'.")


@mcp.tool()
def set_print_area(
    path: str,
    sheet: str,
    address: str | None,
    confirm: bool = False,
) -> dict:
    """Set or clear the print area for a worksheet.

    Pass address=None to clear the print area.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.set_print_area(path, sheet, address, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# Phase 2.4 tools  (cell comments, sheet properties, data validation, workbook protection)
# ---------------------------------------------------------------------------

@mcp.tool()
def cell_comment(
    operation: Literal["get", "delete"],
    path: str,
    sheet: str,
    address: str,
    confirm: bool = False,
) -> dict:
    """Get or delete a cell comment.

    operation='get': returns comment text and author (read-only, no confirm needed).
    operation='delete': removes the comment. Requires EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "get":
        try:
            return wb.get_cell_comment(path, sheet, address)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "delete":
        try:
            return wb.delete_cell_comment(path, sheet, address, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'get' or 'delete'.")


@mcp.tool()
def sheet_properties(
    operation: Literal["get", "set"],
    path: str,
    sheet: str,
    tab_color: str | None = None,
    state: str | None = None,
    show_gridlines: bool | None = None,
    show_row_col_headers: bool | None = None,
    zoom_scale: int | None = None,
    confirm: bool = False,
) -> dict:
    """Get or set sheet display properties.

    operation='get': returns tab_color, state, show_gridlines, show_row_col_headers, zoom_scale.
    operation='set': updates one or more properties. Requires EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "get":
        try:
            return wb.get_sheet_properties(path, sheet)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "set":
        try:
            return wb.set_sheet_properties(
                path, sheet, tab_color=tab_color, state=state,
                show_gridlines=show_gridlines,
                show_row_col_headers=show_row_col_headers,
                zoom_scale=zoom_scale,
                confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'get' or 'set'.")


@mcp.tool()
def data_validation(
    operation: Literal["get", "add"],
    path: str,
    sheet: str,
    validations: list[dict] | None = None,
    confirm: bool = False,
) -> dict:
    """Get or add data validation rules in a worksheet.

    operation='get': returns all validation rules for the sheet.
    operation='add': adds validation rules from the validations list. Each item must have 'address' key.
    Requires EXCEL_ENABLE_WRITE=true and confirm=True for 'add'.
    """
    if operation == "get":
        try:
            return wb.get_data_validation_info(path, sheet)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "add":
        if not validations:
            raise ToolError("validations list is required for operation='add' and must not be empty")
        try:
            return wb.bulk_add_data_validation(
                path, sheet, validations=validations, confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'get' or 'add'.")


# ---------------------------------------------------------------------------
# Phase 2.5 tools — quick wins
# ---------------------------------------------------------------------------


@mcp.tool()
def validate_formula_syntax(path: str = "", formula: str = "") -> dict:
    """Check whether a formula string is syntactically valid for Excel without writing it."""
    try:
        return wb.validate_formula_syntax(formula, path=path)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def copy_range(
    path: str,
    sheet: str,
    src_address: str,
    dst_address: str,
    confirm: bool = False,
) -> dict:
    """Copy cells (values and formats) from src_address to dst_address within the same sheet.

    Requires EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    try:
        return wb.copy_range(path, sheet, src_address, dst_address, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def create_table(
    path: str,
    sheet: str,
    address: str,
    table_name: str | None = None,
    style: str = "TableStyleMedium9",
    confirm: bool = False,
) -> dict:
    """Create a native Excel table (ListObject) over the specified range in a worksheet.

    address: cell range for the table (e.g. 'A1:D20').
    style: Excel table style name (default 'TableStyleMedium9').
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    if table_name is None:
        raise ToolError("table_name is required")
    try:
        return wb.create_table(path, sheet, address, table_name, style=style, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# Phase 2.6 tools — read data validations, read cell comment
# ---------------------------------------------------------------------------

@mcp.tool()
def set_column_visibility(
    path: str,
    sheet: str,
    columns: list[str],
    hidden: bool,
    confirm: bool = False,
) -> dict:
    """Hide or show one or more columns.

    columns: list of column letters, e.g. ['A', 'C', 'D'].
    hidden: True to hide, False to show.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.set_column_visibility(
            path, sheet, columns=columns, hidden=hidden, confirm=confirm
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def set_workbook_calc_mode(
    path: str,
    calc_mode: str | None = None,
    full_calc_on_load: bool | None = None,
    confirm: bool = False,
) -> dict:
    """Set workbook calculation mode and/or fullCalcOnLoad flag.

    calc_mode: one of 'auto', 'manual', 'autoNoTable'. None = no change.
    full_calc_on_load: if True, Excel recalculates all formulas on next open.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.set_workbook_calc_mode(
            path, calc_mode=calc_mode, full_calc_on_load=full_calc_on_load,
            confirm=confirm,
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# Phase 2.7 — Hyperlinks
# ---------------------------------------------------------------------------


@mcp.tool()
def hyperlink(
    operation: Literal["get", "add", "remove"],
    path: str,
    sheet: str,
    links: list[dict] | None = None,
    apply_style: bool = True,
    address: str | None = None,
    confirm: bool = False,
) -> dict:
    """Manage hyperlinks in a worksheet.

    operation='get': returns all hyperlinks (read-only).
    operation='add': adds hyperlinks from the links list. Validates URL schemes —
        javascript:, data:, vbscript: are always blocked (OWASP A03).
    operation='remove': removes hyperlink from address. Requires confirm=True.
    Write operations require EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "get":
        try:
            return wb.get_hyperlinks(path, sheet)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "add":
        if not links:
            raise ToolError("links list is required for operation='add' and must not be empty")
        try:
            return wb.bulk_add_hyperlinks(path, sheet, links, apply_style=apply_style, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "remove":
        if address is None:
            raise ToolError("address is required for operation='remove'")
        try:
            return wb.remove_hyperlink(path, sheet, address, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'get', 'add', or 'remove'.")
