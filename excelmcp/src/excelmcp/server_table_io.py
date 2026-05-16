"""Table/import-family MCP tools for excelmcp.

Covers: list_tables, read_table, append_rows, find_replace, import_csv_to_sheet,
named_range.

Import side-effect: importing this module registers all tools on the shared
mcp instance from excelmcp._server_instance.
"""
from __future__ import annotations

from typing import Any, Literal

from fastmcp.exceptions import ToolError

from excelmcp import workbook_openpyxl as wb
from excelmcp._server_instance import mcp
from excelmcp.workbook_openpyxl import ExcelMCPError


@mcp.tool()
def list_tables(path: str, sheet: str | None = None) -> list[dict]:
    """List all Excel tables (ListObjects) in *sheet*, or across all sheets when
    *sheet* is omitted.

    Each entry contains sheet, name, display_name, ref, header_row, and totals_row.
    Omitting *sheet* enables workbook-wide table discovery.
    """
    try:
        return wb.list_tables(path, sheet)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def read_table(path: str, sheet: str, table_name: str) -> dict:
    """Read all rows from a named Excel table in *sheet*.

    Returns table metadata, headers, and all data rows.
    """
    try:
        return wb.read_table(path, sheet, table_name)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def append_rows(
    path: str,
    sheet: str,
    rows: list[list[Any]],
    confirm: bool = False,
) -> dict:
    """Append *rows* after the last used row in *sheet*.

    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.append_rows(path, sheet, rows, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def find_replace(
    path: str,
    sheet: str,
    find_value: str,
    replace_value: str,
    match_case: bool = False,
    whole_cell: bool = True,
    confirm: bool = False,
) -> dict:
    """Find and replace string values across all cells in *sheet*.

    Set ``whole_cell=False`` for substring replacement.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.find_replace(
            path, sheet, find_value, replace_value,
            match_case=match_case, whole_cell=whole_cell, confirm=confirm
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def import_csv_to_sheet(
    path: str,
    csv_path: str,
    sheet: str,
    skip_header: bool = False,
    overwrite: bool = False,
    confirm: bool = False,
) -> dict:
    """Import a CSV or TSV file into *sheet*.

    All rows including the header row are imported by default so column names
    are preserved.  Set ``skip_header=True`` to discard the first row.
    Set ``overwrite=True`` to replace an existing sheet with the same name.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.import_csv_to_sheet(
            path, csv_path, sheet,
            skip_header=skip_header, overwrite=overwrite, confirm=confirm
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def named_range(
    operation: Literal["list", "read", "write"],
    path: str,
    name: str | None = None,
    sheet: str | None = None,
    address: str | None = None,
    confirm: bool = False,
) -> dict:
    """List, read, or write Excel named ranges (defined names).

    operation='list': returns {"named_ranges": [{name, refers_to, scope}, ...], "count": N}.
    operation='read': reads the data values at the named range. name is required.
    operation='write': creates or updates a named range. name, sheet, address, confirm all required.
    'write' requires EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "list":
        try:
            result = wb.get_named_ranges(path)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
        if isinstance(result, list):
            return {"named_ranges": result, "count": len(result)}
        return result
    elif operation == "read":
        if name is None:
            raise ToolError("name is required for operation='read'")
        try:
            return wb.read_named_range(path, name)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "write":
        if name is None:
            raise ToolError("name is required for operation='write'")
        if sheet is None:
            raise ToolError("sheet is required for operation='write'")
        if address is None:
            raise ToolError("address is required for operation='write'")
        try:
            return wb.write_named_range(path, name=name, sheet=sheet, range_address=address, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'list', 'read', or 'write'.")
