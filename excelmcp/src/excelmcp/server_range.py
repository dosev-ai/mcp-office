"""Range-family MCP tools for excelmcp.

Covers: get_used_range, cell, range_io, read_sheet_all, get_cell_metadata,
consolidate_ranges.

Import side-effect: importing this module registers all tools on the shared
mcp instance from excelmcp._server_instance.
"""
from __future__ import annotations

import logging
from typing import Any, Literal

from fastmcp.exceptions import ToolError

from excelmcp import workbook_openpyxl as wb
from excelmcp._core import NotAllowedError, ValidationError
from excelmcp._server_instance import mcp
from excelmcp.workbook_openpyxl import ExcelMCPError

logger = logging.getLogger(__name__)


@mcp.tool()
def get_used_range(path: str, sheet: str) -> dict:
    """Return the bounding box of used cells in *sheet* of the workbook at *path*."""
    try:
        return wb.get_used_range(path, sheet)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def cell(
    operation: Literal["read", "write"],
    path: str,
    sheet: str,
    address: str,
    value: object = None,
    confirm: bool = False,
) -> dict:
    """Read or write a single cell value.

    operation='read': returns cell value, type, and formula (read-only).
    operation='write': writes value to the cell. Requires EXCEL_ENABLE_WRITE=true and confirm=True.
    value=None is not allowed for 'write' (use range_io with values=[[None]] to clear a cell).
    """
    if operation == "read":
        try:
            return wb.read_cell(path, sheet, address)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "write":
        if value is None:
            raise ToolError("value is required for operation='write' (use range_io with values=[[None]] to clear a cell)")
        try:
            return wb.write_cell(path, sheet, address, value, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'read' or 'write'.")


@mcp.tool()
def range_io(
    operation: Literal["read", "write"],
    path: str,
    sheet: str,
    address: str,
    values: list[list] | None = None,
    confirm: bool = False,
) -> dict:
    """Read or write a rectangular range of cells.

    operation='read': returns {"data": [[{...},...], ...], "rows": int, "cols": int}.
    operation='write': writes values 2-D array starting from address cell.
    Requires EXCEL_ENABLE_WRITE=true and confirm=True for 'write'.
    """
    if operation == "read":
        try:
            raw = wb.read_range(path, sheet, address)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
        rows_count = len(raw)
        cols_count = len(raw[0]) if rows_count > 0 else 0
        return {"data": raw, "rows": rows_count, "cols": cols_count}
    elif operation == "write":
        if values is None:
            raise ToolError("values is required for operation='write'")
        try:
            return wb.write_range(path, sheet, address, values, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'read' or 'write'.")


@mcp.tool()
def read_sheet_all(path: str, sheet: str) -> dict:
    """Read the entire used range of *sheet* from the workbook at *path*.

    Returns all cell data, row count, and column count.
    Subject to EXCEL_MAX_RANGE_CELLS guard.
    """
    try:
        return wb.read_sheet_all(path, sheet)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def get_cell_metadata(path: str, sheet: str, address: str) -> dict:
    """Return formatting and metadata for a single cell.

    Includes font, fill, number format, comment, and hyperlink information.
    *address* must be in A1 notation (e.g. "B3").
    """
    try:
        return wb.get_cell_metadata(path, sheet, address)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def consolidate_ranges(
    ranges: list[dict],
    source_label_key: str = "source",
    max_rows: int = 500,
) -> dict[str, Any]:
    """Read multiple ranges from workbooks and merge into one table.

    Each entry in ``ranges`` must have keys: ``path``, ``sheet``, ``range``,
    and optionally ``label``. Enforces EXCEL_ALLOWLIST_ROOTS via wb.read_range
    (BLOCKER-2 compliance -- no direct openpyxl.load_workbook calls here).

    Does NOT require EXCEL_ENABLE_WRITE. Does NOT require confirm.
    """
    try:
        if len(ranges) == 0:
            raise ValidationError("consolidate_ranges: ranges must not be empty")
        if len(ranges) > 20:
            raise ValidationError(
                f"consolidate_ranges: too many source ranges ({len(ranges)}); maximum is 20"
            )
        _RESERVED_ROW_KEYS = frozenset({"row_index", "data"})
        if source_label_key in _RESERVED_ROW_KEYS:
            raise ValidationError(
                f"source_label_key {source_label_key!r} conflicts with reserved output keys "
                f"{sorted(_RESERVED_ROW_KEYS)}; choose a different key name"
            )
        for i, t in enumerate(ranges):
            for required_key in ("path", "sheet", "range"):
                if required_key not in t:
                    raise ValidationError(
                        f"ranges[{i}] missing required key: {required_key!r}"
                    )
        all_rows: list[dict] = []
        hard_cap = min(max_rows, 500)
        truncated = False
        for t in ranges:
            path_str = t["path"]
            sheet_name = t["sheet"]
            address = t["range"]
            label = t.get("label", f"{path_str}:{sheet_name}!{address}")
            raw: list[list[dict]] = wb.read_range(
                path=path_str,
                sheet=sheet_name,
                a1_range=address,
            )
            for row_data in raw:
                if len(all_rows) >= hard_cap:
                    truncated = True
                    break
                all_rows.append({
                    source_label_key: label,
                    "row_index": len(all_rows),
                    "data": row_data,
                })
            if truncated:
                break
        return {
            "ok": True,
            "source_count": len(ranges),
            "total_rows": len(all_rows),
            "truncated": truncated,
            "rows": all_rows,
        }
    except (ValidationError, NotAllowedError):
        raise
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e
    except Exception as e:
        logger.error("consolidate_ranges failed: %s", e, exc_info=True)
        raise ToolError("consolidate_ranges failed -- check server logs") from e
