"""Workbook-family MCP tools for excelmcp.

Covers: create_workbook, save, workbook_metadata (dispatcher), evaluate_formula.

Import side-effect: importing this module registers all tools on the shared
mcp instance from excelmcp._server_instance.
"""
from __future__ import annotations

from typing import Any, Literal

from fastmcp.exceptions import ToolError

from excelmcp import workbook_openpyxl as wb
from excelmcp._com import _evaluate_formula_live
from excelmcp._server_instance import mcp
from excelmcp.workbook_openpyxl import ExcelMCPError


@mcp.tool()
def create_workbook(
    path: str,
    sheets: list[str] | None = None,
    overwrite: bool = False,
    confirm: bool = False,
) -> dict:
    """Create a new .xlsx workbook at the specified path.

    sheets: worksheet names to create (default: ["Sheet1"]).
    overwrite: if False (default), raises ExcelMCPError if the file exists.
    confirm: must be True (write gate — prevents accidental creation).
    Returns: {"ok": true, "path": "...", "sheets": [...], "created": true}
    """
    try:
        return wb.create_workbook(
            path, sheets=sheets, overwrite=overwrite, confirm=confirm
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def save(path: str, confirm: bool = False) -> dict:
    """Flush the in-memory workbook at *path* to disk.

    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    Evicts the workbook from the cache after saving.
    """
    try:
        return wb.save(path, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def workbook_metadata(
    operation: Literal["read", "write"],
    path: str,
    payload: dict[str, Any] | None = None,
    confirm: bool = False,
) -> dict:
    """Read or write workbook-level metadata stored on a hidden worksheet.

    operation='read': returns whether metadata is present plus stored payload details.
    operation='write': stores payload on hidden sheet _MCP_META. Requires
    EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "read":
        try:
            return wb.read_workbook_metadata(path)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    if operation == "write":
        if payload is None:
            raise ToolError("payload is required for operation='write'")
        try:
            return wb.write_workbook_metadata(path, payload, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    raise ToolError("Unknown operation: use 'read' or 'write'.")


@mcp.tool()
def evaluate_formula(path: str, sheet: str, address: str, live: bool = False) -> dict:
    """Return the formula string and value for *address*.

    When live=False (default): openpyxl returns the last-saved cached value.
    When live=True: opens workbook via Excel COM for live evaluation.
    Useful for volatile functions (TODAY, NOW, RAND, OFFSET).
    Requires EXCEL_ENABLE_COM=true when live=True.
    """
    if live:
        try:
            return _evaluate_formula_live(path=path, sheet=sheet, cell_address=address)
        except Exception as e:
            raise ToolError(str(e)) from e
    try:
        return wb.evaluate_formula(path, sheet, address)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e
