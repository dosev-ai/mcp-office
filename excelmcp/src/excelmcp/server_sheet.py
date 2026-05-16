"""Sheet-family MCP tools for excelmcp.

Covers: sheet (list/add/rename/delete dispatcher).

Import side-effect: importing this module registers all tools on the shared
mcp instance from excelmcp._server_instance.
"""
from __future__ import annotations

from typing import Literal

from fastmcp.exceptions import ToolError

from excelmcp import workbook_openpyxl as wb
from excelmcp._server_instance import mcp
from excelmcp.workbook_openpyxl import ExcelMCPError


@mcp.tool()
def sheet(
    operation: Literal["list", "add", "rename", "delete"],
    path: str,
    sheet_name: str | None = None,
    new_name: str | None = None,
    position: int | None = None,
    confirm: bool = False,
) -> dict:
    """Manage workbook sheets: list, add, rename, or delete.

    operation='list': returns {"sheets": [str, ...], "count": int} — no confirm needed.
    operation='add': adds a new sheet named sheet_name at optional position.
    operation='rename': renames sheet_name to new_name.
    operation='delete': deletes sheet named sheet_name.
    Write operations require EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "list":
        try:
            result = wb.list_sheets(path)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
        if isinstance(result, list):
            return {"sheets": result, "count": len(result)}
        return result
    elif operation == "add":
        if sheet_name is None:
            raise ToolError("sheet_name is required for operation='add'")
        try:
            return wb.add_sheet(path, sheet_name, position=position, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "rename":
        if sheet_name is None:
            raise ToolError("sheet_name is required for operation='rename'")
        if new_name is None:
            raise ToolError("new_name is required for operation='rename'")
        try:
            return wb.rename_sheet(path, sheet_name, new_name, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "delete":
        if sheet_name is None:
            raise ToolError("sheet_name is required for operation='delete'")
        try:
            return wb.delete_sheet(path, sheet_name, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'list', 'add', 'rename', or 'delete'.")
