"""Chart MCP tools for excelmcp.

Extracted from server.py Phase 2.0 chart section.
Tool: chart (unified — replaces create_chart, list_charts, delete_chart, update_chart_data).

Import side-effect: the chart tool is registered on the shared ``mcp``
instance from ``excelmcp._server_instance`` at import time.
"""
from __future__ import annotations

from typing import Literal

from fastmcp.exceptions import ToolError

from excelmcp import workbook_openpyxl as wb
from excelmcp._server_instance import mcp
from excelmcp.workbook_openpyxl import ExcelMCPError

# ---------------------------------------------------------------------------
# Chart tool  (Phase 2.0 — write ops require EXCEL_ENABLE_WRITE=true + confirm=True)
# ---------------------------------------------------------------------------

@mcp.tool()
def chart(
    operation: Literal["create", "list", "delete", "update"],
    path: str,
    sheet: str,
    chart_type: str | None = None,
    data_address: str | None = None,
    title: str = "",
    target_address: str | None = None,
    anchor: str | None = None,
    width_px: int | None = None,
    height_px: int | None = None,
    chart_index: int | None = None,
    confirm: bool = False,
) -> dict:
    """Manage Excel charts: create, list, delete, or update.

    operation='list': returns {"charts": [...], "count": N} — no confirm needed.
    operation='create': chart_type and data_address required. confirm=True required.
      chart_type: bar, line, scatter, pie.
      data_address: e.g. 'A1:B5' (include headers).
      target_address: cell where chart is placed (default 'D1').
      width_px / height_px: chart size in pixels (1–5000).
    operation='delete': chart_index (0-based) required. confirm=True required.
    operation='update': chart_index and data_address required. confirm=True required.
    Write operations require EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "list":
        try:
            return wb.list_charts(path, sheet)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "create":
        if target_address is None:
            target_address = "D1"
        if not isinstance(target_address, str):
            raise ToolError("target_address must be a string when supplied")
        if target_address.strip() == "":
            raise ToolError("target_address must not be empty when supplied")
        if chart_type is None:
            raise ToolError("chart_type is required for operation='create'")
        if data_address is None:
            raise ToolError("data_address is required for operation='create'")
        try:
            return wb.create_chart(
                path, sheet, chart_type, data_address,
                title=title, target_cell=target_address,
                anchor=anchor,
                width_px=width_px, height_px=height_px,
                confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "delete":
        if chart_index is None:
            raise ToolError("chart_index is required for operation='delete'")
        try:
            return wb.delete_chart(path, sheet, chart_index, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "update":
        if chart_index is None:
            raise ToolError("chart_index is required for operation='update'")
        if data_address is None:
            raise ToolError("data_address is required for operation='update'")
        try:
            return wb.update_chart_data(
                path, sheet, chart_index, data_address, confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'create', 'list', 'delete', or 'update'.")
