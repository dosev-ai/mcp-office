"""
Table dispatcher handler for wordmcp.

Provides a single `table` tool that routes to existing individual table
tools based on the `operation` parameter.  All individual tools remain
registered (with deprecation notes) — this dispatcher is an additive
convenience wrapper.
"""
from __future__ import annotations

import os
from collections.abc import Callable
from typing import Any

from wordmcp import _server_docx_tools
from wordmcp._server_runtime import ServerRuntime

_VALID_OPERATIONS = (
    "list",
    "read",
    "add",
    "update_cell",
    "bulk_update_cells",
)


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _table(
        operation: str,
        path: str,
        # --- read ops ---
        table_index: int | None = None,
        # --- add ---
        rows: int | None = None,
        cols: int | None = None,
        data: list | str | None = None,
        style: str | None = None,
        # --- update_cell ---
        row_index: int | None = None,
        col_index: int | None = None,
        new_text: str | None = None,
        # --- bulk_update_cells ---
        updates: list | str | None = None,
        # --- mutation gate ---
        confirm: bool = False,
    ) -> Any:
        """Unified table dispatcher — routes to the appropriate table tool.

        operation (required): one of list | read | add | update_cell |
            bulk_update_cells.

        path (required): absolute path to the .docx file.

        Per-operation parameters:

          list             — no extra params needed.
          read             — table_index (int, 0-based, required).
          add              — rows (int, required), cols (int, required),
                             data (list[list[str]], optional 2D cell values),
                             style (str, optional Word table style name).
          update_cell      — table_index (int, required), row_index (int,
                             required), col_index (int, required),
                             new_text (str, required).
          bulk_update_cells — updates (list of dicts with keys: table_index,
                             row, col, new_text; max 200 per call, required).

        Write operations (add, update_cell, bulk_update_cells) require
        WORD_ENABLE_WRITE=true and confirm=True.

        Unknown operation returns a structured error dict:
          {"error": "unknown_operation", "operation": op, "valid_operations": [...]}
        """
        op = operation.strip().lower()

        # Read-only operations — no gates needed
        if op == "list":
            return _server_docx_tools.list_tables(runtime, path)

        if op == "read":
            if table_index is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("table_index is required for operation='read'")
            return _server_docx_tools.read_table(runtime, path, table_index)

        # Write operations — gates checked first, before any file I/O
        if op in ("add", "update_cell", "bulk_update_cells"):
            if os.getenv("WORD_ENABLE_WRITE", "false").lower() != "true":
                from fastmcp.exceptions import ToolError
                raise ToolError(
                    "NotAllowedError: WORD_ENABLE_WRITE is not set to 'true'. "
                    "Set WORD_ENABLE_WRITE=true to enable table write operations."
                )
            if not confirm:
                from fastmcp.exceptions import ToolError
                raise ToolError(
                    "ValidationError: confirm=True is required for table write operations. "
                    "Pass confirm=True to proceed."
                )

            if op == "add":
                if rows is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("rows is required for operation='add'")
                if cols is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("cols is required for operation='add'")
                return _server_docx_tools.add_table(
                    runtime,
                    path,
                    rows,
                    cols,
                    data=data,
                    style=style,
                    confirm=confirm,
                )

            if op == "update_cell":
                if table_index is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("table_index is required for operation='update_cell'")
                if row_index is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("row_index is required for operation='update_cell'")
                if col_index is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("col_index is required for operation='update_cell'")
                if new_text is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("new_text is required for operation='update_cell'")
                return _server_docx_tools.update_table_cell(
                    runtime,
                    path=path,
                    table_index=table_index,
                    row=row_index,
                    col=col_index,
                    new_text=new_text,
                    confirm=confirm,
                )

            if op == "bulk_update_cells":
                if updates is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("updates is required for operation='bulk_update_cells'")
                return _server_docx_tools.bulk_update_table_cells(
                    runtime,
                    path=path,
                    updates=updates,
                    confirm=confirm,
                )

        # Unknown operation — return structured error (no exception, caller can inspect)
        return {
            "error": "unknown_operation",
            "operation": operation,
            "valid_operations": list(_VALID_OPERATIONS),
        }

    return {
        "table": bind_tool(_table, "table"),
    }
