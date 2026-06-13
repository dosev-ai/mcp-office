"""
Paragraph dispatcher handler for wordmcp.

Provides a single `paragraph` tool that routes to existing individual paragraph
tools based on the `operation` parameter.  All individual tools remain registered
(with deprecation notes) — this dispatcher is an additive convenience wrapper.
"""
from __future__ import annotations

from collections.abc import Callable
from typing import Any

from wordmcp import _server_docx_tools
from wordmcp._server_runtime import ServerRuntime

_VALID_OPERATIONS = (
    "list",
    "read",
    "add",
    "add_heading",
    "insert",
    "update",
    "delete",
    "set_format",
    "bulk_add",
    "bulk_update",
)


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _paragraph(
        operation: str,
        path: str,
        # --- read ops ---
        paragraph_index: int | None = None,
        # --- add / insert / update ---
        text: str | None = None,
        style: str | None = None,
        level: int = 1,
        new_text: str | None = None,
        # --- set_format ---
        space_before: float | None = None,
        space_after: float | None = None,
        line_spacing: float | None = None,
        table_cell: dict | None = None,
        # --- bulk ops ---
        paragraphs: list | str | None = None,
        updates: list | str | None = None,
        # --- mutation gate ---
        confirm: bool = False,
    ) -> Any:
        """Unified paragraph dispatcher — routes to the appropriate paragraph tool.

        operation (required): one of list | read | add | add_heading | insert |
            update | delete | set_format | bulk_add | bulk_update.

        path (required): absolute path to the .docx file.

        Per-operation parameters:

          list       — no extra params needed.
          read       — paragraph_index (int, 0-based, required).
          add        — text (str, required), style (str, optional).
          add_heading — text (str, required), level (int 0-9, default 1).
          insert     — paragraph_index (int, required), text (str, required),
                       style (str, optional).
          update     — paragraph_index (int, required), new_text (str, required).
          delete     — paragraph_index (int, required).
          set_format — paragraph_index (int, required), space_before (float pt),
                       space_after (float pt), line_spacing (float multiplier),
                       table_cell (dict with row/col keys, optional).
          bulk_add   — paragraphs (list of {text, style?} dicts, max 500).
          bulk_update — updates (list of {paragraph_index, new_text} dicts, max 500).

        Write operations (add, add_heading, insert, update, delete, set_format,
        bulk_add, bulk_update) require WORD_ENABLE_WRITE=true and confirm=True.

        Unknown operation returns a structured error dict:
          {"error": "unknown_operation", "operation": op, "valid_operations": [...]}
        """
        op = operation.strip().lower()

        if op == "list":
            return _server_docx_tools.list_paragraphs(runtime, path)

        if op == "read":
            if paragraph_index is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("paragraph_index is required for operation='read'")
            return _server_docx_tools.read_paragraph(runtime, path, paragraph_index)

        if op == "add":
            if text is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("text is required for operation='add'")
            return _server_docx_tools.add_paragraph(
                runtime, path, text, style=style, confirm=confirm
            )

        if op == "add_heading":
            if text is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("text is required for operation='add_heading'")
            return _server_docx_tools.add_heading(
                runtime, path, text, level=level, confirm=confirm
            )

        if op == "insert":
            if paragraph_index is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("paragraph_index is required for operation='insert'")
            if text is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("text is required for operation='insert'")
            return _server_docx_tools.insert_paragraph(
                runtime,
                path=path,
                paragraph_index=paragraph_index,
                text=text,
                style=style,
                confirm=confirm,
            )

        if op == "update":
            if paragraph_index is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("paragraph_index is required for operation='update'")
            if new_text is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("new_text is required for operation='update'")
            return _server_docx_tools.update_paragraph(
                runtime,
                path=path,
                paragraph_index=paragraph_index,
                new_text=new_text,
                confirm=confirm,
            )

        if op == "delete":
            if paragraph_index is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("paragraph_index is required for operation='delete'")
            return _server_docx_tools.delete_paragraph(
                runtime,
                path=path,
                paragraph_index=paragraph_index,
                confirm=confirm,
            )

        if op == "set_format":
            if paragraph_index is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("paragraph_index is required for operation='set_format'")
            return _server_docx_tools.set_paragraph_format(
                runtime,
                path=path,
                paragraph_index=paragraph_index,
                space_before=space_before,
                space_after=space_after,
                line_spacing=line_spacing,
                table_cell=table_cell,
                confirm=confirm,
            )

        if op == "bulk_add":
            if paragraphs is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("paragraphs is required for operation='bulk_add'")
            return _server_docx_tools.bulk_add_paragraphs(
                runtime, path=path, paragraphs=paragraphs, confirm=confirm
            )

        if op == "bulk_update":
            if updates is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("updates is required for operation='bulk_update'")
            return _server_docx_tools.bulk_update_paragraphs(
                runtime, path=path, updates=updates, confirm=confirm
            )

        # Unknown operation — return structured error (no exception, caller can inspect)
        return {
            "error": "unknown_operation",
            "operation": operation,
            "valid_operations": list(_VALID_OPERATIONS),
        }

    return {
        "paragraph": bind_tool(_paragraph, "paragraph"),
    }
