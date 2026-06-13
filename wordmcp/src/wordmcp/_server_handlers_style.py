"""
Style dispatcher handler for wordmcp.

Provides a single `style` tool that routes to existing individual style
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
    "apply",
)


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _style(
        operation: str,
        path: str,
        # --- list op ---
        style_type: str | None = None,
        # --- apply op ---
        paragraph_index: int | None = None,
        style_name: str | None = None,
        # --- mutation gate ---
        confirm: bool = False,
    ) -> Any:
        """Unified style dispatcher — routes to the appropriate style tool.

        operation (required): one of list | apply.

        path (required): absolute path to the .docx file.

        Per-operation parameters:

          list    — style_type (str, optional — e.g. 'paragraph', 'character',
                    'table', 'numbering'; omit for all styles).
          apply   — paragraph_index (int, 0-based, required), style_name (str,
                    required — the Word style name to apply, e.g. 'Heading 1').

        Write operations (apply) require WORD_ENABLE_WRITE=true and confirm=True.

        Unknown operation returns a structured error dict:
          {"error": "unknown_operation", "operation": op, "valid_operations": [...]}

        Deprecated aliases (still registered on the server module):
          list_styles  — deprecated, use style(operation="list")
          apply_style  — deprecated, use style(operation="apply", ...)
        """
        op = operation.strip().lower()

        # Read-only operations — no gates needed
        if op == "list":
            return _server_docx_tools.list_styles(runtime, path, style_type=style_type)

        # Write operations — gates checked first, before any file I/O
        if op == "apply":
            if os.getenv("WORD_ENABLE_WRITE", "false").lower() != "true":
                from fastmcp.exceptions import ToolError
                raise ToolError(
                    "NotAllowedError: WORD_ENABLE_WRITE is not set to 'true'. "
                    "Set WORD_ENABLE_WRITE=true to enable style write operations."
                )
            if not confirm:
                from fastmcp.exceptions import ToolError
                raise ToolError(
                    "ValidationError: confirm=True is required for style write operations. "
                    "Pass confirm=True to proceed."
                )
            if paragraph_index is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("paragraph_index is required for operation='apply'")
            if style_name is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("style_name is required for operation='apply'")
            return _server_docx_tools.apply_style(
                runtime,
                path=path,
                paragraph_index=paragraph_index,
                style=style_name,
                confirm=confirm,
            )

        # Unknown operation — return structured error (no exception, caller can inspect)
        return {
            "error": "unknown_operation",
            "operation": operation,
            "valid_operations": list(_VALID_OPERATIONS),
        }

    return {
        "style": bind_tool(_style, "style"),
    }
