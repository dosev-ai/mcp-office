"""
Document dispatcher handler for wordmcp.

Provides a single `document` tool that routes to existing individual document
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
    "save",
    "search_text",
    "get_headers_footers",
)


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _document(
        operation: str,
        path: str,
        # --- search_text op ---
        query: str | None = None,
        case_sensitive: bool = True,
        max_results: int = 100,
        # --- mutation gate ---
        confirm: bool = False,
    ) -> Any:
        """Unified document dispatcher — routes to the appropriate document tool.

        operation (required): one of save | search_text | get_headers_footers.

        path (required): absolute path to the .docx file.

        Per-operation parameters:

          save               — Persist the document to disk.
                               Requires WORD_ENABLE_WRITE=true and confirm=True.
          search_text        — Search for text across body paragraphs and table
                               cells.  query (str, required — the text to find),
                               case_sensitive (bool, default True),
                               max_results (int, default 100).
          get_headers_footers — Return header and footer text for each section
                               of the document. No additional parameters needed.

        Write operations (save) require WORD_ENABLE_WRITE=true and confirm=True.

        Unknown operation returns a structured error dict (no exception):
          {"error": "unknown_operation", "operation": op, "valid_operations": [...]}

        Deprecated aliases (still registered on the server module):
          save               — deprecated, use document(operation="save", ...)
          search_text        — deprecated, use document(operation="search_text", ...)
          get_headers_footers — deprecated, use document(operation="get_headers_footers")
        """
        op = operation.strip().lower()

        # Read-only operations — no gates needed
        if op == "search_text":
            if query is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("query is required for operation='search_text'")
            return _server_docx_tools.search_text(
                runtime,
                path=path,
                query=query,
                case_sensitive=case_sensitive,
                max_results=max_results,
            )

        if op == "get_headers_footers":
            return _server_docx_tools.get_headers_footers(runtime, path=path)

        # Write operations — gates checked first, before any file I/O
        if op == "save":
            if os.getenv("WORD_ENABLE_WRITE", "false").lower() != "true":
                from fastmcp.exceptions import ToolError
                raise ToolError(
                    "NotAllowedError: WORD_ENABLE_WRITE is not set to 'true'. "
                    "Set WORD_ENABLE_WRITE=true to enable document write operations."
                )
            if not confirm:
                from fastmcp.exceptions import ToolError
                raise ToolError(
                    "ValidationError: confirm=True is required for document write operations. "
                    "Pass confirm=True to proceed."
                )
            return _server_docx_tools.save(runtime, path=path, confirm=confirm)

        # Unknown operation — return structured error (no exception, caller can inspect)
        return {
            "error": "unknown_operation",
            "operation": operation,
            "valid_operations": list(_VALID_OPERATIONS),
        }

    return {
        "document": bind_tool(_document, "document"),
    }
