"""
Content dispatcher handler for wordmcp.

Provides a single `content` tool that routes to existing individual content
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
    "add_list",
    "add_page_break",
    "insert_image",
    "add_hyperlink",
    "add_footnote",
    "find_replace",
    "set_properties",
)


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _content(
        operation: str,
        path: str,
        # --- add_list ---
        items: list | str | None = None,
        list_type: str = "bullet",
        level: int = 0,
        # --- insert_image ---
        image_path: str | None = None,
        width_inches: float | None = None,
        # --- add_hyperlink / add_footnote ---
        paragraph_index: int | None = None,
        text: str | None = None,
        url: str | None = None,
        # --- find_replace ---
        find_text: str | None = None,
        replace_text: str | None = None,
        occurrence: int | None = None,
        # --- set_properties ---
        title: str | None = None,
        author: str | None = None,
        subject: str | None = None,
        keywords: str | None = None,
        category: str | None = None,
        # --- mutation gate ---
        confirm: bool = False,
    ) -> Any:
        """Unified content dispatcher — routes to the appropriate content tool.

        operation (required): one of add_list | add_page_break | insert_image |
            add_hyperlink | add_footnote | find_replace | set_properties.

        path (required): absolute path to the .docx file.

        Per-operation parameters:

          add_list        — items (list[str], required), list_type (str, default
                            'bullet'; use 'number' for numbered lists), level
                            (int, default 0 — nesting depth).
          add_page_break  — no extra params needed.
          insert_image    — image_path (str, required — absolute path to the
                            image file), width_inches (float, optional — render
                            width; aspect ratio preserved).
          add_hyperlink   — paragraph_index (int, 0-based, required), text (str,
                            required — display text), url (str, required — must
                            start with http:// or https://).
          add_footnote    — paragraph_index (int, 0-based, required), text (str,
                            required — footnote body text).
          find_replace    — find_text (str, required), replace_text (str,
                            required), paragraph_index (int, optional — limit
                            to a single paragraph), occurrence (int, optional —
                            1-based index of match to replace; omit for all).
          set_properties  — title (str, optional), author (str, optional),
                            subject (str, optional), keywords (str, optional),
                            category (str, optional).

        All operations are write-gated: WORD_ENABLE_WRITE=true and confirm=True
        are required before any file I/O is performed.

        Unknown operation returns a structured error dict:
          {"error": "unknown_operation", "operation": op, "valid_operations": [...]}

        Deprecated aliases (still registered on the server module):
          add_list              — deprecated, use content(operation="add_list", ...)
          add_page_break        — deprecated, use content(operation="add_page_break", ...)
          insert_image          — deprecated, use content(operation="insert_image", ...)
          add_hyperlink         — deprecated, use content(operation="add_hyperlink", ...)
          add_footnote          — deprecated, use content(operation="add_footnote", ...)
          find_replace          — deprecated, use content(operation="find_replace", ...)
          set_document_properties — deprecated, use content(operation="set_properties", ...)
        """
        op = operation.strip().lower()

        # All content operations are write-gated — check gates first, before any I/O
        if op in _VALID_OPERATIONS:
            if os.getenv("WORD_ENABLE_WRITE", "false").lower() != "true":
                from fastmcp.exceptions import ToolError
                raise ToolError(
                    "NotAllowedError: WORD_ENABLE_WRITE is not set to 'true'. "
                    "Set WORD_ENABLE_WRITE=true to enable content write operations."
                )
            if not confirm:
                from fastmcp.exceptions import ToolError
                raise ToolError(
                    "ValidationError: confirm=True is required for content write operations. "
                    "Pass confirm=True to proceed."
                )

            if op == "add_list":
                if items is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("items is required for operation='add_list'")
                return _server_docx_tools.add_list(
                    runtime,
                    path=path,
                    items=items,
                    list_type=list_type,
                    level=level,
                    confirm=confirm,
                )

            if op == "add_page_break":
                return _server_docx_tools.add_page_break(
                    runtime,
                    path=path,
                    confirm=confirm,
                )

            if op == "insert_image":
                if image_path is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("image_path is required for operation='insert_image'")
                return _server_docx_tools.insert_image(
                    runtime,
                    path=path,
                    image_path=image_path,
                    width_inches=width_inches,
                    confirm=confirm,
                )

            if op == "add_hyperlink":
                if paragraph_index is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("paragraph_index is required for operation='add_hyperlink'")
                if text is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("text is required for operation='add_hyperlink'")
                if url is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("url is required for operation='add_hyperlink'")
                return _server_docx_tools.add_hyperlink(
                    runtime,
                    path=path,
                    paragraph_index=paragraph_index,
                    text=text,
                    url=url,
                    confirm=confirm,
                )

            if op == "add_footnote":
                if paragraph_index is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("paragraph_index is required for operation='add_footnote'")
                if text is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("text is required for operation='add_footnote'")
                return _server_docx_tools.add_footnote(
                    runtime,
                    path=path,
                    paragraph_index=paragraph_index,
                    text=text,
                    confirm=confirm,
                )

            if op == "find_replace":
                if find_text is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("find_text is required for operation='find_replace'")
                if replace_text is None:
                    from fastmcp.exceptions import ToolError
                    raise ToolError("replace_text is required for operation='find_replace'")
                return _server_docx_tools.find_replace(
                    runtime,
                    path=path,
                    find_text=find_text,
                    replace_text=replace_text,
                    paragraph_index=paragraph_index,
                    occurrence=occurrence,
                    confirm=confirm,
                )

            if op == "set_properties":
                return _server_docx_tools.set_document_properties(
                    runtime,
                    path=path,
                    title=title,
                    author=author,
                    subject=subject,
                    keywords=keywords,
                    category=category,
                    confirm=confirm,
                )

        # Unknown operation — return structured error (no exception, caller can inspect)
        return {
            "error": "unknown_operation",
            "operation": operation,
            "valid_operations": list(_VALID_OPERATIONS),
        }

    return {
        "content": bind_tool(_content, "content"),
    }
