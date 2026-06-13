"""MCP tool handler registration for assembly tools.

Wires bulk_find_replace, manage_hyperlinks, and verify_no_placeholders
from wordmcp._docx.assembly into the FastMCP server.

Layer contract: NO COM imports, NO MCP framework imports beyond the
bind_tool callable passed in from server.py.
"""
from __future__ import annotations

from typing import Any

from fastmcp.exceptions import ToolError

from wordmcp._docx import assembly as _asm
from wordmcp.document_docx import NotAllowedError, ValidationError


def _tool_error(exc: Exception) -> ToolError:
    return ToolError(str(exc))


def register(bind_tool: Any, runtime: Any) -> dict[str, Any]:
    """Register assembly tools with the FastMCP server."""

    def bulk_find_replace(
        path: str,
        replacements: dict[str, str],
        confirm: bool = False,
    ) -> dict[str, Any]:
        """Replace multiple {{TOKEN}} placeholders in a Word document in a single pass.

        Requires WORD_ENABLE_WRITE=true and confirm=True. Returns replaced count,
        per-token results, and any tokens not found in the document.
        """
        try:
            return _asm.bulk_find_replace(path, replacements, confirm)
        except (ValidationError, NotAllowedError) as exc:
            raise _tool_error(exc) from exc

    def manage_hyperlinks(
        path: str,
        hyperlinks: dict[str, dict[str, Any]],
        confirm: bool = False,
    ) -> dict[str, Any]:
        """Update hyperlink URLs and labels for named display-text entries in a Word document.

        Requires WORD_ENABLE_WRITE=true and confirm=True. Only http, https, and mailto
        URL schemes are accepted. Returns update status per display-text key.
        """
        try:
            return _asm.manage_hyperlinks(path, hyperlinks, confirm)
        except (ValidationError, NotAllowedError) as exc:
            raise _tool_error(exc) from exc

    def verify_no_placeholders(
        path: str,
        ignore: list[str] | None = None,
    ) -> dict[str, Any]:
        """Verify that no unreplaced {{TOKEN}} placeholders remain in a Word document.

        Read-only operation. Returns pass/fail status and list of any residual tokens
        found in body paragraphs and table cells. Headers and footers are not scanned.
        """
        try:
            return _asm.verify_no_placeholders(path, ignore)
        except (ValidationError, NotAllowedError) as exc:
            raise _tool_error(exc) from exc

    return {
        "bulk_find_replace": bind_tool(bulk_find_replace, "bulk_find_replace"),
        "manage_hyperlinks": bind_tool(manage_hyperlinks, "manage_hyperlinks"),
        "verify_no_placeholders": bind_tool(verify_no_placeholders, "verify_no_placeholders"),
    }
