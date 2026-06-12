"""
FastMCP handler for manage_comments (list / add / delete).

No COM logic, no python-pptx / lxml imports — delegates to _pptx_comments.
"""
from __future__ import annotations

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

from pptmcp._pptx_comments import (
    _DEFAULT_X_EMU,
    _DEFAULT_Y_EMU,
    add_comment,
    delete_comment,
    list_comments,
)
from pptmcp.presentation_pptx import PPTMCPError


def register_comments_tools(mcp: FastMCP) -> None:
    """Register manage_comments on the given FastMCP instance."""

    @mcp.tool()
    def manage_comments(
        path: str,
        operation: str,
        slide_index: int | None = None,
        comment_id: int | None = None,
        text: str | None = None,
        author: str = "MCP",
        slide_x_emu: int = _DEFAULT_X_EMU,
        slide_y_emu: int = _DEFAULT_Y_EMU,
        confirm: bool = False,
    ) -> dict:
        """List, add, or delete comments on a PowerPoint slide.

        operation values:
          list   — list all comments on slide_index (or all slides if omitted).
                   Read-only; confirm not required.
          add    — add a comment. Requires PPT_ENABLE_WRITE=true, confirm=True,
                   slide_index, and text.
          delete — delete comment by comment_id. Requires PPT_ENABLE_WRITE=true,
                   confirm=True, slide_index, and comment_id.

        Returns:
          list   -> {comments, total_found, truncated, max_comments}
          add    -> {status, comment_id, author, slide_index}
          delete -> {status, comment_id, slide_index}
        """
        try:
            op = operation.strip().lower()

            if op == "list":
                return list_comments(path, slide_index=slide_index)

            if op == "add":
                if slide_index is None:
                    raise ValueError("slide_index is required for add operation")
                if not text:
                    raise ValueError("text is required for add operation")
                return add_comment(
                    path,
                    slide_index=slide_index,
                    text=text,
                    author=author,
                    slide_x_emu=slide_x_emu,
                    slide_y_emu=slide_y_emu,
                    confirm=confirm,
                )

            if op == "delete":
                if slide_index is None:
                    raise ValueError("slide_index is required for delete operation")
                if comment_id is None:
                    raise ValueError("comment_id is required for delete operation")
                return delete_comment(
                    path,
                    slide_index=slide_index,
                    comment_id=comment_id,
                    confirm=confirm,
                )

            raise ValueError(
                f"Unknown operation: {operation!r}. Valid values: list, add, delete"
            )

        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc
        except ValueError as exc:
            raise ToolError(str(exc)) from exc
