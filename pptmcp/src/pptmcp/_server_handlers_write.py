"""
Write (mutation) MCP tool registrations for pptmcp.
All tools require PPT_ENABLE_WRITE=true and confirm=True.

Module-level add_content and set_format remain importable for backwards-compat
(server.py re-exports them; test_dispatcher_tools.py imports them directly).
MCP registration uses inner @mcp.tool() functions inside register_write_tools().
"""
from __future__ import annotations

__all__: list[str] = []

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

from pptmcp import presentation_pptx as prs
from pptmcp import shapes_pptx as shapes
from pptmcp.presentation_pptx import PPTMCPError

# Re-export shape ops for backwards compatibility — implementations live in _server_handlers_shape_ops
from pptmcp._server_handlers_shape_ops import (  # noqa: F401
    add_content,
    set_format,
    set_table_style,
    shape,
)


def slide(
    operation: str,
    path: str,
    confirm: bool = False,
    layout_index: int = 1,
    title: str | None = None,
    layout_name: str | None = None,
    slide_index: int | None = None,
    new_order: list[int] | None = None,
    source_path: str | None = None,
    source_slide_index: int | None = None,
    target_path: str | None = None,
    target_slide_index: int | None = None,
    suppress_content_placeholder: bool = True,
) -> dict:
    """Unified slide mutation dispatcher. Requires PPT_ENABLE_WRITE=true + confirm=True.

    operation: 'add' | 'delete' | 'reorder' | 'copy'
      'add'     - params: layout_index (default 1), title (optional),
                  layout_name (optional, takes precedence over layout_index),
                  suppress_content_placeholder (default True)
      'delete'  - params: slide_index (required, 0-based)
      'reorder' - params: new_order (required, list of all slide indices as a permutation)
      'copy'    - params: source_slide_index (or slide_index alias),
                  source_path (defaults to path), target_path (optional, defaults to path),
                  target_slide_index (optional)
    """
    try:
        if operation == "add":
            return prs.add_slide(
                path,
                layout_index,
                title,
                confirm,
                layout_name,
                suppress_content_placeholder=suppress_content_placeholder,
            )
        elif operation == "delete":
            if slide_index is None:
                raise PPTMCPError("slide_index is required for operation='delete'")
            return prs.delete_slide(path, slide_index, confirm)
        elif operation == "reorder":
            if new_order is None:
                raise PPTMCPError("new_order is required for operation='reorder'")
            return prs.reorder_slides(path, new_order, confirm)
        elif operation == "copy":
            effective_source = source_path if source_path is not None else path
            effective_slide_idx = (
                source_slide_index if source_slide_index is not None else slide_index
            )
            return prs.copy_slide(
                source_path=effective_source,
                source_slide_index=effective_slide_idx,
                target_path=target_path,
                target_slide_index=target_slide_index,
                confirm=confirm,
            )
        else:
            raise PPTMCPError(
                f"Unknown operation: {operation!r}. Must be one of: add, delete, reorder, copy"
            )
    except ToolError:
        raise
    except PPTMCPError as e:
        raise ToolError(str(e)) from e
    except Exception as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# MCP tool registration
# ---------------------------------------------------------------------------

def register_write_tools(mcp: FastMCP) -> None:
    """Register the 10 core mutation tools on the given FastMCP instance."""

    @mcp.tool()
    def add_slide(
        path: str,
        layout_index: int = 1,
        title: str | None = None,
        confirm: bool = False,
        layout_name: str | None = None,
        suppress_content_placeholder: bool = False,
    ) -> dict:
        """Append a new slide. layout_name takes precedence over layout_index when both are given. Requires PPT_ENABLE_WRITE=true and confirm=True."""
        try:
            return prs.add_slide(
                path,
                layout_index,
                title,
                confirm,
                layout_name,
                suppress_content_placeholder=suppress_content_placeholder,
            )
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def edit_text_placeholder(
        path: str,
        slide_index: int,
        placeholder_idx: int = 0,
        text: str = "",
        confirm: bool = False,
        shape_id: int | None = None,
    ) -> dict:
        """Replace text in a placeholder by index, or in any text-containing shape by shape_id.

        If shape_id is provided: finds shape by shape_id and replaces its text (works for text
        boxes, content shapes, etc. - not just native placeholders).
        If shape_id is NOT provided: uses placeholder_idx to find placeholder by
        placeholder_format.idx (layout - NOT list position).
        Requires PPT_ENABLE_WRITE=true and confirm=True.
        """
        try:
            return shapes.edit_text_placeholder(
                path, slide_index, placeholder_idx, text, confirm, shape_id=shape_id
            )
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def set_speaker_notes(
        path: str,
        slide_index: int,
        notes_text: str | None = None,
        confirm: bool = False,
        notes: str | None = None,
    ) -> dict:
        """Write speaker notes for a slide. Requires PPT_ENABLE_WRITE=true and confirm=True.

        Accepts notes_text (primary, backward-compat) or notes (preferred alias).
        Changes are held in memory - call save(path, confirm=True) to persist.
        """
        try:
            return shapes.set_speaker_notes(path, slide_index, notes_text, confirm, notes=notes)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def delete_shape(
        path: str,
        slide_index: int,
        shape_id: int,
        confirm: bool = False,
    ) -> dict:
        """Remove a shape from a slide by shape_id. Requires PPT_ENABLE_WRITE=true and confirm=True."""
        try:
            return shapes.delete_shape(path, slide_index, shape_id, confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def reorder_slides(path: str, new_order: list[int], confirm: bool = False) -> dict:
        """Reorder slides to new_order permutation. Requires PPT_ENABLE_WRITE=true and confirm=True."""
        try:
            return prs.reorder_slides(path, new_order, confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def delete_slide(path: str, slide_index: int, confirm: bool = False) -> dict:
        """Remove a slide by index. Requires PPT_ENABLE_WRITE=true and confirm=True."""
        try:
            return prs.delete_slide(path, slide_index, confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    mcp.tool()(add_content)
    mcp.tool()(set_format)
    mcp.tool()(set_table_style)
    mcp.tool()(slide)
    mcp.tool()(shape)
