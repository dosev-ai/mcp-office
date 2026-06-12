"""
File-management and export MCP tool registrations for pptmcp.
Contains tools that save, copy, or mutate presentation files and slides
(save, create_presentation, copy_slide, clear_slide_content,
apply_slide_layout, remove_empty_placeholders, insert_image,
replace_slide_text, manage_hyperlinks, add_hyperlink).
Note: the unified export() tool is registered via register_read_tools()
in _server_handlers_read.py, not here.
All mutating tools require PPT_ENABLE_WRITE=true and confirm=True.
"""
from __future__ import annotations

__all__: list[str] = []

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

from pptmcp import presentation_pptx as prs
from pptmcp import shapes_pptx as shapes
from pptmcp._presentation_cache import _atomic_save, _evict_prs, _load_prs
from pptmcp.presentation_pptx import PPTMCPError


def register_export_tools(mcp: FastMCP) -> None:
    """Register file-management and export tools on the given FastMCP instance."""

    @mcp.tool()
    def save(path: str, confirm: bool = False) -> dict:
        """Persist the in-memory presentation to disk. Requires PPT_ENABLE_WRITE=true and confirm=True."""
        try:
            return prs.save(path, confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def create_presentation(
        path: str,
        title: str | None = None,
        template_path: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """Create a brand-new .pptx file at path. Optionally add a title slide.
        Optionally initialise with slide masters/theme from a template .pptx.

        Requires PPT_ENABLE_WRITE=true and confirm=True.
        Fails if file already exists - use a new destination path.
        """
        try:
            return prs.create_presentation(path, title, template_path, confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def manage_hyperlinks(
        path: str,
        slide_index: int,
        operation: str,
        shape_id: int | None = None,
        target_url: str | None = None,
        tooltip: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """List, add, or remove hyperlinks on a slide's shapes.

        operation: 'list' (read-only) | 'add' | 'remove'
        add/remove require PPT_ENABLE_WRITE=true and confirm=True.
        """
        try:
            return shapes.manage_hyperlinks(
                path=path,
                slide_index=slide_index,
                operation=operation,
                shape_id=shape_id,
                target_url=target_url,
                tooltip=tooltip,
                confirm=confirm,
            )
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def copy_slide(
        source_path: str | None = None,
        source_slide_index: int | None = None,
        target_path: str | None = None,
        target_slide_index: int | None = None,
        confirm: bool = False,
        path: str | None = None,
        slide_index: int | None = None,
        clear_content: bool = False,
    ) -> dict:
        """Copy a slide from a source presentation into a target presentation.

        Param aliases (any of these work):
          source_path  OR  path          - source .pptx file
          source_slide_index  OR  slide_index  - 0-based slide to copy
          target_path  (default: same as source - within-deck copy)

        target_slide_index: optional 0-based position in target deck (default: append at end).
        clear_content: when True, clear all placeholder text on the newly copied slide
          (empties every text frame's runs and paragraphs) so the slide is a blank
          layout template.  The structural copy (shapes, layout) is preserved; only
          text content is erased.
        Requires PPT_ENABLE_WRITE=true + confirm=True.
        """
        try:
            result = prs.copy_slide(
                source_path, source_slide_index, target_path, target_slide_index, confirm,
                path=path, slide_index=slide_index,
            )
            if clear_content:
                dest = result["target_path"]
                new_idx = result["new_slide_index"]
                resolved = prs._check_path(dest)
                deck = _load_prs(resolved)
                slide = deck.slides[new_idx]
                cleared = 0
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                run.text = ""
                        cleared += 1
                _atomic_save(deck, resolved)
                _evict_prs(resolved)
                result["cleared_placeholders"] = cleared
            return result
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def add_hyperlink(
        path: str,
        slide_index: int,
        shape_id: int,
        run_index: int = 0,
        url: str = "",
        display_text: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """Apply a URL hyperlink to a specific text run in a shape's first paragraph.

        Only http://, https://, and mailto: URLs are accepted.
        Requires confirm=True to execute.
        """
        try:
            return shapes.add_hyperlink(path, slide_index, shape_id, run_index, url, display_text, confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def replace_slide_text(
        path: str,
        find: str,
        replace: str,
        confirm: bool = False,
    ) -> dict:
        """Bulk find-and-replace across all slides. Handles <a:br> line-break elements.

        Requires PPT_ENABLE_WRITE=true and confirm=True.
        """
        try:
            return shapes.replace_slide_text(path, find, replace, confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def insert_image(
        path: str,
        slide_index: int,
        image_path: str,
        left: float = 1.0,
        top: float = 1.0,
        width: float | None = None,
        height: float | None = None,
        confirm: bool = False,
    ) -> dict:
        """Insert an allowlisted image onto a slide. Requires PPT_ENABLE_WRITE=true and confirm=True."""
        try:
            return shapes.insert_image(path, slide_index, image_path, left, top, width, height, confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def clear_slide_content(
        path: str,
        slide_index: int,
        confirm: bool = False,
    ) -> dict:
        """Remove ALL shapes from slide ``slide_index`` in ``path``. confirm=True required.

        Uses snapshot deletion — safe against index corruption.
        Requires PPT_ENABLE_WRITE=true and confirm=True.
        """
        try:
            return prs.clear_slide_content(path, slide_index, confirm=confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def apply_slide_layout(
        path: str,
        slide_index: int,
        layout_index: int,
        remove_placeholders: bool = True,
        confirm: bool = False,
    ) -> dict:
        """Apply layout ``layout_index`` to slide ``slide_index``. confirm=True required.

        Swaps the OPC SLIDE_LAYOUT relationship directly (python-pptx has no setter).
        If ``remove_placeholders=True`` (default), existing placeholder shapes are
        cleared so the slide can be re-populated from scratch.
        Requires PPT_ENABLE_WRITE=true and confirm=True.
        """
        try:
            return prs.apply_slide_layout(path, slide_index, layout_index, remove_placeholders, confirm=confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def remove_empty_placeholders(
        path: str,
        slide_index: int,
        confirm: bool = False,
    ) -> dict:
        """Remove empty / default-text placeholder shapes from slide ``slide_index``. confirm=True required.

        Removes placeholders whose text is blank or matches well-known PowerPoint
        default prompts ('Click to add title', etc.).  Non-placeholder shapes are
        never touched.
        Requires PPT_ENABLE_WRITE=true and confirm=True.
        """
        try:
            return shapes.remove_empty_placeholders(path, slide_index, confirm=confirm)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e
