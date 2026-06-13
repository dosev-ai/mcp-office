"""
Read-only MCP tool registrations for pptmcp.
All tools in this module are read-only and do not require PPT_ENABLE_WRITE or confirm=True.
"""
from __future__ import annotations

__all__: list[str] = []
from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from mcpshared.inline_artifact_response import build_inline_artifact_response

from pptmcp import content_pptx as content
from pptmcp import contract_pptx as contract
from pptmcp import presentation_pptx as prs
from pptmcp import shapes_pptx as shapes
from pptmcp.presentation_pptx import PPTMCPError


def export(  # noqa: A001
    scope: str,
    path: str,
    fmt: str = "png",
    slide_index: int | None = None,
    output_path: str | None = None,
    dpi: int | None = None,
    width_px: int = 0,
    height_px: int = 0,
    return_inline: bool = False,
    response_mode: str = "auto",
    max_inline_bytes: int = 3_000_000,
    options: dict | None = None,
    confirm: bool = False,
) -> dict | list:
    """Unified export dispatcher. Read-only scopes need no write gate; COM scopes require confirm=True.

    scope: 'slide_text' | 'slide_images' | 'slide_png' | 'deck_pdf' | 'slide'
      'slide_text'   - Extract text from one slide (slide_index) or all slides (None). Read-only.
      'slide_images' - List all embedded images in the deck with metadata. Read-only.
      'slide_png'    - Export a single slide as PNG via COM (requires PPT_ENABLE_COM=true).
                       Params: slide_index (required), output_path (required), dpi (optional),
                       width_px, height_px. Requires confirm=True.
                       DEPRECATED SCOPE: "slide_png" is deprecated. Use scope="slide" with return_inline=False for equivalent raw file output.
      'deck_pdf'     - Export the full deck as PDF via COM (requires PPT_ENABLE_COM=true).
                           Params: output_path (required), dpi (optional). Requires confirm=True.
      'slide'        - Consolidated visual export. Currently supports fmt='png'.
                       Params: output_path (required), options.slide_index (required),
                       options.inline_dpi (optional). return_inline=True packages PNG via mcpshared.

    Deprecated per scope:
      slide_text   -> was export_slide_as_text
      slide_images -> was extract_images
      slide_png    -> was export_slide (COM tool)
      deck_pdf     -> was export_deck (COM tool)
    """
    try:
        export_options = options or {}
        if scope == "slide_text":
            _reject_options(export_options, scope=scope, allowed=set())
            return content.export_slide_as_text(path, slide_index)
        elif scope == "slide_images":
            _reject_options(export_options, scope=scope, allowed=set())
            return content.extract_images(path)
        elif scope in {"slide_png", "slide"}:
            # both scopes accept the same options dict so
            # callers can use `options={"slide_index": N}` consistently.
            _reject_options(
                export_options, scope=scope, allowed={"slide_index", "inline_dpi"}
            )
            if scope == "slide" and fmt.lower() != "png":
                raise PPTMCPError("scope='slide' currently supports fmt='png' only")
            option_slide_index = export_options.get("slide_index")
            if option_slide_index is not None:
                slide_index = option_slide_index
            option_dpi = export_options.get("inline_dpi")
            if option_dpi is not None:
                dpi = option_dpi
            if scope == "slide_png":
                fmt = "png"
            if slide_index is None:
                raise PPTMCPError(f"slide_index is required for scope={scope!r}")
            if output_path is None:
                raise PPTMCPError(f"output_path is required for scope={scope!r}")
            # Import COM tools only when the COM export path is actually needed.
            from pptmcp import _server_handlers_com_tools as _com_tools  # noqa: PLC0415
            result = _com_tools.export_slide(
                file_path=path,
                slide_index=slide_index,
                fmt="png",
                output_path=output_path,
                dpi=dpi,
                width_px=width_px,
                height_px=height_px,
                confirm=confirm,
            )
            if scope != "slide":
                return result
            if result.get("success") is False:
                return result
            inline = build_inline_artifact_response(
                artifact_type="slide_render",
                mime_type="image/png",
                output_path=result.get("output_path", output_path),
                return_inline=return_inline,
                response_mode=response_mode,
                max_inline_bytes=max_inline_bytes,
                metadata={
                    "scope": scope,
                    "fmt": "png",
                    "slide_index": slide_index,
                    "dpi": dpi,
                },
            )
            return {**result, **inline}
        elif scope == "deck_pdf":
            _reject_options(export_options, scope=scope, allowed=set())
            if output_path is None:
                raise PPTMCPError("output_path is required for scope='deck_pdf'")
            # Import COM tools only when the COM export path is actually needed.
            from pptmcp import _server_handlers_com_tools as _com_tools  # noqa: PLC0415
            return _com_tools.export_deck(
                file_path=path,
                fmt="pdf",
                output_path=output_path,
                dpi=dpi,
                confirm=confirm,
            )
        else:
            raise PPTMCPError(
                f"Unknown scope: {scope!r}. "
                "Must be one of: slide_text, slide_images, slide_png, deck_pdf, slide"
            )
    except ToolError:
        raise
    except PPTMCPError as e:
        raise ToolError(str(e)) from e
    except Exception as e:
        raise ToolError(str(e)) from e


def _reject_options(options: dict, *, scope: str, allowed: set[str]) -> None:
    unsupported = set(options) - allowed
    if unsupported:
        raise PPTMCPError(
            f"Unsupported options for scope={scope!r}: {sorted(unsupported)}. "
            f"Allowed: {sorted(allowed)}"
        )


def register_read_tools(mcp: FastMCP) -> None:  # noqa: C901
    """Register read-only tools and the unified export() dispatcher on the given FastMCP instance.

    Despite its name, this function also registers export(), which handles both
    read-only scopes (slide_text, slide_images) and write/COM scopes (slide_png,
    deck_pdf, slide). The export() tool is placed here because it is the primary
    read/inspect entry point for agents; COM-gated paths inside export() require
    confirm=True at call time.
    """

    @mcp.tool()
    def capabilities() -> dict:
        """Return metadata about this MCP phase and its available tools."""
        try:
            result = prs.capabilities()
            result["tools"] = result["tools"] + contract.CONTRACT_TOOLS
            return result
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def read_presentation(path: str) -> dict:
        """Return slide count and per-slide summary for the presentation at path."""
        try:
            return prs.read_presentation(path)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def get_presentation_metadata(path: str) -> dict:
        """Return title, author, slide count, created/modified for the presentation."""
        try:
            return prs.get_presentation_metadata(path)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def list_layouts(path: str) -> list[dict]:
        """Return all slide layouts in a presentation with index, name, and placeholder info.

        Use the returned indices with add_slide(layout_index=N) to pick the right layout.
        """
        try:
            return prs.list_layouts(path)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def list_slides(path: str) -> list[dict]:
        """Return index, title, layout and has_notes for every slide."""
        try:
            return prs.list_slides(path)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def read_slide(path: str, slide_index: int) -> dict:
        """Return full shape list and notes for a single slide."""
        try:
            return prs.read_slide(path, slide_index)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def read_speaker_notes(path: str, slide_index: int | None = None) -> list[dict]:
        """Return speaker notes for one slide (slide_index) or all slides (None)."""
        try:
            return prs.read_speaker_notes(path, slide_index)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def list_shapes(path: str, slide_index: int) -> list[dict]:
        """Return all shapes on a slide with type, position, and size."""
        try:
            return prs.list_shapes(path, slide_index)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def get_shape(path: str, slide_index: int, shape_id: int) -> dict:
        """Return full detail for a specific shape by shape_id."""
        try:
            return prs.get_shape(path, slide_index, shape_id)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def extract_tables(path: str, slide_index: int | None = None) -> dict:
        """Extract all tables from one slide or the whole deck as row/cell arrays.

        Returns ``{"tables": [...]}``.
        """
        try:
            return content.extract_tables(path, slide_index)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def extract_images(path: str) -> dict:
        """List all embedded images across the deck with dimensions and content type.

        Returns ``{"images": [...]}``.
        """
        try:
            return content.extract_images(path)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def export_slide_as_text(path: str, slide_index: int | None = None) -> list[dict]:
        """Extract all text from shapes on one slide or the entire deck.

        Returns ``list[dict]`` — one entry per slide, each with:
        - ``slide_index`` (int): 0-based slide position
        - ``texts`` (list[str]): non-empty text strings from all shapes on that slide
        """
        try:
            return content.export_slide_as_text(path, slide_index)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def extract_presentation_text(path: str) -> dict:
        """Export full deck text with per-slide structure for RAG indexing.

        Returns title, body text, notes, and per-shape text for every slide in one
        structured response. Read-only — no write gate required.
        """
        try:
            return content.extract_presentation_text(path)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    @mcp.tool()
    def detect_overlapping_shapes(path: str, slide_index: int) -> dict:
        """Return all overlapping shape pairs on a slide with overlap area in square inches.
        Read-only — does not require PPT_ENABLE_WRITE or confirm=True.
        Returns: {slide_index, overlapping_pairs: [{shape_a, shape_b, overlap_area_in2}], pair_count}.
        shape_a and shape_b are {shape_id, name} dicts.
        """
        try:
            return shapes.detect_overlapping_shapes(path, slide_index)
        except PPTMCPError as e:
            raise ToolError(str(e)) from e

    mcp.tool()(export)

