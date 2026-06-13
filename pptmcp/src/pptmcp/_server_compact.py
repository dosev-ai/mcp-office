"""Compact agent-facing MCP tool registration for pptmcp.

This mode keeps the implementation surface unchanged, but exposes a smaller
tool list for agents that should prefer dispatchers over many narrow tools.
"""
from __future__ import annotations

from typing import Literal

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

import pptmcp.review_pptx as review
from pptmcp import content_pptx as content
from pptmcp import contract_pptx as contract
from pptmcp import presentation_pptx as prs
from pptmcp import shapes_pptx as shapes
from pptmcp._server_handlers_read import export
from pptmcp._server_handlers_write import add_content, set_format, slide, shape
from pptmcp.presentation_pptx import PPTMCPError

COMPACT_TOOL_SURFACE_ENV = "PPT_COMPACT_TOOL_SURFACE"

COMPACT_TOOL_NAMES: tuple[str, ...] = (
    "capabilities",
    "read_presentation",
    "list_layouts",
    "list_slides",
    "read_slide",
    "list_shapes",
    "get_shape",
    "extract_tables",
    "export_slide_as_text",
    "extract_presentation_text",
    "detect_overlapping_shapes",
    "declare_slide_contract",
    "validate_contract",
    "check_presentation_against_contract",
    "review_slide_export",
    "produce_evidence_bundle",
    "get_presentation_context",
    "add_content",  # deprecated_alias of shape — kept for 2-release transition window
    "set_format",
    "slide",
    "shape",
    "export",
    "manage_comments",
    "batch_set_text",
    "pptmcp_add_column",
    "pptmcp_edit_table_cell",
    "pptmcp_set_column_width",
)


def compact_surface_enabled() -> bool:
    """Return True when the server should expose the compact agent tool list."""
    import os

    return os.environ.get(COMPACT_TOOL_SURFACE_ENV, "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }


def register_compact_tools(mcp: FastMCP) -> None:
    """Register the compact dispatcher-first tool surface."""

    @mcp.tool()
    def capabilities() -> dict:
        """Return metadata about this MCP phase and its available tools."""
        try:
            return prs.capabilities()
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def read_presentation(path: str) -> dict:
        """Return slide count and per-slide summary for the presentation at path."""
        try:
            return prs.read_presentation(path)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def list_layouts(path: str) -> list[dict]:
        """Return all slide layouts with indices, names, and placeholder info."""
        try:
            return prs.list_layouts(path)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def list_slides(path: str) -> list[dict]:
        """Return index, title, layout, and has_notes for every slide."""
        try:
            return prs.list_slides(path)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def read_slide(path: str, slide_index: int) -> dict:
        """Return full shape list and notes for a single slide."""
        try:
            return prs.read_slide(path, slide_index)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def list_shapes(path: str, slide_index: int) -> list[dict]:
        """Return all shapes on a slide with type, position, and size."""
        try:
            return prs.list_shapes(path, slide_index)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def get_shape(path: str, slide_index: int, shape_id: int) -> dict:
        """Return full detail for a specific shape by shape_id."""
        try:
            return prs.get_shape(path, slide_index, shape_id)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def extract_tables(path: str, slide_index: int | None = None) -> dict:
        """Extract all tables from one slide or the whole deck as row/cell arrays."""
        try:
            return content.extract_tables(path, slide_index)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def export_slide_as_text(path: str, slide_index: int | None = None) -> list[dict]:
        """Extract all text from shapes on one slide or the entire deck."""
        try:
            return content.export_slide_as_text(path, slide_index)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def extract_presentation_text(path: str) -> dict:
        """Export full deck text with per-slide structure for RAG indexing."""
        try:
            return content.extract_presentation_text(path)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def detect_overlapping_shapes(path: str, slide_index: int) -> dict:
        """Return all overlapping shape pairs on a slide."""
        try:
            return shapes.detect_overlapping_shapes(path, slide_index)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def declare_slide_contract(
        slide_path: str,
        spec: dict,
        confirm: bool = False,
    ) -> dict:
        """Persist a slide-export contract beside the PPTX (RCI-US-01)."""
        try:
            return contract.declare_slide_contract(slide_path, spec, confirm=confirm)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def validate_contract(contract_path: str) -> dict:
        """Validate a JSON Output Contract file against the pptmcp v1.0 schema."""
        try:
            return contract.validate_contract(contract_path)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def check_presentation_against_contract(path: str, contract_path: str) -> dict:
        """Check a presentation against its Output Contract specification."""
        try:
            return contract.check_presentation_against_contract(path, contract_path)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def review_slide_export(
        path: str,
        slide_index: int,
        png_path: str,
        contract_path: str | None = None,
    ) -> dict:
        """Review a slide export against quality criteria."""
        try:
            return review.review_slide_export(path, slide_index, png_path, contract_path)
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def produce_evidence_bundle(
        path: str,
        review_results: list,
        contract_path: str | None = None,
        exports_dir: str | None = None,
    ) -> dict:
        """Aggregate review results into a machine-verifiable evidence bundle."""
        try:
            return content.produce_evidence_bundle(
                path, review_results, contract_path, exports_dir
            )
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    @mcp.tool()
    def get_presentation_context(
        path: str,
        level: Literal["index", "focused", "deep"] = "focused",
        annotations: list[dict] | None = None,
    ) -> dict:
        """Return an Artifact Context Packet for a PowerPoint presentation."""
        from pptmcp._pptx_acp_adapter import build_presentation_context  # noqa: PLC0415

        try:
            return build_presentation_context(path, level=level, annotations=annotations)
        except ValueError as exc:
            raise ToolError(str(exc)) from exc
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc

    mcp.tool()(add_content)
    mcp.tool()(set_format)
    mcp.tool()(slide)
    mcp.tool()(shape)
    mcp.tool()(export)

