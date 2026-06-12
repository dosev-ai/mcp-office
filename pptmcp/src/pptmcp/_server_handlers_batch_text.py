"""FastMCP handler for batch_set_text. No python-pptx or pptx imports."""
from __future__ import annotations

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

from pptmcp.presentation_pptx import PPTMCPError


def register_batch_text_tools(mcp: FastMCP) -> None:
    """Register batch_set_text on the given FastMCP instance."""

    @mcp.tool()
    def batch_set_text(
        path: str,
        slide_index: int,
        updates: list[dict],
        confirm: bool = False,
    ) -> dict:
        """Update text on multiple shapes in a single slide in one open/save cycle.

        Each entry in `updates` must be {"shape_id": int, "text": str}.
        Shapes that are missing or have no text frame are collected in `skipped`
        rather than aborting the batch.
        Requires PPT_ENABLE_WRITE=true and confirm=True.

        Returns: {updated, path, slide_index, results[{shape_id, status, previous_text, new_text}],
                  skipped[{shape_id, reason}], error_count}
        """
        try:
            from pptmcp._pptx_batch_text import batch_set_text as _impl  # noqa: PLC0415
            return _impl(path, slide_index, updates, confirm=confirm)
        except (PPTMCPError, ValueError) as exc:
            raise ToolError(str(exc)) from exc
