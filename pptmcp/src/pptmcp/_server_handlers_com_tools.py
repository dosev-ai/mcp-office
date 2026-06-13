"""
FastMCP tool registrations for pptmcp COM-backed PowerPoint tools.
Windows-only tools that require a live PowerPoint COM session.
Module-level functions are re-exported via server.py for backward compatibility.
"""
from __future__ import annotations

import importlib
__all__: list[str] = []

import sys
from typing import Any

from fastmcp.exceptions import ToolError

from pptmcp.presentation_pptx import PPTMCPError

# ── COM availability check ──────────────────────────────────────────────────
def _load_com_backend() -> tuple[Any | None, bool]:
    if sys.platform != "win32":
        return None, False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return None, False
    pcom = importlib.import_module("pptmcp.presentation_com")
    return pcom, True


_pcom, _COM_AVAILABLE = _load_com_backend()


def _require_pcom() -> Any:
    if _pcom is None:
        raise RuntimeError("COM backend not available")
    return _pcom


# ---------------------------------------------------------------------------
# Module-level COM tool functions (always-registered; graceful on non-COM)
# ---------------------------------------------------------------------------

def export_slide(
    file_path: str,
    slide_index: int,
    fmt: str,
    output_path: str,
    dpi: int | None = None,
    width_px: int = 0,
    height_px: int = 0,
    confirm: bool = False,
) -> dict:
    """Export a single slide as PNG, PDF, or SVG.

    Renders one slide to a file on disk using a COM-based PowerPoint process.

    Args:
        file_path: Absolute path to the .pptx source file.
        slide_index: 0-based slide index (0 = first slide). Matches all other pptmcp tools.
        fmt: Export format - one of 'png', 'pdf', 'svg'.
             SVG requires Office 2019+ / M365 Build 13328+.
        output_path: Absolute path to the output file (must match the format extension).
        dpi: Dots-per-inch for raster formats (PNG). Clamped to 72-600. Ignored for SVG.
        confirm: Must be True to execute the export (write guard).

    Returns:
        dict with success, output_path, format, slide_number, or error message.
    """
    if not _COM_AVAILABLE:
        return {"success": False, "error": "COM not available on this platform"}
    pcom = _require_pcom()
    try:
        return pcom.ppt_export_slide(
            file_path=file_path,
            slide_index=slide_index,
            fmt=fmt,
            output_path=output_path,
            dpi=dpi,
            width_px=width_px,
            height_px=height_px,
            confirm=confirm,
        )
    except PPTMCPError as e:
        raise ToolError(str(e)) from e


def export_deck(
    file_path: str,
    fmt: str,
    output_path: str,
    dpi: int | None = None,
    confirm: bool = False,
) -> dict:
    """Export an entire presentation as PDF, PPTX, or a folder of PNG images.

    Args:
        file_path: Absolute path to the .pptx source file.
        fmt: Export format - one of 'pdf', 'pptx', 'images'.
             'pdf' -> single PDF file.
             'pptx' -> saves a copy as .pptx (cannot overwrite source).
             'images' -> exports every slide as Slide1.PNG, Slide2.PNG into output_path dir.
        output_path: Absolute path to output file (pdf/pptx) or output directory (images).
        dpi: DPI for 'images' format, clamped 72-600. Ignored for pdf/pptx.
        confirm: Must be True to execute the export (write guard).

    Returns:
        dict with success, output_path, format, or error message.
    """
    if not _COM_AVAILABLE:
        return {"success": False, "error": "COM not available on this platform"}
    pcom = _require_pcom()
    try:
        return pcom.ppt_export_deck(
            file_path=file_path,
            fmt=fmt,
            output_path=output_path,
            dpi=dpi,
            confirm=confirm,
        )
    except PPTMCPError as e:
        raise ToolError(str(e)) from e


def export_slides_to_stamped_dir(
    path: str,
    base_dir: str,
    slide_indices: list[int] | None = None,
    width_px: int = 0,
    height_px: int = 0,
    confirm: bool = False,
) -> dict:
    """Export selected slides as PNG images to a timestamped directory.

    Creates <base_dir>/exports/run-<UTC-timestamp>/png/slide-NNN.png for each
    selected slide. Returns the run directory path and list of exported file paths.

    Args:
        path: Path to the .pptx file.
        base_dir: Root directory for exports (must be in PPT_ALLOWLIST_ROOTS).
        slide_indices: 0-based indices of slides to export. None = all slides.
        width_px: PNG output width in pixels (0 = PowerPoint default).
        height_px: PNG output height in pixels (0 = PowerPoint default).
        confirm: Must be True to execute (write-operation guard).

    Returns:
        run_dir: Path to the stamped export directory.
        exported_slides: List of exported PNG file paths.
        slide_count: Number of slides exported.
    """
    if not _COM_AVAILABLE:
        return {"success": False, "error": "COM not available on this platform"}
    pcom = _require_pcom()
    try:
        return pcom.export_slides_to_stamped_dir(
            path=path,
            base_dir=base_dir,
            slide_indices=slide_indices,
            width_px=width_px,
            height_px=height_px,
            confirm=confirm,
        )
    except PPTMCPError as e:
        raise ToolError(str(e)) from e


def export_changed_slides_only(
    path: str,
    base_dir: str,
    previous_hashes: dict,
    width_px: int = 0,
    height_px: int = 0,
    confirm: bool = False,
) -> dict:
    """Re-export only slides whose content changed since previous_hashes snapshot.

    previous_hashes: dict {str(slide_index): md5_hex_32_chars}.
    Requires PPT_ENABLE_COM=true and PPT_ENABLE_WRITE=true.
    Returns: {run_dir, exported_slides, unchanged_slides, new_hashes}
    """
    if not _COM_AVAILABLE:
        return {"success": False, "error": "COM not available on this platform"}
    pcom = _require_pcom()
    try:
        return pcom.export_changed_slides_only(
            path, base_dir, previous_hashes, width_px, height_px, confirm
        )
    except PPTMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# MCP tool registration
# ---------------------------------------------------------------------------

def register_com_tools(mcp) -> None:
    """Register COM export tools on the given FastMCP instance.

    export_slide, export_deck, export_slides_to_stamped_dir, and
    export_changed_slides_only are always registered (they return a
    graceful error dict when COM is unavailable).

    run_slide_show and recalculate_charts are only registered when
    _COM_AVAILABLE is True (Windows + win32com present).
    """
    mcp.tool()(export_slide)
    mcp.tool()(export_deck)
    mcp.tool()(export_slides_to_stamped_dir)
    mcp.tool()(export_changed_slides_only)

    if _COM_AVAILABLE:
        pcom = _require_pcom()

        @mcp.tool()
        def run_slide_show(path: str, confirm: bool = False) -> dict:
            """Open the presentation in slide-show mode in PowerPoint (requires PPT_ENABLE_COM=true, PPT_ENABLE_WRITE=true, Windows only)."""
            try:
                return pcom.run_slide_show(path, confirm=confirm)
            except PPTMCPError as e:
                raise ToolError(str(e)) from e

        @mcp.tool()
        def recalculate_charts(
            path: str,
            confirm: bool = False,
        ) -> dict:
            """Activate embedded charts for recalculation only; linked and external refresh is skipped for safety. Requires PPT_ENABLE_COM=true, PPT_ENABLE_WRITE=true, and confirm=True. Windows only."""
            try:
                return pcom.recalculate_charts(path, confirm)
            except PPTMCPError as e:
                raise ToolError(str(e)) from e
