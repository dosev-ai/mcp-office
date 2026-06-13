"""Compatibility shim for legacy imports of pptmcp._server_handlers_com.

FastMCP tool registration moved to pptmcp._server_handlers_com_tools so that
MCP wiring no longer lives in a *_com.py module. This file re-exports the
public symbols to preserve existing imports during the transition.
"""
from __future__ import annotations

from . import _server_handlers_com_tools as _tools

_COM_AVAILABLE = _tools._COM_AVAILABLE
_pcom = _tools._pcom


def _sync_impl_state() -> None:
    _tools._COM_AVAILABLE = _COM_AVAILABLE
    _tools._pcom = _pcom


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
    _sync_impl_state()
    return _tools.export_slide(
        file_path=file_path,
        slide_index=slide_index,
        fmt=fmt,
        output_path=output_path,
        dpi=dpi,
        width_px=width_px,
        height_px=height_px,
        confirm=confirm,
    )


def export_deck(
    file_path: str,
    fmt: str,
    output_path: str,
    dpi: int | None = None,
    confirm: bool = False,
) -> dict:
    _sync_impl_state()
    return _tools.export_deck(
        file_path=file_path,
        fmt=fmt,
        output_path=output_path,
        dpi=dpi,
        confirm=confirm,
    )


def export_slides_to_stamped_dir(
    path: str,
    base_dir: str,
    slide_indices: list[int] | None = None,
    width_px: int = 0,
    height_px: int = 0,
    confirm: bool = False,
) -> dict:
    _sync_impl_state()
    return _tools.export_slides_to_stamped_dir(
        path=path,
        base_dir=base_dir,
        slide_indices=slide_indices,
        width_px=width_px,
        height_px=height_px,
        confirm=confirm,
    )


def export_changed_slides_only(
    path: str,
    base_dir: str,
    previous_hashes: dict,
    width_px: int = 0,
    height_px: int = 0,
    confirm: bool = False,
) -> dict:
    _sync_impl_state()
    return _tools.export_changed_slides_only(
        path=path,
        base_dir=base_dir,
        previous_hashes=previous_hashes,
        width_px=width_px,
        height_px=height_px,
        confirm=confirm,
    )


def register_com_tools(mcp) -> None:
    _sync_impl_state()
    _tools.register_com_tools(mcp)

__all__ = [
    "_COM_AVAILABLE",
    "_pcom",
    "export_changed_slides_only",
    "export_deck",
    "export_slide",
    "export_slides_to_stamped_dir",
    "register_com_tools",
]
