"""_shapes_links — hyperlink and placeholder removal functions for pptmcp."""
from __future__ import annotations

import logging
from urllib.parse import urlparse

from pptmcp._shapes_helpers import _is_empty_placeholder
from pptmcp.presentation_pptx import (
    ValidationError,
    _atomic_save,
    _audit_log,
    _check_confirm,
    _check_path,
    _check_write,
    _evict_prs,
    _find_shape_by_id,
    _load_prs,
    _validate_slide_index,
)

_log = logging.getLogger(__name__)


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

    Only http://, https://, and mailto: URLs are accepted (OWASP A03).
    Requires PPT_ENABLE_WRITE=true and confirm=True.
    """
    _check_write()
    _check_confirm(confirm)
    resolved = _check_path(path)
    url = url.strip()
    _AH_ALLOWED_SCHEMES = {"https", "http", "mailto"}
    scheme = url.lower().split("://")[0].split(":")[0].strip()
    if scheme not in _AH_ALLOWED_SCHEMES:
        raise ValidationError(
            f"URL scheme {scheme!r} is not allowed. Permitted: {sorted(_AH_ALLOWED_SCHEMES)}"
        )
    prs = _load_prs(resolved)
    slide = prs.slides[slide_index] if 0 <= slide_index < len(prs.slides) else None
    if slide is None:
        raise ValidationError(f"slide_index {slide_index} out of range")
    shape = _find_shape_by_id(slide, shape_id)
    if not shape.has_text_frame:
        raise ValidationError(f"Shape {shape_id} has no text frame")
    paragraph = shape.text_frame.paragraphs[0]
    if run_index < 0 or run_index >= len(paragraph.runs):
        raise ValidationError(f"run_index {run_index} out of range for paragraph shape")
    run = paragraph.runs[run_index]
    if display_text is not None:
        run.text = display_text
    run.hyperlink.address = url
    _audit_log("add_hyperlink", resolved, slide_index, {"shape_id": shape_id, "url": url})
    return {"path": str(resolved), "slide_index": slide_index, "shape_id": shape_id, "url": url}


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

    operation: 'list' (read-only), 'add' (requires write gate), 'remove' (requires write gate).
    """
    VALID_OPERATIONS = {"list", "add", "remove"}
    if operation not in VALID_OPERATIONS:
        raise ValidationError(f"operation must be one of {sorted(VALID_OPERATIONS)}, got {operation!r}")
    if operation in ("add", "remove"):
        _check_write()
        _check_confirm(confirm)
    if operation in ("add", "remove") and shape_id is None:
        raise ValidationError("shape_id is required for add/remove operations")
    if operation == "add" and not target_url:
        raise ValidationError("target_url is required for add operation")
    if operation == "add" and target_url:
        target_url = target_url.strip()  # H-013: strip before urlparse to avoid empty-scheme false-positive
        _ALLOWED_SCHEMES = {"https", "http", "mailto"}
        parsed = urlparse(target_url)
        scheme = parsed.scheme.lower().strip()
        if scheme not in _ALLOWED_SCHEMES:
            raise ValidationError(f"URL scheme {scheme!r} is not allowed. Permitted: {sorted(_ALLOWED_SCHEMES)}")
    resolved = _check_path(path)
    prs = _load_prs(resolved)
    if slide_index < 0 or slide_index >= len(prs.slides):
        raise ValidationError(f"slide_index {slide_index} out of range (0–{len(prs.slides) - 1})")
    slide = prs.slides[slide_index]

    if operation == "list":
        hyperlinks: list[dict] = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.hyperlink and run.hyperlink.address:
                        hyperlinks.append({
                            "shape_id": shape.shape_id, "shape_name": shape.name,
                            "run_text": run.text, "url": run.hyperlink.address,
                            "tooltip": getattr(run.hyperlink, "tooltip", "") or "",
                        })
        return {"slide_index": slide_index, "operation": "list",
                "hyperlink_count": len(hyperlinks), "links": hyperlinks}

    target_shape = None
    for shape in slide.shapes:
        if shape.shape_id == shape_id:
            target_shape = shape
            break
    if target_shape is None:
        raise ValidationError(f"Shape with shape_id={shape_id} not found on slide {slide_index}")
    if not target_shape.has_text_frame:
        raise ValidationError(f"Shape {shape_id} has no text frame — cannot manage hyperlinks")

    if operation == "add":
        added = 0
        for para in target_shape.text_frame.paragraphs:
            for run in para.runs:
                run.hyperlink.address = target_url.strip()
                if tooltip:
                    try:
                        run.hyperlink.tooltip = tooltip
                    except AttributeError:
                        pass
                added += 1
        if added == 0:
            raise ValidationError(f"Shape {shape_id} has no text runs to attach a hyperlink to")
        _audit_log("manage_hyperlinks_add", resolved, slide_index, {"shape_id": shape_id, "url": target_url})
        return {"ok": True, "slide_index": slide_index, "shape_id": shape_id,
                "operation": "add", "url": target_url, "runs_updated": added}

    # operation == "remove"
    removed = 0
    for para in target_shape.text_frame.paragraphs:
        for run in para.runs:
            if run.hyperlink and run.hyperlink.address:
                run.hyperlink.address = None
                removed += 1
    _audit_log("manage_hyperlinks_remove", resolved, slide_index, {"shape_id": shape_id})
    return {"ok": True, "slide_index": slide_index, "shape_id": shape_id,
            "operation": "remove", "runs_cleared": removed}


def remove_empty_placeholders(
    path: str,
    slide_index: int,
    confirm: bool = False,
) -> dict:
    """Remove placeholder shapes that are empty or contain only default prompt text.

    A placeholder is considered empty when its text frame text (stripped) is either
    empty or matches one of the well-known PowerPoint default prompt strings such as
    'Click to add title', 'Click to add text', 'Click to add subtitle', etc.

    Non-placeholder shapes (text boxes, auto-shapes) are never touched.

    Uses a snapshot of the shape list before the deletion loop to avoid index
    corruption during live-list mutation.

    Requires PPT_ENABLE_WRITE=true and confirm=True.

    Returns: {"removed_count": int, "remaining_shapes": int}
    """
    _check_write()
    _check_confirm(confirm)
    resolved = _check_path(path)
    prs = _load_prs(resolved)
    slide = _validate_slide_index(prs, slide_index)
    ph_snapshot = [s for s in slide.shapes if _is_empty_placeholder(s)]
    for ph in ph_snapshot:
        ph._element.getparent().remove(ph._element)
    _atomic_save(prs, resolved)
    _evict_prs(resolved)
    removed = len(ph_snapshot)
    remaining = len(list(slide.shapes))
    _audit_log("remove_empty_placeholders", resolved, slide_index, {"removed_count": removed})
    return {"removed_count": removed, "remaining_shapes": remaining}
