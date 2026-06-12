"""
Backend for batch_set_text — update text on multiple shapes on a single slide
in one open/save cycle. No MCP/FastMCP imports.
"""
from __future__ import annotations

import logging

from pptmcp.presentation_pptx import (
    PPTMCPError,  # noqa: F401
    NotAllowedError,  # noqa: F401  (raised by _check_write/_check_confirm)
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

logger = logging.getLogger(__name__)

_MAX_UPDATES = 200
_TEXT_MAX = 50_000


def batch_set_text(
    path: str,
    slide_index: int,
    updates: list[dict],
    confirm: bool = False,
) -> dict:
    """Update text on multiple shapes on a single slide in one open/save cycle.

    Gate order: _check_write -> _check_confirm -> validate inputs -> _check_path -> file I/O.

    Args:
        path: Path to the .pptx file (must be in PPT_ALLOWLIST_ROOTS).
        slide_index: Zero-based slide index.
        updates: List of dicts, each with "shape_id" (int) and "text" (str).
        confirm: Must be True to proceed (mutation gate).

    Returns:
        {updated, path, slide_index, results[{shape_id, status, previous_text, new_text}],
         skipped[{shape_id, reason}], error_count}

    Raises:
        NotAllowedError: PPT_ENABLE_WRITE not set or confirm=False.
        ValidationError: updates list is invalid or slide_index out of range.
    """
    # Gate 1: env var write check
    _check_write()
    # Gate 2: confirm parameter check
    _check_confirm(confirm)

    # Gate 3: validate updates list
    if not updates:
        raise ValidationError("updates must be a non-empty list")
    if len(updates) > _MAX_UPDATES:
        raise ValidationError(
            f"updates length {len(updates)} exceeds maximum {_MAX_UPDATES}"
        )
    for i, entry in enumerate(updates):
        if not isinstance(entry, dict):
            raise ValidationError(f"updates[{i}] must be a dict")
        if "shape_id" not in entry:
            raise ValidationError(f"updates[{i}] missing required key 'shape_id'")
        if not isinstance(entry["shape_id"], int):
            raise ValidationError(
                f"updates[{i}]['shape_id'] must be an int, got {type(entry['shape_id']).__name__}"
            )
        if "text" not in entry:
            raise ValidationError(f"updates[{i}] missing required key 'text'")
        if not isinstance(entry["text"], str):
            raise ValidationError(
                f"updates[{i}]['text'] must be a str, got {type(entry['text']).__name__}"
            )
        if len(entry["text"]) > _TEXT_MAX:
            raise ValidationError(
                f"updates[{i}]['text'] length {len(entry['text'])} exceeds maximum {_TEXT_MAX}"
            )

    # Gate 4: path validation (allowlist + extension)
    resolved = _check_path(path)

    # File I/O
    prs = _load_prs(resolved)
    slide = _validate_slide_index(prs, slide_index)

    results: list[dict] = []
    skipped: list[dict] = []

    for entry in updates:
        sid = entry["shape_id"]
        text = entry["text"]

        try:
            shape = _find_shape_by_id(slide, sid)
        except ValidationError as exc:
            skipped.append({"shape_id": sid, "reason": str(exc)})
            continue

        try:
            tf = shape.text_frame
        except AttributeError:
            skipped.append({"shape_id": sid, "reason": f"shape {sid} has no text frame"})
            continue

        previous_text = tf.text
        tf.text = text
        results.append({
            "shape_id": sid,
            "status": "ok",
            "previous_text": previous_text,
            "new_text": text,
        })

    _atomic_save(prs, resolved)
    _evict_prs(resolved)
    _audit_log(
        "batch_set_text",
        resolved,
        slide_idx=slide_index,
        extra={"updated_count": len(results), "skipped_count": len(skipped)},
    )

    return {
        "updated": True,
        "path": str(resolved),
        "slide_index": slide_index,
        "results": results,
        "skipped": skipped,
        "error_count": len(skipped),
    }
