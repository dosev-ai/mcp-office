"""_shapes_overlap — detect_overlapping_shapes function for pptmcp.

Extracted from _shapes_geometry.py to keep that file within AMBER bounds.
Re-exported from _shapes_geometry for backwards compatibility.
"""
from __future__ import annotations

from pptmcp.presentation_pptx import _check_path, _load_prs, _validate_slide_index


def detect_overlapping_shapes(path: str, slide_index: int) -> dict:
    """Return all overlapping shape pairs on a slide with overlap area in sq.in (read-only)."""
    # BLK-01: _check_path MUST be first statement
    resolved = _check_path(path)
    prs = _load_prs(resolved)
    # BLK-04: use _validate_slide_index helper (not inline bounds check)
    slide = _validate_slide_index(prs, slide_index)
    shapes_list = list(slide.shapes)
    pairs = []
    for i in range(len(shapes_list)):
        for j in range(i + 1, len(shapes_list)):
            sa, sb = shapes_list[i], shapes_list[j]
            if any(v is None for v in [sa.left, sa.top, sa.width, sa.height,
                                        sb.left, sb.top, sb.width, sb.height]):
                continue
            ox = max(0, min(sa.left + sa.width, sb.left + sb.width) - max(sa.left, sb.left))
            oy = max(0, min(sa.top + sa.height, sb.top + sb.height) - max(sa.top, sb.top))
            area_emu2 = ox * oy
            if area_emu2 > 0:
                pairs.append({
                    "shape_a": {"shape_id": sa.shape_id, "name": sa.name},
                    "shape_b": {"shape_id": sb.shape_id, "name": sb.name},
                    "overlap_area_in2": round(area_emu2 / (914400 ** 2), 4),
                })
    return {"slide_index": slide_index, "overlapping_pairs": pairs, "pair_count": len(pairs)}


__all__ = ["detect_overlapping_shapes"]
