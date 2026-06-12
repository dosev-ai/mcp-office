"""review_pptx — re-export shim; implementation in _review_detectors + _review_export."""
from pptmcp._review_detectors import (  # noqa: F401
    DEFAULT_TEXT,
    _EMU_PER_INCH,
    _PT_HIGH,
    _PT_MEDIUM,
    _SPARSE_FILL_THRESHOLD,
    _SPARSE_COUNT_THRESHOLD,
    _TEXT_TRUNCATION_LIMIT,
    _detect_low_font_size,
    _detect_png_issues,
    _detect_shapes_outside_slide,
    _detect_sparse_slide,
    _detect_table_content_clipping,
    _detect_text_overflow_heuristic,
    _detect_unfinished_placeholder,
)
from pptmcp._review_export import review_slide_export  # noqa: F401

__all__ = [
    "review_slide_export",
    "DEFAULT_TEXT",
    "_detect_png_issues",
    "_detect_text_overflow_heuristic",
    "_detect_shapes_outside_slide",
    "_detect_table_content_clipping",
    "_detect_unfinished_placeholder",
    "_detect_low_font_size",
    "_detect_sparse_slide",
]
