"""shapes_pptx — re-export shim; implementation in _shapes_helpers, _shapes_text, _shapes_geometry, _shapes_overlap, _shapes_links."""
from pptmcp._shapes_helpers import (  # noqa: F401
    _EMPTY_PLACEHOLDER_PATTERNS, _apply_cell_fill, _apply_cell_text_style,
    _calculate_overflow_risk, _first_overlap_for_shape, _is_empty_placeholder,
    _is_title_placeholder, _normalize_color_hex,
    _remove_empty_overlapping_content_placeholders, _shapes_overlap,
)
from pptmcp._shapes_text import (  # noqa: F401
    _A_BR, _A_PPR, _A_R, _A_RPR, _A_T, _XML_SPACE,
    _collect_para_text, _rebuild_para_text,
    edit_text_placeholder, replace_slide_text, set_paragraph_format,
    set_speaker_notes, set_text_format,
)
from pptmcp._shapes_geometry import (  # noqa: F401
    _style_table_shape, add_shape, add_table_to_slide, add_textbox,
    delete_shape, detect_overlapping_shapes, insert_image,
    set_shape_fill, set_shape_geometry, set_table_style,
)
from pptmcp._shapes_links import add_hyperlink, manage_hyperlinks, remove_empty_placeholders  # noqa: F401
