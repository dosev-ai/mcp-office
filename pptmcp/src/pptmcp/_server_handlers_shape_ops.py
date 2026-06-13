"""
Shape operation helpers and dispatchers for pptmcp write tools.
Extracted from _server_handlers_write.py to reduce that file from RED (614L) to GREEN.
Backwards-compatible re-exports remain in _server_handlers_write.py.
"""
from __future__ import annotations

__all__: list[str] = []

from fastmcp.exceptions import ToolError

from pptmcp import shapes_pptx as shapes
from pptmcp.presentation_pptx import PPTMCPError

# ---------------------------------------------------------------------------
# Private helpers (module-level so add_content/set_format can call them)
# ---------------------------------------------------------------------------

def _add_textbox(
    path: str,
    slide_index: int,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str = "",
    confirm: bool = False,
) -> dict:
    try:
        return shapes.add_textbox(path, slide_index, left, top, width, height, text, confirm)
    except PPTMCPError as e:
        raise ToolError(str(e)) from e


def _add_shape(
    path: str,
    slide_index: int,
    shape_type: str,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str = "",
    confirm: bool = False,
) -> dict:
    try:
        return shapes.add_shape(path, slide_index, shape_type, left, top, width, height, text, confirm)
    except PPTMCPError as e:
        raise ToolError(str(e)) from e


def _add_table_to_slide(
    path: str,
    slide_index: int,
    rows: int,
    cols: int,
    left: float,
    top: float,
    width: float,
    height: float,
    data: list[list[str]] | None = None,
    confirm: bool = False,
    font_size_pt: float | None = None,
    header_font_size_pt: float | None = None,
    header_bold: bool | None = None,
    row_height_pt: float | None = None,
    suppress_content_placeholder: bool = True,
) -> dict:
    try:
        return shapes.add_table_to_slide(
            path,
            slide_index,
            rows,
            cols,
            left,
            top,
            width,
            height,
            data,
            confirm,
            font_size_pt=font_size_pt,
            header_font_size_pt=header_font_size_pt,
            header_bold=header_bold,
            row_height_pt=row_height_pt,
            suppress_content_placeholder=suppress_content_placeholder,
        )
    except PPTMCPError as e:
        raise ToolError(str(e)) from e


def _set_text_format(
    path: str,
    slide_index: int,
    shape_id: int,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size_pt: float | None = None,
    font_name: str | None = None,
    color_hex: str | None = None,
    confirm: bool = False,
    paragraph_range: list[int] | None = None,
    run_index: int | None = None,
) -> dict:
    try:
        return shapes.set_text_format(
            path, slide_index, shape_id, bold, italic, font_size_pt, font_name,
            color_hex, confirm, paragraph_range=paragraph_range, run_index=run_index,
        )
    except PPTMCPError as e:
        raise ToolError(str(e)) from e


def _set_paragraph_format(
    path: str,
    slide_index: int,
    shape_id: int,
    paragraph_index: int = 0,
    alignment: str | None = None,
    line_spacing: float | None = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    confirm: bool = False,
    paragraph_range: list[int] | None = None,
) -> dict:
    try:
        return shapes.set_paragraph_format(
            path, slide_index, shape_id, paragraph_index, alignment, line_spacing,
            space_before_pt, space_after_pt, confirm, paragraph_range=paragraph_range,
        )
    except PPTMCPError as e:
        raise ToolError(str(e)) from e


def set_table_style(
    path: str,
    slide_index: int,
    shape_id: int,
    font_size_pt: float | None = None,
    bold: bool | None = None,
    color_hex: str | None = None,
    header_font_size_pt: float | None = None,
    header_bold: bool | None = None,
    header_color_hex: str | None = None,
    fill_color_hex: str | None = None,
    header_fill_color_hex: str | None = None,
    row_height_pt: float | None = None,
    header_row_height_pt: float | None = None,
    row_heights_pt: list[float] | None = None,
    confirm: bool = False,
) -> dict:
    """Style a table shape. Requires PPT_ENABLE_WRITE=true + confirm=True."""
    try:
        return shapes.set_table_style(
            path,
            slide_index,
            shape_id,
            font_size_pt=font_size_pt,
            bold=bold,
            color_hex=color_hex,
            header_font_size_pt=header_font_size_pt,
            header_bold=header_bold,
            header_color_hex=header_color_hex,
            fill_color_hex=fill_color_hex,
            header_fill_color_hex=header_fill_color_hex,
            row_height_pt=row_height_pt,
            header_row_height_pt=header_row_height_pt,
            row_heights_pt=row_heights_pt,
            confirm=confirm,
        )
    except PPTMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# Dispatcher tools — module-level so they are importable and testable directly
# ---------------------------------------------------------------------------

def add_content(
    content_type: str,
    path: str,
    slide_index: int,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str = "",
    confirm: bool = False,
    shape_type: str = "",
    rows: int = 0,
    cols: int = 0,
    data: list[list[str]] | None = None,
    font_size_pt: float | None = None,
    header_font_size_pt: float | None = None,
    header_bold: bool | None = None,
    row_height_pt: float | None = None,
    suppress_content_placeholder: bool = True,
) -> dict:
    """DEPRECATED: Use shape(operation="add_text_box"/"add_autoshape"/"add_table") instead.
    Replacement mapping:
      add_content(content_type="textbox") → shape(operation="add_text_box")
      add_content(content_type="shape")   → shape(operation="add_autoshape")
      add_content(content_type="table")   → shape(operation="add_table")
    This alias will be removed after 2 minor releases.

    Add content to a slide. Requires PPT_ENABLE_WRITE=true + confirm=True.

    content_type: 'textbox' | 'shape' | 'table'
      'textbox'  - requires: left, top, width, height; optional: text
      'shape'    - requires: shape_type (RECTANGLE, OVAL, ROUNDED_RECTANGLE, etc.),
                   left, top, width, height; optional: text
      'table'    - requires: rows (>0), cols (>0), left, top, width, height;
                   optional: data (2-D list of strings)
    """
    try:
        if content_type == "textbox":
            return _add_textbox(path, slide_index, left, top, width, height, text, confirm)
        elif content_type == "shape":
            if not shape_type:
                raise PPTMCPError("shape_type is required for content_type='shape'")
            return _add_shape(path, slide_index, shape_type, left, top, width, height, text, confirm)
        elif content_type == "table":
            if rows <= 0 or cols <= 0:
                raise PPTMCPError("rows and cols must be > 0 for content_type='table'")
            return _add_table_to_slide(
                path,
                slide_index,
                rows,
                cols,
                left,
                top,
                width,
                height,
                data,
                confirm,
                font_size_pt=font_size_pt,
                header_font_size_pt=header_font_size_pt,
                header_bold=header_bold,
                row_height_pt=row_height_pt,
                suppress_content_placeholder=suppress_content_placeholder,
            )
        else:
            raise PPTMCPError(f"Unknown content_type: {content_type!r}. Must be one of: textbox, shape, table")
    except ToolError:
        raise
    except PPTMCPError as e:
        raise ToolError(str(e)) from e
    except Exception as e:
        raise ToolError(str(e)) from e


def set_format(
    target: str,
    path: str,
    slide_index: int,
    shape_id: int,
    confirm: bool = False,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size_pt: float | None = None,
    font_name: str | None = None,
    color_hex: str | None = None,
    paragraph_index: int = 0,
    alignment: str | None = None,
    line_spacing: float | None = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    paragraph_range: list[int] | None = None,
    run_index: int | None = None,
) -> dict:
    """Set formatting in a shape. Requires PPT_ENABLE_WRITE=true + confirm=True.

    target: 'text' | 'paragraph'
      'text'      - bold, italic, font_size_pt, font_name, color_hex (6-char hex, e.g. 'FF0000')
                    paragraph_range: optional [start, end] (inclusive) to limit paragraphs
                    run_index: optional 0-based run index within each matched paragraph
      'paragraph' - paragraph_index (default 0), alignment (CENTER/LEFT/RIGHT/JUSTIFY/DISTRIBUTE),
                    line_spacing, space_before_pt, space_after_pt
                    paragraph_range: optional [start, end] (inclusive) overrides paragraph_index
    """
    try:
        if target == "text":
            return _set_text_format(
                path, slide_index, shape_id, bold, italic, font_size_pt, font_name,
                color_hex, confirm, paragraph_range=paragraph_range, run_index=run_index,
            )
        elif target == "paragraph":
            return _set_paragraph_format(
                path, slide_index, shape_id, paragraph_index, alignment, line_spacing,
                space_before_pt, space_after_pt, confirm, paragraph_range=paragraph_range,
            )
        else:
            raise PPTMCPError(f"Unknown target: {target!r}. Must be one of: text, paragraph")
    except ToolError:
        raise
    except PPTMCPError as e:
        raise ToolError(str(e)) from e
    except Exception as e:
        raise ToolError(str(e)) from e


def shape(
    operation: str,
    path: str,
    slide_index: int,
    confirm: bool = False,
    left: float = 0.0,
    top: float = 0.0,
    width: float = 0.0,
    height: float = 0.0,
    text: str = "",
    shape_type: str = "",
    rows: int = 0,
    cols: int = 0,
    data: list[list[str]] | None = None,
    shape_id: int | None = None,
    target: str = "text",
    bold: bool | None = None,
    italic: bool | None = None,
    font_size_pt: float | None = None,
    font_name: str | None = None,
    color_hex: str | None = None,
    paragraph_index: int = 0,
    alignment: str | None = None,
    line_spacing: float | None = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    header_font_size_pt: float | None = None,
    header_bold: bool | None = None,
    row_height_pt: float | None = None,
    suppress_content_placeholder: bool = True,
    header_color_hex: str | None = None,
    fill_color_hex: str | None = None,
    header_fill_color_hex: str | None = None,
    header_row_height_pt: float | None = None,
    row_heights_pt: list[float] | None = None,
    new_left_in: float | None = None,
    new_top_in: float | None = None,
    new_width_in: float | None = None,
    new_height_in: float | None = None,
    paragraph_range: list[int] | None = None,
    run_index: int | None = None,
) -> dict:
    """Unified shape mutation dispatcher. Requires PPT_ENABLE_WRITE=true + confirm=True.

    operation: 'add_text_box' | 'add_autoshape' | 'add_table' | 'delete'
               | 'set_properties' | 'set_table_style' | 'set_geometry'
      'add_text_box'   - params: left, top, width, height; optional: text
      'add_autoshape'  - params: shape_type (RECTANGLE/OVAL/ROUNDED_RECTANGLE/etc.),
                         left, top, width, height; optional: text
      'add_table'      - params: rows (>0), cols (>0), left, top, width, height;
                         optional: data (2-D list of strings)
      'delete'         - params: shape_id (required)
      'set_properties' - params: shape_id (required), target ('text', 'paragraph', or 'fill')
                         'text' sub-target: bold, italic, font_size_pt, font_name,
                                            color_hex (6-char hex, e.g. 'FF0000')
                         'paragraph' sub-target: paragraph_index, alignment
                                     (CENTER/LEFT/RIGHT/JUSTIFY/DISTRIBUTE),
                                     line_spacing, space_before_pt, space_after_pt
                         'fill' sub-target: fill_color_hex (6-char hex, e.g. 'FF0000')
      'set_geometry'   - params: shape_id (required); any subset of
                         new_left_in/new_top_in/new_width_in/new_height_in
                         (inches; omitted axes unchanged). Off-slide results
                         surface as advisory warnings, never block.
    """
    try:
        if operation == "add_text_box":
            return add_content("textbox", path, slide_index, left, top, width, height, text, confirm)
        elif operation == "add_autoshape":
            return add_content(
                "shape", path, slide_index, left, top, width, height, text, confirm,
                shape_type=shape_type,
            )
        elif operation == "add_table":
            return add_content(
                "table", path, slide_index, left, top, width, height, "", confirm,
                rows=rows, cols=cols, data=data,
                font_size_pt=font_size_pt,
                header_font_size_pt=header_font_size_pt,
                header_bold=header_bold,
                row_height_pt=row_height_pt,
                suppress_content_placeholder=suppress_content_placeholder,
            )
        elif operation == "delete":
            if shape_id is None:
                raise PPTMCPError("shape_id is required for operation='delete'")
            return shapes.delete_shape(path, slide_index, shape_id, confirm)
        elif operation == "set_properties":
            if shape_id is None:
                raise PPTMCPError("shape_id is required for operation='set_properties'")
            if target == "fill":
                if fill_color_hex is None:
                    raise PPTMCPError("fill_color_hex is required for target='fill'")
                return shapes.set_shape_fill(
                    path, slide_index, shape_id, fill_color_hex, confirm
                )
            result = set_format(
                target, path, slide_index, shape_id, confirm,
                bold, italic, font_size_pt, font_name, color_hex,
                paragraph_index, alignment, line_spacing, space_before_pt, space_after_pt,
                paragraph_range=paragraph_range, run_index=run_index,
            )
            if fill_color_hex is not None:
                fill_result = shapes.set_shape_fill(
                    path, slide_index, shape_id, fill_color_hex, confirm
                )
                result["fill_color_hex"] = fill_result["fill_color_hex"]
                result["fill_updated"] = fill_result["fill_updated"]
            return result
        elif operation == "set_table_style":
            if shape_id is None:
                raise PPTMCPError("shape_id is required for operation='set_table_style'")
            return set_table_style(
                path,
                slide_index,
                shape_id,
                font_size_pt=font_size_pt,
                bold=bold,
                color_hex=color_hex,
                header_font_size_pt=header_font_size_pt,
                header_bold=header_bold,
                header_color_hex=header_color_hex,
                fill_color_hex=fill_color_hex,
                header_fill_color_hex=header_fill_color_hex,
                row_height_pt=row_height_pt,
                header_row_height_pt=header_row_height_pt,
                row_heights_pt=row_heights_pt,
                confirm=confirm,
            )
        elif operation == "set_geometry":
            if shape_id is None:
                raise PPTMCPError("shape_id is required for operation='set_geometry'")
            return shapes.set_shape_geometry(
                path,
                slide_index,
                shape_id,
                left_in=new_left_in,
                top_in=new_top_in,
                width_in=new_width_in,
                height_in=new_height_in,
                confirm=confirm,
            )
        else:
            raise PPTMCPError(
                f"Unknown operation: {operation!r}. "
                "Must be one of: add_text_box, add_autoshape, add_table, delete, "
                "set_properties, set_table_style, set_geometry"
            )
    except ToolError:
        raise
    except PPTMCPError as e:
        raise ToolError(str(e)) from e
    except Exception as e:
        raise ToolError(str(e)) from e
