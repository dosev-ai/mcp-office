"""Formatting MCP tools for excelmcp.

Extracted from server.py. Tools cover number formats, cell style, alignment,
borders, multi-sheet formatting, format-only copy, write+format combined,
and conditional formatting.

Import side-effect: importing this module registers all tools on the shared
mcp instance from excelmcp._server_instance.
"""
from __future__ import annotations

from fastmcp.exceptions import ToolError

from excelmcp import workbook_openpyxl as wb
from excelmcp._core import _check_confirm, _check_write
from excelmcp._server_instance import mcp
from excelmcp.workbook_openpyxl import ExcelMCPError

# ---------------------------------------------------------------------------
# Formatting tools  (Phase 2.0 — require EXCEL_ENABLE_WRITE=true + confirm=True)
# ---------------------------------------------------------------------------

@mcp.tool()
def apply_number_format(
    path: str,
    sheet: str,
    address: str,
    format_code: str,
    confirm: bool = False,
) -> dict:
    """Apply a number format code (e.g. '#,##0.00', 'YYYY-MM-DD') to a cell or range.

    Does NOT auto-save. Call save() to persist changes to disk.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        result = wb.apply_number_format(
            path, sheet, address, format_code, confirm=confirm
        )
        return result
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def apply_style(
    path: str,
    sheet: str,
    ranges: list[str] | None = None,
    address: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_color: str | None = None,
    bg_color: str | None = None,
    confirm: bool = False,
) -> dict:
    """Apply styling to one or more cell ranges.

    address: single range string (e.g. 'A1:B2') — preferred.
    ranges: list of range strings — still supported for backward compat.
    Provide address OR ranges, not both.
    font_color and bg_color must be 6 hex digits without '#', e.g. 'FF0000'.
    Requires EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if address is not None and ranges is not None:
        raise ToolError("Provide address or ranges, not both")
    if address is not None:
        resolved_ranges = [address]
    elif ranges is not None:
        resolved_ranges = ranges
    else:
        raise ToolError("address or ranges is required")
    try:
        return wb.apply_style_to_ranges(
            path, sheet, resolved_ranges,
            bold=bold, italic=italic,
            font_size=font_size, font_color=font_color, bg_color=bg_color,
            confirm=confirm,
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def apply_alignment(
    path: str,
    sheet: str,
    address: str,
    horizontal: str | None = None,
    vertical: str | None = None,
    wrap_text: bool | None = None,
    confirm: bool = False,
) -> dict:
    """Apply alignment to a cell or range.

    *horizontal*: left, center, right, fill, justify.
    *vertical*: top, center, bottom.
    Does NOT auto-save. Call save() to persist changes to disk.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        result = wb.apply_alignment(
            path, sheet, address,
            horizontal=horizontal, vertical=vertical,
            wrap_text=wrap_text, confirm=confirm,
        )
        return result
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def add_border(
    path: str,
    sheet: str,
    address: str,
    border_style: str = "thin",
    color: str = "000000",
    confirm: bool = False,
) -> dict:
    """Apply a border to all sides of a cell or range.

    *border_style*: thin, medium, thick, dashed, dotted, double.
    *color*: 6 hex digits without '#', e.g. '000000'.
    Does NOT auto-save. Call save() to persist changes to disk.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        result = wb.add_border(
            path, sheet, address,
            border_style=border_style, color=color, confirm=confirm,
        )
        return result
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def apply_format_to_sheet_list(
    path: str,
    sheets: list[str] | None = None,
    address: str | None = None,
    style: dict | None = None,
    number_format: str | None = None,
    alignment: dict | None = None,
    column_widths: dict | None = None,
    row_heights: dict | None = None,
    confirm: bool = False,
) -> dict:
    """Apply formatting to the same address across multiple sheets.

    sheets=None applies to all non-underscore-prefixed sheets (_Lists etc skipped).
    Explicit sheet list overrides the skip guard.
    style keys: bold, italic, font_size, font_color, bg_color.
    alignment keys: horizontal, vertical, wrap_text.
    column_widths: {"A": 15, "B:D": 12} — sets column widths on every target sheet.
    row_heights: {"1": 25, "2:10": 18} — sets row heights on every target sheet.
    EXCEL_MAX_RANGE_CELLS is enforced per sheet.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.apply_format_to_sheet_list(
            path, sheets=sheets, address=address, style=style,
            number_format=number_format, alignment=alignment,
            column_widths=column_widths, row_heights=row_heights,
            confirm=confirm,
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def copy_range_format(
    path: str,
    src_sheet: str,
    src_address: str,
    dst_sheet: str,
    dst_address: str,
    copy_border: bool = False,
    confirm: bool = False,
) -> dict:
    """Copy formatting only (no values) from src to dst range. Cross-sheet. dst can be single cell (auto-expands).

    Copies: font, fill, number_format, alignment. Border only if copy_border=True.
    Does NOT copy: values, formulas, comments, hyperlinks.
    MergedCell slave cells in src are skipped silently.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.copy_range_format(
            path, src_sheet=src_sheet, src_address=src_address,
            dst_sheet=dst_sheet, dst_address=dst_address,
            copy_border=copy_border, confirm=confirm,
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def write_range_with_format(
    path: str,
    sheet: str,
    address: str,
    data: list,
    style: dict | None = None,
    number_format: str | None = None,
    alignment: dict | None = None,
    confirm: bool = False,
) -> dict:
    """Write a 2-D array of values to a range and apply consistent formatting in one call.

    Args:
        path: Absolute path to the .xlsx workbook.
        sheet: Worksheet name.
        address: Top-left cell or full range address (e.g. "B2" or "B2:E10").
        data: 2-D list (rows x columns) of values to write.
        style: Optional style dict. Allowed keys: bold, italic, underline, font_color,
               bg_color, font_size, font_name.
        number_format: Optional Excel number format string (e.g. "#,##0.00").
        alignment: Optional alignment dict. Allowed keys: horizontal, vertical,
                   wrap_text, shrink_to_fit, text_rotation, indent.
        confirm: Must be True to execute the write.
    """
    try:
        from excelmcp.workbook_openpyxl import write_range_with_format as _fn
        return _fn(path=path, sheet=sheet, address=address, data=data, style=style,
                   number_format=number_format, alignment=alignment, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def apply_conditional_format(
    path: str,
    sheet: str,
    address: str,
    rule_type: str,
    condition: str,
    format_spec: dict,
    formula: str | None = None,
    confirm: bool = False,
) -> dict:
    """Apply a conditional-formatting rule to a range.

    Args:
        path: Absolute path to the .xlsx workbook.
        sheet: Worksheet name.
        address: Cell range to apply CF to (e.g. "A1:Z100").
        rule_type: "cell_is" (compare cell value) or "formula" (arbitrary formula).
        condition: For cell_is: "<operator> <value>" e.g. "greaterThan 100",
                   "between 10 20". Operator choices: greaterThan, greaterThanOrEqual,
                   lessThan, lessThanOrEqual, equal, notEqual, between, notBetween.
                   For formula: unused (pass formula arg instead).
        format_spec: Formatting to apply when rule fires. Allowed keys: bold, italic,
                     font_color, bg_color, number_format (number_format silently ignored).
        formula: Required when rule_type="formula". Must start with "=".
        confirm: Must be True to execute the write.
    """
    try:
        from excelmcp.workbook_openpyxl import apply_conditional_format as _fn
        return _fn(path=path, sheet=sheet, address=address, rule_type=rule_type,
                   condition=condition, format_spec=format_spec, formula=formula, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# format dispatcher (Phase 3E) — single entry-point for all format operations
# ---------------------------------------------------------------------------

_FORMAT_VALID_OPERATIONS = (
    "alignment",
    "border",
    "conditional",
    "copy_format",
    "number_format",
    "sheet_list",
    "style",
    "write_with_format",
)


@mcp.tool(name="format")
def format_dispatcher(
    operation: str,
    path: str,
    # common
    sheet: str | None = None,
    address: str | None = None,
    # apply_style extras
    ranges: list[str] | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_color: str | None = None,
    bg_color: str | None = None,
    # apply_alignment
    horizontal: str | None = None,
    vertical: str | None = None,
    wrap_text: bool | None = None,
    # add_border
    border_style: str = "thin",
    color: str = "000000",
    # apply_format_to_sheet_list
    sheets: list[str] | None = None,
    style: dict | None = None,
    number_format: str | None = None,
    alignment: dict | None = None,
    column_widths: dict | None = None,
    row_heights: dict | None = None,
    # copy_range_format
    src_sheet: str | None = None,
    src_address: str | None = None,
    dst_sheet: str | None = None,
    dst_address: str | None = None,
    copy_border: bool = False,
    # write_range_with_format
    data: list | None = None,
    # apply_number_format
    format_code: str | None = None,
    # apply_conditional_format
    rule_type: str | None = None,
    condition: str | None = None,
    format_spec: dict | None = None,
    formula: str | None = None,
    confirm: bool = False,
) -> dict:
    """Unified format dispatcher — routes 8 format operations through a single tool.

    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.

    operation vocabulary
    --------------------
    "number_format"   Apply a number-format code to a cell or range.
                      Params: path, sheet, address, format_code, confirm.

    "style"           Apply font/fill styling to one or more ranges.
                      Params: path, sheet, address (or ranges), bold, italic,
                      font_size, font_color, bg_color, confirm.

    "alignment"       Apply horizontal/vertical/wrap_text alignment.
                      Params: path, sheet, address, horizontal, vertical,
                      wrap_text, confirm.

    "border"          Apply border to all sides of a cell or range.
                      Params: path, sheet, address, border_style, color, confirm.

    "sheet_list"      Apply formatting to the same address across multiple sheets.
                      Params: path, sheets, address, style, number_format,
                      alignment, column_widths, row_heights, confirm.

    "copy_format"     Copy formatting only (no values) from src to dst range.
                      Params: path, src_sheet, src_address, dst_sheet,
                      dst_address, copy_border, confirm.

    "write_with_format"  Write a 2-D array and apply formatting in one call.
                      Params: path, sheet, address, data, style, number_format,
                      alignment, confirm.

    "conditional"     Apply a conditional-formatting rule to a range.
                      Params: path, sheet, address, rule_type, condition,
                      format_spec, formula, confirm.

    Deprecated aliases
    ------------------
    The 8 standalone tools (apply_number_format, apply_style, apply_alignment,
    add_border, apply_format_to_sheet_list, copy_range_format,
    write_range_with_format, apply_conditional_format) remain callable but are
    deprecated. Prefer this dispatcher for all new integrations.
    """
    _check_write()
    _check_confirm(confirm)

    if operation == "number_format":
        try:
            return wb.apply_number_format(
                path, sheet, address, format_code, confirm=confirm
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e

    elif operation == "style":
        if address is not None and ranges is not None:
            raise ToolError("Provide address or ranges, not both")
        if address is not None:
            resolved_ranges = [address]
        elif ranges is not None:
            resolved_ranges = ranges
        else:
            raise ToolError("address or ranges is required for operation='style'")
        try:
            return wb.apply_style_to_ranges(
                path, sheet, resolved_ranges,
                bold=bold, italic=italic,
                font_size=font_size, font_color=font_color, bg_color=bg_color,
                confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e

    elif operation == "alignment":
        try:
            return wb.apply_alignment(
                path, sheet, address,
                horizontal=horizontal, vertical=vertical,
                wrap_text=wrap_text, confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e

    elif operation == "border":
        try:
            return wb.add_border(
                path, sheet, address,
                border_style=border_style, color=color, confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e

    elif operation == "sheet_list":
        try:
            return wb.apply_format_to_sheet_list(
                path, sheets=sheets, address=address, style=style,
                number_format=number_format, alignment=alignment,
                column_widths=column_widths, row_heights=row_heights,
                confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e

    elif operation == "copy_format":
        try:
            return wb.copy_range_format(
                path, src_sheet=src_sheet, src_address=src_address,
                dst_sheet=dst_sheet, dst_address=dst_address,
                copy_border=copy_border, confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e

    elif operation == "write_with_format":
        try:
            from excelmcp.workbook_openpyxl import write_range_with_format as _fn
            return _fn(
                path=path, sheet=sheet, address=address, data=data,
                style=style, number_format=number_format,
                alignment=alignment, confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e

    elif operation == "conditional":
        try:
            from excelmcp.workbook_openpyxl import apply_conditional_format as _fn
            return _fn(
                path=path, sheet=sheet, address=address,
                rule_type=rule_type, condition=condition,
                format_spec=format_spec, formula=formula, confirm=confirm,
            )
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e

    else:
        raise ToolError(
            f"Unknown format operation: {operation!r}. "
            f"Valid: {', '.join(sorted(_FORMAT_VALID_OPERATIONS))}"
        )
