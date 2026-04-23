"""Formatting prompts: suggest, prepare-for-print, bulk format, style, header, import-CSV,
highlight exceptions, annotate, style-and-write, annotate workbook.

Registers 10 @mcp.prompt() on the shared mcp instance at import time.
"""
from __future__ import annotations

import logging

from excelmcp._server_instance import mcp

logger = logging.getLogger(__name__)


@mcp.prompt()
def suggest_formatting(sheet_json: str, report_purpose: str) -> str:
    """Given sheet data and purpose, suggest a formatting tool call sequence."""
    return (
        f"Sheet data:\n{sheet_json}\n\n"
        f"Report purpose: {report_purpose}\n\n"
        "Suggest a minimal, professional formatting sequence using these available tools: "
        "apply_style, apply_alignment, apply_number_format, set_column_widths, set_row_heights, add_border. "
        "For each suggestion output: tool name, parameters (as JSON), and a one-line reason. "
        "Focus on making the report immediately usable -- avoid unnecessary decorative formatting."
    )


@mcp.prompt()
def prepare_for_print(sheet_description: str, page_size: str = "A4") -> str:
    """Given a sheet description and page size, suggest print area and page setup for export."""
    return (
        f"Sheet description:\n{sheet_description}\n\nTarget page size: {page_size}\n\n"
        "Suggest a professional print configuration for this sheet using:\n"
        "- set_print_area: define the exact data range to print (e.g. 'A1:H40')\n"
        "- freeze_panes: freeze header rows so they print on each page if the sheet is tall\n"
        "- set_column_widths: adjust widths so data fits within page margins\n\n"
        "Output a JSON object with:\n"
        '{"print_area": "A1:H40", "address": "A2", "column_widths": {"A": 15, "B": 20}}\n\n'
        "Keep the print area tight to the data -- do not include blank rows/columns. "
        f"For {page_size}, typical column width budget is ~100 character units total."
    )


@mcp.prompt()
def bulk_format_report(sheet_json: str, brand_colors: str = "") -> str:
    """Orchestrate bulk styling for a report sheet.

    sheet_json: output from read_sheet_all.
    brand_colors: optional JSON with primary/secondary hex colors.
    """
    palette = (
        "Brand colors: " + brand_colors
        if brand_colors
        else "Use default navy/grey consulting palette: primary=#1F3864, accent=#2E74B5, "
             "header_bg=#1F3864, header_fg=#FFFFFF, alt_row=#F2F2F2"
    )
    return f"""You must apply professional formatting to a report sheet.

Sheet data summary:
{sheet_json}

{palette}

Apply formatting in this order using excelmcp tools:
1. apply_format_to_sheet_list -- apply table formatting to all content rows at once (most efficient)
2. apply_style -- style the header row: bold=true, font_color=#FFFFFF, bg_color=<primary>
3. set_column_widths -- auto-size: set data columns to 14-18, label columns to 20-25
4. apply_number_format -- numbers with commas, dates as "DD-MMM-YYYY", currencies as "#,##0.00"
5. freeze_panes -- freeze at address="A2" to keep headers visible
6. set_print_area -- set to match the data extent
7. apply_conditional_format (optional) -- highlight top/bottom 10% if a KPI column exists

Output the exact tool call sequence with populated arguments before executing.
Keep formatting declarations minimal -- no decorative borders or excessive colors."""


@mcp.prompt()
def style_dashboard_range(range_description: str, purpose: str) -> str:
    """Apply polished cell formatting to a dashboard range.

    Orchestrates: apply_style, apply_alignment, add_border,
    apply_number_format, apply_style_to_ranges.
    """
    return f"""You must style this dashboard range for professional presentation:

Range description: {range_description}
Purpose: {purpose}

Apply formatting in this exact sequence (ordering matters -- styles applied last win):
1. `apply_style` -- set base fill and font for the range:
   - Header cells: bold=true, bg_color="1F3864", font_color="FFFFFF", font_size=11
   - Data cells: bg_color=null (transparent), font_size=10
   - KPI metric cells: bold=true, font_size=14, font_color="1F3864"
2. `apply_alignment` -- set text alignment:
   - Label columns: horizontal="left", vertical="center"
   - Numeric columns: horizontal="right", vertical="center"
   - Header row: horizontal="center", vertical="center", wrap_text=true
3. `apply_number_format` -- apply Excel format codes per column type:
   - Integer counts: "#,##0"
   - Currency: "#,##0.00"
   - Percentage: "0.0%"
   - Date: "DD-MMM-YYYY"
   - Text: "@" (no format code needed -- leave as default)
4. `add_border` -- apply borders after fills (fills do not erase borders):
   - Outer border of the whole range: border_style="medium", color="1F3864"
   - Inner grid lines: border_style="thin", color="D9D9D9"
   Note: apply outer border first, then inner -- call `add_border` once per border style.
5. `apply_style_to_ranges` -- for non-contiguous ranges (e.g. alternating rows), pass a
   ranges list: ["A3:H3", "A5:H5", ...] with bg_color="F2F2F2" for alt-row banding.
6. Call `save` to persist.

Purpose-specific adjustments for "{purpose}":
- Executive dashboard: larger fonts (12-14pt KPIs), generous row heights (24+), minimal borders
- Operational report: standard 10pt, dense layout, thin grid borders throughout
- Data entry form: white bg for editable cells, colour-coded mandatory fields (light yellow bg)"""


@mcp.prompt()
def format_header_section(header_description: str) -> str:
    """Merge header cells, apply bold/background, and add auto-filter below headers.

    Orchestrates: merge_cells, unmerge_cells, list_merged_cells, auto_filter,
    apply_style, apply_alignment.
    """
    return f"""You must format the header section of this table:

{header_description}

Follow this sequence:
1. Identify the header row range (e.g. A1:H1 for a title, A2:H2 for column headers).
2. Check existing merges: call `list_merged_cells` to see any current merged ranges.
   If stale merges exist, call `unmerge_cells` on each before reapplying.
3. Merge title cells (row 1 only, if a title row is needed):
   Call `merge_cells` with address="A1:H1" (adjust columns to match data width).
   Only merge the TITLE row -- do NOT merge the column header row (auto-filter breaks on merged headers).
4. Style the merged title cell:
   Call `apply_style` with bold=true, bg_color="1F3864", font_color="FFFFFF", font_size=13.
5. Style the column header row (one row below the title, e.g. row 2):
   Call `apply_style` with bold=true, bg_color="2E74B5", font_color="FFFFFF", font_size=10.
6. Apply centre alignment to the title cell:
   Call `apply_alignment` with horizontal="center", vertical="center", wrap_text=false.
7. Apply auto-filter to the column header row:
   Call `auto_filter` with address="A2:H2" (use the column header range, NOT the title row).
   This enables dropdown filter arrows on each column header.
8. Call `save` to persist.

Common mistakes to avoid:
- Do NOT merge the auto-filter row -- Excel will disable filters on merged headers.
- The auto_filter address must cover all data columns including the last data column.
- To remove an auto-filter, call `auto_filter` with address=null."""


@mcp.prompt()
def import_and_style_csv(csv_path: str, target_sheet: str) -> str:
    """Import a CSV file, auto-fit columns, apply a table style, and save.

    Orchestrates: import_csv_to_sheet, set_column_widths, create_table, save.
    """
    return f"""You must import and style a CSV file into Excel.

CSV path: {csv_path}
Target sheet: {target_sheet}

Follow this sequence:
1. Import the CSV:
   Call `import_csv_to_sheet` with:
   - csv_path: "{csv_path}"
   - sheet_name: "{target_sheet}"
   - overwrite: false (set to true only if you intentionally want to replace existing data)
   - skip_header: false (keep the CSV header row as Excel row 1)
   All rows including the header are imported; dates and numbers remain as strings initially.
2. Get the used range to know the data extent:
   Call `get_used_range` on sheet "{target_sheet}" -- note the last row and column.
3. Auto-fit column widths:
   Call `set_column_widths` with col_range="A:<last_col>" and autofit=true, max_width=40.
   This approximates width from cell content with 2-char padding.
4. Convert the range to an Excel Table for structured filtering and referencing:
   Call `create_table` with:
   - table_range: "<A1:LastCol+LastRow>" (e.g. "A1:F500")
   - table_name: "<short_name>_Table" (no spaces, no special chars)
   - style: "TableStyleMedium9" (medium blue banded rows -- professional default)
5. Freeze the header row:
   Call `freeze_panes` with freeze_at="A2".
6. Call `save` to persist the workbook to disk.

Notes:
- The import does not auto-convert data types. If you need numeric columns, use
  `apply_number_format` after import (e.g. "#,##0" for integers, "0.0%" for percentages).
- If the CSV has a BOM or encoding issues, the first column header may have a garbage prefix --
  verify with `range_io` on "A1:A1" and use `cell` to fix if needed."""


@mcp.prompt()
def highlight_exceptions(range_address: str, rule_description: str) -> str:
    """Apply conditional formatting rules for exception highlighting.

    Orchestrates: apply_conditional_format.
    """
    return f"""You must apply conditional formatting to highlight exceptions in a range.

Target range: {range_address}
Rule intent: {rule_description}

Follow this sequence -- choose the rule type that best matches the intent:

## Option A -- Cell value threshold (most common)
Call `apply_conditional_format` with rule_type="cell_is":
- address: "{range_address}"
- rule_type: "cell_is"
- condition: "<operator> <value>" -- examples:
    "greaterThan 100"          -- highlight cells > 100
    "lessThan 0"               -- highlight negative values
    "between 80 120"           -- highlight values in range
    "equal Overdue"            -- highlight exact text match
- format_spec: {{"bold": true, "bg_color": "FF0000", "font_color": "FFFFFF"}}
  Standard exception colours: Red=FF0000, Amber=FFC000, Green=00B050

## Option B -- Formula-based rule (for cross-column logic)
Call `apply_conditional_format` with rule_type="formula":
- address: "{range_address}"
- rule_type: "formula"
- condition: "" (unused -- pass empty string)
- formula: "=<Excel formula starting with = -- use relative ref for first cell of range>"
  Examples:
    "=$D2>$E2"                 -- highlight rows where actual > budget
    "=ISBLANK(B2)"             -- highlight empty mandatory fields
    "=TODAY()-A2>30"           -- highlight dates older than 30 days
- format_spec: {{"bg_color": "FFC000", "font_color": "000000"}}

## Colour convention (traffic light)
- Critical / overdue / error: bg_color="FF0000", font_color="FFFFFF"
- Warning / at-risk / amber: bg_color="FFC000", font_color="000000"
- Good / on-track / pass:    bg_color="00B050", font_color="FFFFFF"

## Multiple rules on the same range
Apply rules in PRIORITY ORDER -- Excel evaluates rules top-to-bottom and stops at first match.
Call `apply_conditional_format` once per rule; first call = highest priority.

Confirm the rule applied by calling `get_used_range` to verify no errors, then call `save`."""


@mcp.prompt()
def annotate_and_configure_sheet(path: str, sheet_name: str) -> str:
    """Get/delete cell comments, manage hyperlinks, and configure sheet display properties.

    Orchestrates: cell_comment, hyperlink, sheet_properties.
    """
    return f"""You must annotate cells and configure the display properties of a sheet.

Workbook: {path}
Sheet: {sheet_name}

## Cell comments -- read or delete
Call `cell_comment`:
- operation="get": inspect the comment on a cell.
  Returns: {{has_comment, text, author, address}}. No confirm needed.
  Example: cell_comment(operation="get", path="{path}", sheet="{sheet_name}", address="B5")
- operation="delete": remove a comment from a cell.
  Requires EXCEL_ENABLE_WRITE=true and confirm=True.
NOTE: To ADD comments, use `bulk_add_comments` (multiple cells, preferred) or
`cell_comment` with `operation='add'` for a single cell.

## Hyperlinks -- get, add, or remove
Call `hyperlink`:
- operation="get": list all hyperlinks on the sheet.
  Returns: {{sheet, count, hyperlinks: [{{address, url, display_text, tooltip}}, ...]}}
  No confirm needed.
- operation="add": add one or more hyperlinks.
  links: [{{"address": "A1", "url": "https://...", "display_text": "...", "tooltip": "..."}}]
  Only https, http, and mailto schemes are permitted -- javascript:, data:, and vbscript:
  are always blocked (OWASP A03). Requires EXCEL_ENABLE_WRITE=true and confirm=True.
- operation="remove": remove a single hyperlink.
  address: "<cell address, e.g. A1>"
  Requires EXCEL_ENABLE_WRITE=true and confirm=True.

## Sheet display properties -- get or set
Call `sheet_properties`:
- operation="get": read the current display settings.
  Returns: {{tab_color, state, show_gridlines, show_row_col_headers, zoom_scale}}
  No confirm needed.
- operation="set": update one or more display settings.
  tab_color: 6-hex-digit colour (no #), e.g. "1F3864" for dark blue.
  state: "visible", "hidden", or "veryHidden" (hidden from sheet-tab right-click too).
  show_gridlines: true/false -- hide gridlines for dashboard sheets.
  show_row_col_headers: true/false -- hide row/column headers for clean presentations.
  zoom_scale: integer 10-400 -- set the zoom level.
  Requires EXCEL_ENABLE_WRITE=true and confirm=True.

Typical workflow for preparing a presentation sheet:
1. sheet_properties(operation="get") -- check current state
2. sheet_properties(operation="set", show_gridlines=false, show_row_col_headers=false, zoom_scale=100)
3. hyperlink(operation="get") -- audit existing links
4. hyperlink(operation="add", links=[...]) -- add navigation links
5. save -- persist all changes"""


@mcp.prompt()
def style_and_write_range(range_description: str, style_intent: str) -> str:
    """Write data with inline formatting, apply style to multiple ranges, or copy format only.

    Orchestrates: apply_style, write_range_with_format, copy_range_format.
    """
    return f"""You must apply styling or write formatted data to an Excel range.

Range description: {range_description}
Style intent: {style_intent}

Choose the right tool for the task:

## Write data AND format in one atomic call (preferred for new data)
Call `write_range_with_format`:
- path, sheet, address: target workbook/sheet/range (e.g. "B2" or "B2:E10")
- data: 2-D list of values [[row1], [row2], ...]
- style: optional dict -- keys: bold, italic, underline, font_color, bg_color, font_size, font_name
  Example: {{"bold": true, "bg_color": "1F3864", "font_color": "FFFFFF", "font_size": 11}}
- number_format: optional Excel format string e.g. "#,##0.00", "DD-MMM-YYYY", "0.0%"
- alignment: optional dict -- keys: horizontal, vertical, wrap_text, shrink_to_fit,
  text_rotation, indent
Apply consistent style across ALL cells in the data range in a single call.
Requires EXCEL_ENABLE_WRITE=true and confirm=True.

## Apply style to one or more ranges (no data write)
Call `apply_style`:
- path, sheet
- ranges: list of A1 range strings -- single range: ["A1:H1"] or multi: ["A1:H1", "A3:H3", "A5:H5"]
- bold, italic, font_size, font_color, bg_color -- all optional; omit unchanged properties
  font_color / bg_color must be 6-hex no #, e.g. "FF0000"
Useful for: header row styling, alternating row banding, KPI cell emphasis.
More efficient than calling apply_style once per range for non-contiguous groups.
Requires EXCEL_ENABLE_WRITE=true and confirm=True.

## Copy formatting WITHOUT values (cross-sheet or same-sheet)
Call `copy_range_format`:
- path
- src_sheet, src_address: source range to copy formatting FROM
- dst_sheet, dst_address: destination (can be a single cell -- auto-expands)
- copy_border: false (default) -- set true to also copy border styles
Copies: font, fill, number_format, alignment.
Does NOT copy: values, formulas, comments, or hyperlinks.
MergedCell slave cells in src are silently skipped.
Use case: copy a styled header template from a _Template sheet to all project sheets.
Requires EXCEL_ENABLE_WRITE=true and confirm=True.

Decision guide:
- New data + styling together -> `write_range_with_format`
- Styling existing cells in multiple non-contiguous ranges -> `apply_style`
- Reapply a template's formatting to a destination without overwriting values -> `copy_range_format`

Always call `save` after styling operations."""


@mcp.prompt()
def annotate_workbook(items_json: str) -> str:
    """Add cell comments and hyperlinks to annotate a set of cells.

    Orchestrates: bulk_add_comments, cell_comment, add_hyperlink, bulk_add_hyperlinks,
    remove_hyperlink, get_hyperlinks.

    items_json: JSON array of {cell, comment, url} objects.
    """
    return f"""You must annotate a workbook with comments and hyperlinks.

Items to annotate (JSON array of {{cell, comment, url}}):
{items_json}

Follow this sequence:
1. Parse the items_json above. For each item extract: cell address, comment text, optional URL.
2. Inspect existing annotations first:
   - Call `get_hyperlinks` to see all current hyperlinks on the target sheet.
   - Call `get_cell_comment` (address=<cell>) for any cell where you suspect an existing comment.
   - If outdated: call `remove_hyperlink` or `delete_cell_comment` to clean up first.
3. Add comments in bulk (preferred for multiple cells):
   Call `bulk_add_comments` with a list: [{{"address": "A1", "text": "...", "author": "MCP"}}].
   For a single cell, use `bulk_add_comments` with a single-item list.
4. Add hyperlinks in bulk (preferred for multiple cells):
   Call `bulk_add_hyperlinks` with a list:
   [{{"address": "A1", "url": "https://...", "display_text": "Link", "tooltip": "..."}}].
   Only https, http, and mailto schemes are permitted -- no file:// unless allowlisted.
   For a single cell, use `add_hyperlink`.
5. Verify: call `get_cell_comment` on a sample cell to confirm the comment was saved.
6. Call `save` to persist all annotations.

Best practices:
- Keep comment text concise (1-3 lines). Long comments truncate in Excel's comment box.
- Use display_text for hyperlinks -- the raw URL as cell value is hard to read.
- Set tooltip on hyperlinks to give users a hover preview of the destination.
- Bulk operations are more efficient than one-by-one calls for 5+ items."""
