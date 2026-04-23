"""Workflow prompts: design, audit, chart, validation, pivot, recalculate, grow, read/write,
discover, export, configure.

Registers 12 @mcp.prompt() on the shared mcp instance at import time.
"""
from __future__ import annotations

import logging

from excelmcp._server_instance import mcp

logger = logging.getLogger(__name__)


@mcp.prompt()
def design_workbook(business_context: str) -> str:
    """Plan a new workbook structure (sheets, columns, types) ready for create_workbook."""
    return (
        f"Business context:\n{business_context}\n\n"
        "Design a workbook structure for this use case. Output a JSON object with:\n"
        '{"sheet_names": ["Sheet1", ...], "sheets": {"Sheet1": {"columns": [{"name": "...", "type": "string|number|date|boolean", "example": "..."}], "description": "..."}}}\n\n'
        "Choose sheet names that are short, descriptive, and valid Excel sheet names (no special characters, max 31 chars). "
        "Match column types to the data. Include 2-3 example values per column."
    )


@mcp.prompt()
def audit_workbook(path: str) -> str:
    """Discovery prompt: orchestrate all resources to understand a workbook before editing.

    Call sequence: READ excelmcp://workbook/{path}/sheets THEN excelmcp://workbook/{path}/tables
    THEN excelmcp://workbook/{path}/charts THEN excelmcp://workbook/{path}/named_ranges.
    Summarise what you find: sheet inventory, table structure, chart types, named ranges.
    Identify the main data table(s) and suggest next operations.
    """
    return f"""You are auditing the Excel workbook at: {path}

Follow this exact sequence:
1. Read resource: excelmcp://workbook/{path}/sheets \u2014 get all sheet names and dimensions
2. Read resource: excelmcp://workbook/{path}/tables \u2014 get all Excel Tables with column names
3. Read resource: excelmcp://workbook/{path}/charts \u2014 get chart inventory
4. Read resource: excelmcp://workbook/{path}/named_ranges \u2014 get defined names

Then provide:
- Sheet inventory table (name, dimensions, tables present, charts present)
- For each table: name, sheet, ref, column list
- For each chart: type, anchor cell, data source
- Named ranges with their destination refs
- Recommended next action (edit, format, chart, validate \u2014 based on what you found)

Use only the resource data above \u2014 do not call range_io or read_sheet_all during this audit."""


@mcp.prompt()
def chart_from_data(sheet_json: str, insight: str) -> str:
    """Guide for creating a chart from sheet data.

    sheet_json: JSON output from read_sheet_all.
    insight: the business insight to visualise.
    """
    return f"""You must create an Excel chart to visualise: {insight}

Sheet data (from read_sheet_all):
{sheet_json}

Follow this sequence:
1. Identify the columns that best express the insight \u2014 pick at most 2 data series
2. Choose chart type:
   - Time trend (date column + 1 numeric) \u2192 line chart (chart_type="line")
   - Category comparison (text labels + 1-2 numeric) \u2192 bar chart (chart_type="bar")
   - Part-to-whole \u2192 pie chart (chart_type="pie")
   - Correlation (2 numeric series) \u2192 scatter chart (chart_type="scatter")
3. Identify the data range (include the header row)
4. Call: excelmcp chart create with:
   - data_address: "<A1 range including headers>"
   - chart_type: "<chosen type>"
   - title: "<descriptive title>"
   - target_address: "<anchor cell below data, e.g. A20>"
5. Report the chart location and confirm it is visible

One chart per insight. Do not include decorative chart elements."""


@mcp.prompt()
def build_data_validation(field_description: str) -> str:
    """Generate a complete data_validation add tool call spec for a field.

    field_description: describe the field and its allowed values.
    """
    return f"""You must build an Excel data validation rule for: {field_description}

Output a complete JSON spec for the excelmcp data_validation tool (operation="add"):
{{
  "path": "<workbook path>",
  "sheet_name": "<target sheet>",
  "operation": "add",
  "validations": [
    {{
      "address": "<A1 range, e.g. B2:B100>",
      "type": "<whole/decimal/list/date/textLength/custom>",
      "operator": "<between/equal/greaterThan/lessThan/notBetween \u2014 omit for list type>",
      "formula1": "<value or =SheetName!$A$1:$A$10 for list>",
      "formula2": "<upper bound if operator=between, else omit>",
      "allow_blank": true,
      "show_dropdown": true,
      "error_message": "<user-friendly error>",
      "prompt_message": "<what the user should enter>"
    }}
  ]
}}

Rules:
- Use type="list" + formula1="=_Lists!$A$2:$A$20" for dropdown lists from a named range
- Use type="whole" + operator="between" + formula1/2 for integer ranges
- Use type="date" + operator="greaterThan" + formula1="2000-01-01" for date constraints
- Prefer allow_blank=true unless the field is mandatory
- Always set error_message and prompt_message to user-readable text

Provide the complete JSON ready to pass directly as the tool argument."""


@mcp.prompt()
def pivot_workflow(source_description: str) -> str:
    """Step-by-step guide for creating a pivot table.

    source_description: describe the source data and desired pivot.
    """
    return f"""You must create an Excel pivot table for: {source_description}

IMPORTANT: Pivot table creation requires COM (Excel must be installed and file must be .xlsx).
Check: is COM available? If not, skip to the fallback section.

## COM Pathway (preferred)
1. Identify source data range using excelmcp read_sheet_all or the tables resource
2. Call excelmcp pivot_table with operation="create":
   {{
     "path": "<workbook path>",
     "sheet_name": "<source sheet>",
     "operation": "create",
     "source_address": "<data range including headers, e.g. A1:G500>",
     "dest_address": "<upper-left cell of pivot output, on a new or existing sheet>",
     "dest_sheet": "<sheet name for pivot \u2014 create new if needed>",
     "row_fields": ["<field names for row grouping>"],
     "col_fields": ["<field names for column grouping, or []>"],
     "data_fields": [{{"field": "<numeric field>", "function": "Sum"}}],
     "filter_fields": ["<optional slicer fields>"]
   }}
3. Confirm pivot was created: use read_sheet_all on the dest_sheet to verify output

## Fallback (no COM)
If COM is not available, simulate the pivot with COUNTIFS/SUMIFS formulas:
1. Read the source data range
2. Identify unique values in the grouping field using range_io
3. Use fill_formula_range to write SUMIFS/COUNTIFS formulas summarising the data
4. Format the output range with apply_style for readability"""


@mcp.prompt()
def validate_and_recalculate(path: str) -> str:
    """Validate formula syntax, recalculate the workbook, inspect the used range.

    Orchestrates: validate_formula_syntax, recalculate_workbook, get_used_range,
    evaluate_formula, freeze_panes.
    """
    return f"""You must validate formulas and recalculate a workbook to ensure data integrity.

Workbook: {path}

Follow this sequence:
1. Validate formula syntax BEFORE writing to the workbook:
   For any formula you plan to write, first call `validate_formula_syntax`:
   - formula: "=<your formula string>"
   This is a dry-run check \u2014 it does NOT write to the file. Fix any reported errors before proceeding.
2. Recalculate all formulas in the workbook:
   Call `recalculate_workbook` with:
   - path: "{path}"
   - full_calc: true  -- rebuilds the entire dependency tree (safe, slower)
   REQUIRES: EXCEL_ENABLE_COM=true. If COM is unavailable, skip to step 3 --
   openpyxl cached values may be stale but the file is still readable.
3. Inspect the used data extent:
   Call `get_used_range` on each key sheet to confirm the row/column boundaries are correct.
   Watch for unexpected expansion (extra blank rows with formatting = inflated range).
4. Verify key KPI cells:
   For each KPI formula cell (e.g. summary totals, COUNTIFS, dashboard metrics):
   Call `evaluate_formula`:
   - address: "<KPI cell, e.g. B5>"
   Returns both the formula string and the last-saved cached value.
   If the cached value looks stale (e.g. 0 when data exists), recalculate_workbook must
   be run with COM before the value is trustworthy.
5. Set freeze panes on the primary data sheet if not already set:
   Call `freeze_panes` with freeze_at="A2" (freezes row 1 as a sticky header).
6. Call `save` to persist the recalculated state.

Validation checklist:
- [ ] validate_formula_syntax passed for all new formulas
- [ ] recalculate_workbook completed without error
- [ ] get_used_range shows expected bounds (no ghost rows)
- [ ] evaluate_formula confirmed KPI cells return non-zero/non-null values
- [ ] save called after recalculation"""


@mcp.prompt()
def grow_table(path: str, sheet_name: str, table_name: str) -> str:
    """Read table structure, append new rows with correct column order, verify expansion.

    Orchestrates: list_tables, read_table, append_rows, get_used_range.
    """
    return f"""You must append new rows to an existing Excel Table and verify it expanded correctly.

Workbook: {path}
Sheet: {sheet_name}
Table name: {table_name}

Follow this sequence:
1. Discover the table structure:
   Call `list_tables` (path="{path}") to confirm "{table_name}" exists and get its ref range.
2. Read the current table data and schema:
   Call `read_table` (path="{path}", sheet="{sheet_name}", table_name="{table_name}").
   Note the headers list \u2014 NEW rows MUST provide values in this exact column order.
3. Prepare new rows:
   Construct each new row as a Python list in the same column order as the headers.
   - Leave cells empty with null or "" (not 0) where the field is optional.
   - Dates: use ISO format strings "YYYY-MM-DD" \u2014 Excel will auto-parse.
   - Formulas: pass the formula string starting with "=" (e.g. "=SUM(B2:D2)").
4. Append the rows:
   Call `append_rows` with:
   - path: "{path}"
   - sheet: "{sheet_name}"
   - rows: [[<row 1 values>], [<row 2 values>], ...]
   append_rows always writes after the last used row in the sheet, which is the correct
   position for table expansion.
5. Verify the table auto-expanded:
   Call `get_used_range` to confirm the last row increased by the number of rows appended.
   Call `read_table` again to confirm the new rows appear in the table output.
6. Call `save` to persist.

Common pitfalls:
- Column ORDER matters \u2014 mismatched columns silently corrupt data. Always read_table first.
- append_rows writes to the sheet, not the table object directly \u2014 Excel auto-expands the
  table boundary when new rows are written immediately below the last table row.
- If the table has a Totals Row, append_rows may write into it. Disable Totals Row first."""


@mcp.prompt()
def read_write_and_inspect_range(path: str, sheet_name: str, address: str) -> str:
    """Read, write, copy, inspect, and find-replace in a cell range.

    Orchestrates: range_io, copy_range, get_cell_metadata, find_replace.
    """
    return f"""You must read, write, copy, and inspect cells in a workbook range.

Workbook: {path}
Sheet: {sheet_name}
Target address: {address}

Choose the right operation for your task:

## Reading a range
Call `range_io` with operation="read":
- path: "{path}", sheet: "{sheet_name}", address: "{address}"
Returns a 2-D data array with row count and column count.
For a single cell, prefer `cell` (operation="read") -- simpler output.
For the entire used range, use `read_sheet_all`.

## Writing data to a range
Call `range_io` with operation="write":
- path: "{path}", sheet: "{sheet_name}", address: "<top-left cell, e.g. B2>"
- values: [[row1col1, row1col2, ...], [row2col1, ...]]
Provide values as a 2-D list. Strings, numbers, dates (ISO), and formulas (starting with =) are accepted.
Requires EXCEL_ENABLE_WRITE=true and confirm=True.

## Copying cells (values + formats) within the same sheet
Call `copy_range`:
- path: "{path}", sheet: "{sheet_name}"
- src_address: "<source range, e.g. A1:D10>"
- dst_address: "<destination top-left cell or matching range>"
Copies both cell values and cell formatting. Does NOT copy comments or hyperlinks.
Dst can be a single cell -- the range is auto-expanded to match src dimensions.

## Inspecting a single cell's formatting and metadata
Call `get_cell_metadata`:
- path: "{path}", sheet: "{sheet_name}", address: "<single cell, e.g. C5>"
Returns: font (bold, italic, size, color), fill (bg_color), number_format,
alignment, any comment text, and any hyperlink URL.
Use this before styling to avoid overwriting existing formatting accidentally.

## Finding and replacing text across a sheet
Call `find_replace`:
- path: "{path}", sheet: "{sheet_name}"
- find_value: "<exact text to find>"
- replace_value: "<replacement text>"
- match_case: false (default) or true for case-sensitive matching
- whole_cell: true (default, exact match) or false for substring replacement
Returns replacement_count.
Requires EXCEL_ENABLE_WRITE=true and confirm=True.
NOTE: find_replace operates on string/text values only -- formulas and numbers are not matched.

Always call `save` after write operations."""


@mcp.prompt()
def discover_server_and_manage_named_ranges(path: str) -> str:
    """Discover server capabilities and manage named ranges in a workbook.

    Orchestrates: capabilities, named_range.
    """
    return f"""You must discover what the Excel MCP server supports and manage named ranges.

Workbook: {path}

## Step 1 -- Discover server capabilities
Call `capabilities` (no parameters required):
Returns the server phase, version, available tool categories, and active feature flags
(e.g. EXCEL_ENABLE_WRITE, EXCEL_ENABLE_COM). Call this at the start of a session to
confirm which write/COM tools are available before building an automation plan.
No confirm or write gate required -- it is always safe to call.

## Step 2 -- Manage named ranges via the unified tool
Call `named_range` with one of three operations:

### List all named ranges
named_range(operation="list", path="{path}")
Returns: {{named_ranges: [{{name, refers_to, scope}}, ...], count: N}}
Review existing names before creating new ones -- prefer reusing existing names.

### Read values from a named range
named_range(operation="read", path="{path}", name="<range name>")
Returns a 2-D list of cell values for the named range.
Use to verify the range resolves correctly and contains expected data.

### Create or update a named range
named_range(
    operation="write",
    path="{path}",
    name="<short_descriptive_name>",      -- no spaces, use underscores, max 255 chars
    sheet="<sheet where the range lives>",
    address="<A1-style range, e.g. A2:A100>",  -- do NOT include sheet prefix here
    confirm=True
)
Requires EXCEL_ENABLE_WRITE=true and confirm=True.
Example: create "rng_StatusList" on "_Lists" sheet at "A2:A20" for use as a
data-validation source.

## Named range best practices
- Prefix lists with "rng_" (e.g. "rng_PriorityOpts") to distinguish from cell names.
- Scope is always workbook-level when created via this tool.
- After writing a named range, call named_range(operation="read") to confirm it resolves.
- Call save after writing named ranges."""


@mcp.prompt()
def export_workbook_report(path: str, sheet_names: str) -> str:
    """Finalise formatting, export each sheet as PDF, and save the workbook.

    Orchestrates: set_print_area, freeze_panes, export_as_pdf, save.

    sheet_names: comma-separated list of sheet names to export.
    """
    return f"""You must finalise and export a workbook report to PDF.

Workbook path: {path}
Sheets to export: {sheet_names}

Follow this sequence:
1. For EACH sheet in [{sheet_names}]:
   a. Call `get_used_range` to determine the data extent (e.g. "A1:H45").
   b. Call `set_print_area` with address=<used range> to restrict printing to data only.
   c. Call `freeze_panes` with freeze_at="A2" if the sheet has a header row.
   d. Call `set_column_widths` with autofit=true to clean up column widths before PDF export.
2. Export individual sheets:
   For each sheet, call `export_as_pdf`:
   - path: "{path}" (absolute source workbook path)
   - scope: "sheet"
   - sheet: "<sheet name>"
   - output_path: "<absolute path to output PDF>/<SheetName>_Report.pdf"
   - quality: 0 (standard quality)
   - confirm: true
   Returns evidence fields including sha256, size_bytes, exported_at, and elapsed_ms.
   REQUIRES: EXCEL_ENABLE_COM=true environment variable must be set.
3. Export the full workbook (all visible sheets):
   Call `export_as_pdf`:
   - path: "{path}" (absolute source workbook path)
   - scope: "workbook"
   - output_path: "<absolute path to output PDF>/<WorkbookName>_Full.pdf"
   - quality: 0
   - confirm: true
4. Call `save` to persist any formatting changes to the source workbook.
5. Confirm each PDF file was freshly created by checking ok=True plus sha256, size_bytes,
   and exported_at in each export response.

Pre-export checklist:
- Verify no sheet names contain characters invalid in file paths (/, \\, *, ?, :, |, <, >).
- Ensure EXCEL_ENABLE_COM=true is set -- PDF export uses Excel COM; openpyxl cannot export PDF.
- output_path must be an absolute .pdf path inside EXCEL_EXPORT_ROOTS."""


@mcp.prompt()
def export_or_hydrate(intent: str) -> str:
    """Export a workbook/sheet to PDF, or hydrate a template with project data.

    Orchestrates: export_as_pdf, hydrate_template.
    """
    return f"""You must export a workbook to PDF or write project data into a template.

Intent: {intent}

Both tools require EXCEL_ENABLE_COM=true -- Excel must be installed.

---

## Export to PDF
Call `export_as_pdf`:

### Export a single sheet
export_as_pdf(
    scope="sheet",
    path="<absolute source workbook path>",
    output_path="<absolute output path ending in .pdf -- must be within EXCEL_EXPORT_ROOTS>",
    sheet="<exact sheet name>",
    quality=0,        -- 0=standard (default), 1=minimum
    confirm=True
)
Returns: {{ok: True, sheet: str, output_path: str, quality: int, sha256: str, size_bytes: int, exported_at: str, elapsed_ms: int}}

### Export the entire workbook (all visible sheets)
export_as_pdf(
    scope="workbook",
    path="<absolute source workbook path>",
    output_path="<absolute output .pdf path>",
    quality=0,
    confirm=True
)
Returns: {{ok: True, output_path: str, quality: int, source: str, sha256: str, size_bytes: int, exported_at: str, elapsed_ms: int}}

Key notes for export_as_pdf:
- Source workbook is opened READ-ONLY -- it is not modified.
- `path` must be an absolute workbook path inside EXCEL_ALLOWLIST_ROOTS.
- `scope` must be exactly "sheet" or "workbook"; other values are rejected.
- Does NOT require EXCEL_ENABLE_WRITE.
- output_path must be within EXCEL_EXPORT_ROOTS (set in env vars).
- output_path must be absolute; relative paths are rejected.
- If a PDF already exists at output_path, excelmcp validates the full request first,
  then removes the stale file immediately before export so returned evidence always
  describes fresh bytes.
- Pre-set print areas with set_print_area before exporting for clean page breaks.

---

## Hydrate a template with project data
Call `hydrate_template` to fill a .xlsx/.xlsm template with data and save as a new file:

hydrate_template(
    template_path="<path to .xlsx/.xlsm template>",
    output_path="<path for new output file -- SaveAs, template is NOT modified>",
    cell_writes=[
        {{"ref": "ProjectName",   "value": "Procurement Tracker Q2"}},
        {{"ref": "A1",            "value": "Header text"}},
        {{"ref": "B2:D4",         "value": [[1,2,3],[4,5,6],[7,8,9]]}},
        {{"ref": "Sheet2!B3",     "value": 42.0}},
    ],
    sheet="<default sheet for cell lookups -- if None, uses first sheet>",
    recalculate=True,
    confirm=True
)

Returns: {{ok: True, output_path: str, writes_applied: int, recalculated: bool}}
Requires: EXCEL_ENABLE_WRITE=true, EXCEL_ENABLE_COM=true, confirm=True.

Security: all string values in cell_writes are checked for formula injection (=, +, -, @
prefix strings). Pass numeric types instead of strings for numbers to avoid the check.

Named range refs (e.g. "ProjectName") are resolved via Excel's COM Names collection --
the template must have the named range defined before calling hydrate_template."""


@mcp.prompt()
def configure_workbook_calculation(path: str) -> str:
    """Control when Excel recalculates formulas -- set auto, manual, or full-calc-on-load.

    Orchestrates: set_workbook_calc_mode.
    """
    return f"""You must configure the calculation mode for a workbook.

Workbook: {path}

Call `set_workbook_calc_mode`:
- path: "{path}"
- calc_mode: one of "auto", "manual", "autoNoTable" -- or None to leave unchanged
- full_calc_on_load: true/false -- or None to leave unchanged
- confirm: True (EXCEL_ENABLE_WRITE=true required)

## Calculation modes explained

**"auto"** (Excel default):
Excel recalculates all formulas immediately whenever any cell changes.
Use for: small workbooks, interactive use, workbooks with few complex formulas.

**"manual"**:
Excel recalculates only when the user presses F9 or when SaveAs is triggered.
Use for: large workbooks with many SUMIFS/COUNTIFS/array formulas where auto-calc
causes noticeable lag during data entry. Remember to press F9 (or call
`recalculate_workbook`) before reading formula results.

**"autoNoTable"**:
Auto-recalculate formulas EXCEPT for those inside data tables (What-If Analysis tables).
Use for: workbooks with large data tables that are expensive to recalculate.

## full_calc_on_load flag

Setting full_calc_on_load=True marks the workbook so that Excel performs a full
dependency-tree recalculation the next time the file is opened. Useful when:
- The workbook was last saved with stale cached values.
- You have volatile functions (TODAY, NOW, RAND, OFFSET, INDIRECT) that need refresh.
- You distribute the workbook and want recipients to see fresh values immediately.

## Recommended settings by workbook type

| Workbook type              | calc_mode    | full_calc_on_load |
|----------------------------|--------------|-------------------|
| Interactive dashboard      | "auto"       | false             |
| Large PM tracker (500+ rows)| "manual"    | true              |
| Template for distribution  | "auto"       | true              |
| Batch automation output    | "manual"     | true              |

After changing to "manual" mode, call `recalculate_workbook` (COM) or
`validate_and_recalculate` flow to force an immediate calculation before saving."""
