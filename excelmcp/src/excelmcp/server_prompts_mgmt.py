"""Management prompts: sheet lifecycle, table reshape, named ranges, lock template,
duplicate sheet.

Registers 5 @mcp.prompt() on the shared mcp instance at import time.
"""
from __future__ import annotations

import logging

from excelmcp._server_instance import mcp

logger = logging.getLogger(__name__)


@mcp.prompt()
def manage_sheets(workbook_description: str) -> str:
    """Audit, rename, reorder, add, and delete sheets in a workbook.

    Orchestrates: list_sheets, add_sheet, rename_sheet, delete_sheet, copy_sheet,
    reorder_sheets.
    """
    return f"""You must audit and reorganise the sheets in a workbook.

Workbook context: {workbook_description}

Follow this sequence:
1. Audit current sheets:
   Call `list_sheets` (path="<path>") -- returns a list of sheet names in order.
   Map out what exists before making any changes.

2. Add new sheets (if required):
   Call `add_sheet`:
   - path, name="<SheetName>", position=<0-based index or -1 for end>
   - confirm=True   (EXCEL_ENABLE_WRITE=true required)
   Avoid creating sheets that already exist -- list_sheets first.

3. Rename a sheet:
   Call `rename_sheet`:
   - path, old_name="<current name>", new_name="<desired name>"
   - confirm=True
   Rules: max 31 chars, no special characters (/ \\ ? * [ ] :), not blank.

4. Reorder sheets:
   Call `reorder_sheets`:
   - path, sheet_order=["Sheet1", "Sheet2", ...]   -- complete list in desired order
   - confirm=True
   Provide ALL sheet names in the desired final order. Sheets omitted from the list
   keep their relative order after the listed sheets.

5. Copy a sheet within the workbook:
   Call `copy_sheet`:
   - path, source_sheet="<name>", dest_sheet="<new name>"
   - position=<0-based index for destination> or -1 for end
   - confirm=True
   Copies the sheet with all values, formulas, and formatting. Hyperlinks and COM
   objects (charts created via COM) are not guaranteed to copy.

6. Delete a sheet (DESTRUCTIVE -- irreversible without a backup):
   Call `delete_sheet`:
   - path, sheet_name="<name>"
   - confirm=True
   WARNING: delete_sheet permanently removes the sheet and all its data. Verify
   list_sheets output before deleting. Only delete sheets that are truly unused.

7. Call `save` after all sheet operations to persist changes.

Post-audit checklist:
- [ ] list_sheets confirms correct sheet names and count
- [ ] No duplicate sheet names exist
- [ ] Sheets are in logical tab order (e.g. README, PM_Log, Dashboard, Views, _Lists)
- [ ] Hidden helper sheets (_Lists, _Config) are at the end and set to invisible:
  Use `set_sheet_visibility` (visible=False) for each helper sheet after confirm."""


@mcp.prompt()
def reshape_table(table_description: str, change_description: str) -> str:
    """Insert or delete rows and columns in a table or sheet range.

    Orchestrates: insert_rows_cols, delete_rows_cols, list_tables, read_table.

    table_description: describe the table or sheet range to change.
    change_description: describe the rows/columns to insert or delete.
    """
    return f"""You must reshape a table or range by inserting or deleting rows/columns.

Table/range context: {table_description}
Requested change: {change_description}

## Before any modification -- read the current state
1. If working with a named Excel Table, call `list_tables` then `read_table` to understand
   the column structure and row count.
2. Call `get_used_range` to know the exact bounds of data before inserting/deleting.

## Insert rows or columns
Call `insert_rows_cols`:
- path: "<workbook path>"
- sheet: "<sheet name>"
- dimension: "rows" or "columns"
- start_index: <1-based row or column number where insertion begins>
- count: <number of rows/columns to insert>
- confirm: True  (EXCEL_ENABLE_WRITE=true required)

Example: insert 3 blank rows above row 5 on "PM_Log":
  insert_rows_cols(path=..., sheet="PM_Log", dimension="rows", start_index=5, count=3, confirm=True)

After insertion:
- Formula references below/right of the insertion point automatically shift (Excel-native).
- Data validation ranges do NOT automatically expand -- update them manually after.
- Named ranges that include the insertion row auto-extend if they are contiguous.

## Delete rows or columns (DESTRUCTIVE)
Call `delete_rows_cols`:
- path: "<workbook path>"
- sheet: "<sheet name>"
- dimension: "rows" or "columns"
- start_index: <1-based row or column number to start deleting at>
- count: <number of rows/columns to delete>
- confirm: True

WARNING: deletion is irreversible. Always back up the workbook before deleting rows with data.
After deletion, formula references to the deleted rows return #REF! -- inspect formulas first.

## After reshaping
1. Call `read_table` (if a named Table is involved) to confirm row count changed correctly.
2. Call `get_used_range` to verify the new extent.
3. If the table has data validation, call `data_validation` with operation="list" to check
   validation ranges still cover the correct rows.
4. Call `save`."""


@mcp.prompt()
def named_range_workflow(intent: str) -> str:
    """End-to-end workflow: list, read, create/update named ranges.

    Orchestrates: named_range (list / read / write).
    """
    return f"""You must manage named ranges in a workbook.

Intent: {intent}

All three operations are provided by the single `named_range` tool:

## 1. List all named ranges
named_range(operation="list", path="<path>")

Returns: {{named_ranges: [{{name, refers_to, scope}}, ...], count: N}}
Always call this first to discover existing names and avoid duplicates.

## 2. Read values from a named range
named_range(operation="read", path="<path>", name="<range name>")

Returns: {{name, refers_to, values: [[...], ...], row_count, col_count}}
Use to verify the range resolves correctly. If refers_to includes a sheet prefix
(e.g. "PM_Log!$A$2:$A$100") the values array will contain the data from that range.

## 3. Create or update a named range
named_range(
    operation="write",
    path="<path>",
    name="<name>",        -- valid Excel name: no spaces, start with letter or _, max 255 chars
    sheet="<sheet>",      -- sheet where the range resides
    address="<A1-range>", -- do NOT include the sheet prefix; tool adds it
    confirm=True
)
Requires EXCEL_ENABLE_WRITE=true.

If the name already exists it is overwritten (same semantics as Excel Name Manager "Edit").

## Naming conventions for this project
| Prefix  | Scope     | Example                |
|---------|-----------|------------------------|
| rng_    | List/range| rng_StatusOptions      |
| cfg_    | Config    | cfg_ProjectID          |
| kpi_    | KPI cell  | kpi_TotalActions       |

## Common use cases

### Data-validation dropdown source
1. Write values to a column in _Lists sheet (use range_io write).
2. Create named range: name="rng_StatusOptions", sheet="_Lists", address="A2:A10"
3. Use formula1="=rng_StatusOptions" in data_validation add call.

### Config cell referencing
1. Write a project ID string to _Config!B2 using range_io.
2. Create named range: name="cfg_ProjectID", sheet="_Config", address="B2"
3. Reference in formulas as =cfg_ProjectID.

Always call `save` after write operations."""


@mcp.prompt()
def lock_template(path: str, editable_ranges: str) -> str:
    """Protect the workbook structure and lock all cells except specified editable ranges.

    Orchestrates: apply_cell_style (locked/unlocked), protect_sheet, protect_workbook.
    Requires EXCEL_ENABLE_WRITE=true.

    editable_ranges: comma-separated A1 ranges that users should be allowed to edit.
    """
    return f"""You must lock a workbook template, protecting structure while leaving data-entry cells editable.

Workbook: {path}
Editable ranges (users CAN edit these): {editable_ranges}

Follow this sequence:

## Step 1 -- Unlock the editable input cells
For each range in [{editable_ranges}]:
   Call `apply_cell_style`:
   - path: "{path}"
   - sheet: "<sheet>"
   - address: "<range>"
   - locked: false
   - confirm: True
   This flags these cells as unlocked in the cell format record.
   IMPORTANT: Unlocking cells has NO effect until the sheet is protected (Step 2).

## Step 2 -- Lock all other cells (lock the full used range first)
   Call `apply_cell_style`:
   - path: "{path}"
   - sheet: "<sheet>"
   - address: "<full used range, e.g. A1:ZZ9999>"
   - locked: true
   - confirm: True
   Then re-apply Step 1 to re-unlock the editable ranges (apply_cell_style merges, not replaces).

## Step 3 -- Protect each sheet
   Call `protect_sheet` for each sheet that contains the template:
   - path: "{path}"
   - sheet: "<sheet name>"
   - password: "<optional password string, or omit for no password>"
   - allow_insert_rows: false
   - allow_insert_columns: false
   - allow_delete_rows: false
   - allow_delete_columns: false
   - allow_sort: false
   - allow_autofilter: true    -- allow filtering, useful for PM tables
   - allow_format_cells: false
   - confirm: True
   After protect_sheet, only cells with locked=false are editable.

## Step 4 -- Protect workbook structure (prevents sheet add/delete/rename)
   Call `protect_workbook`:
   - path: "{path}"
   - password: "<same password as protect_sheet, or omit>"
   - protect_structure: true
   - protect_windows: false    -- window protection is legacy, leave false
   - confirm: True

## Step 5 -- Save
   Call `save` to persist protection settings.

## Verification
   After saving, reopen the workbook check flow by calling `get_cell_metadata` on:
   - One of the editable range cells -- confirm locked=false
   - One of the template formula cells -- confirm locked=true

Note: protection without a password can be removed by any Excel user. For distribution
templates, always set a password and record it securely."""


@mcp.prompt()
def duplicate_sheet_template(path: str, source_sheet: str, new_names: str) -> str:
    """Copy one sheet multiple times with new names for per-project/per-month templating.

    Orchestrates: copy_sheet (one call per new name).

    new_names: comma-separated list of new sheet names to create from the source.
    """
    return f"""You must create multiple copies of a template sheet with distinct names.

Workbook: {path}
Source sheet to copy: {source_sheet}
New sheet names to create: {new_names}

This is a bulk copy operation using repeated `copy_sheet` calls.

## Sequence

1. Verify source sheet exists:
   Call `list_sheets` (path="{path}") and confirm "{source_sheet}" is present.
   Also check that none of the target names in [{new_names}] already exist to avoid errors.

2. For each name in [{new_names}]:
   Call `copy_sheet`:
   - path: "{path}"
   - source_sheet: "{source_sheet}"
   - dest_sheet: "<name from the list>"
   - position: -1   -- append each copy at the end of the tab bar
   - confirm: True  (EXCEL_ENABLE_WRITE=true required)

   copy_sheet copies the full sheet: values, formulas, formatting, print settings,
   freeze panes, data validation. Named ranges that reference only the source sheet
   are NOT automatically duplicated for the new sheet -- update them manually if needed.

3. Verify all copies were created:
   Call `list_sheets` again and confirm the tab count increased by the number of names.

4. Post-copy customisation (optional but recommended):
   For each new sheet, call `range_io` (write) to update any header cell that contains
   the project/month name -- typically cell A1 or the first row. For example:
   range_io(operation="write", path="{path}", sheet="<new name>",
            address="A1", values=[["<new name>"]], confirm=True)

5. Call `save`.

## Sheet name rules
- Max 31 characters
- No characters: / \\ ? * [ ] :
- Not blank
- Not a duplicate of an existing sheet name (case-insensitive on some systems)

If any name in [{new_names}] violates these rules, fix it before calling copy_sheet.
Suggested safe names: "Jan-2025", "Q1_Summary", "Project_Alpha" (underscores, hyphens, dots OK)."""
