"""
Sheet CRUD operations: add/rename/delete/copy sheets, row/column insert/delete,
and CSV import.
"""
from __future__ import annotations

import csv as _csv_module

from ._active_wb import _check_file_not_open_in_excel
from ._core import (
    ValidationError,
    _adjust_table_refs_after_delete,
    _adjust_table_refs_after_insert,
    _audit_log,
    _check_confirm,
    _check_csv_path,
    _check_path,
    _check_write,
    _flush_if_dirty,
    _load_wb,
    _save_and_evict,
    _sync_autofilter_to_table,
)


def add_sheet(
    path: str,
    sheet_name: str,
    position: int | None = None,
    confirm: bool = False,
) -> dict:
    """Create a new sheet in the workbook.

    Returns ``{"ok": True, "sheet_name": sheet_name, "position": int}``.
    """
    _check_write()
    _check_confirm(confirm)
    if not sheet_name:
        raise ValidationError("sheet_name must be non-empty")
    resolved = _check_path(path)
    _check_file_not_open_in_excel(resolved)   # LC-2: block if file is open in Excel
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)
    if sheet_name in wb.sheetnames:
        raise ValidationError(f"Sheet already exists: {sheet_name!r}")
    wb.create_sheet(title=sheet_name, index=position)
    position_out = wb.sheetnames.index(sheet_name)
    _save_and_evict(wb, resolved)
    _audit_log("add_sheet", str(resolved), sheet_name, None, 0)
    return {"ok": True, "sheet_name": sheet_name, "position": position_out}


def rename_sheet(
    path: str,
    old_name: str,
    new_name: str,
    confirm: bool = False,
) -> dict:
    """Rename a sheet from *old_name* to *new_name*.

    Returns ``{"ok": True, "old_name": old_name, "new_name": new_name}``.
    """
    _check_write()
    _check_confirm(confirm)
    if not old_name or not new_name:
        raise ValidationError("Both old_name and new_name must be non-empty")
    resolved = _check_path(path)
    _check_file_not_open_in_excel(resolved)   # LC-2: block if file is open in Excel
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)
    if old_name not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {old_name!r}")
    if new_name in wb.sheetnames:
        raise ValidationError(f"Sheet already exists: {new_name!r}")
    wb[old_name].title = new_name
    _save_and_evict(wb, resolved)
    _audit_log("rename_sheet", str(resolved), new_name, None, 0)
    return {"ok": True, "old_name": old_name, "new_name": new_name}


def delete_sheet(path: str, sheet_name: str, confirm: bool = False) -> dict:
    """Delete a sheet from the workbook.

    Returns ``{"ok": True, "sheet_name": sheet_name}``.
    """
    _check_write()
    _check_confirm(confirm)
    resolved = _check_path(path)
    _check_file_not_open_in_excel(resolved)   # LC-2: block if file is open in Excel
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)
    if sheet_name not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet_name!r}")
    if len(wb.sheetnames) == 1:
        raise ValidationError("Cannot delete the only sheet in a workbook")
    del wb[sheet_name]
    _save_and_evict(wb, resolved)
    _audit_log("delete_sheet", str(resolved), sheet_name, None, 0)
    return {"ok": True, "sheet_name": sheet_name}


def copy_sheet(
    path: str,
    source_sheet: str,
    new_name: str,
    confirm: bool = False,
) -> dict:
    """Copy *source_sheet* to a new sheet named *new_name*.

    Returns ``{"ok": True, "source_sheet": source_sheet, "new_name": new_name}``.
    """
    _check_write()
    _check_confirm(confirm)
    resolved = _check_path(path)
    _check_file_not_open_in_excel(resolved)   # LC-2: block if file is open in Excel
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)
    if source_sheet not in wb.sheetnames:
        raise ValidationError(f"Source sheet not found: {source_sheet!r}")
    if new_name in wb.sheetnames:
        raise ValidationError(f"Sheet already exists: {new_name!r}")
    copied = wb.copy_worksheet(wb[source_sheet])
    copied.title = new_name
    _save_and_evict(wb, resolved)
    _audit_log("copy_sheet", str(resolved), new_name, None, 0)
    return {"ok": True, "source_sheet": source_sheet, "new_name": new_name}


def bulk_copy_sheet(
    path: str,
    source_sheet: str,
    new_names: list[str],
    insert_after: str | None = None,
    confirm: bool = False,
) -> dict:
    """Copy *source_sheet* to multiple new sheets in a single call.

    All validation is performed before any copy takes place (fail-fast atomicity).

    .. note::
        openpyxl ``copy_worksheet`` does NOT copy charts, images, or print
        settings.  Macros and external references are also not preserved.

    Args:
        path: Path to the workbook file.
        source_sheet: Name of the sheet to copy.
        new_names: List of new sheet names to create.  Must be non-empty and
            contain no duplicates.  Each name must be <= 31 characters and must
            not contain any of ``/ \\ * ? [ ]``.
        insert_after: If provided, the new sheets are placed immediately after
            this sheet in worksheet order.  If ``None`` the new sheets are
            appended at the end of the workbook.
        confirm: Must be True to allow writes.

    Returns:
        ``{"ok": True, "source_sheet": str, "sheets_created": int, "names": list[str]}``
    """
    _check_write()
    _check_confirm(confirm)
    if not new_names:
        raise ValidationError("new_names must be non-empty")
    # Validate each name
    _INVALID_CHARS = set("/\\*?[]")
    for name in new_names:
        if len(name) > 31:
            raise ValidationError(
                f"Sheet name too long (max 31 chars): {name!r} ({len(name)} chars)"
            )
        for ch in _INVALID_CHARS:
            if ch in name:
                raise ValidationError(
                    f"Sheet name {name!r} contains invalid character: {ch!r}"
                )
    # Duplicate check within new_names
    seen: set[str] = set()
    for name in new_names:
        if name in seen:
            raise ValidationError(
                f"Duplicate name in new_names: {name!r}"
            )
        seen.add(name)
    resolved = _check_path(path)
    if len(new_names) > 100:
        raise ValidationError(
            f"bulk_copy_sheet: max 100 copies per call (got {len(new_names)})"
        )
    _check_file_not_open_in_excel(resolved)   # LC-2: after arg validation
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)
    if source_sheet not in wb.sheetnames:
        raise ValidationError(f"Source sheet not found: {source_sheet!r}")
    if insert_after is not None and insert_after not in wb.sheetnames:
        raise ValidationError(f"insert_after sheet not found: {insert_after!r}")
    # Check none of new_names already exist in the workbook
    existing = set(wb.sheetnames)
    for name in new_names:
        if name in existing:
            raise ValidationError(f"Sheet already exists: {name!r}")
    # Capture insert position before adding sheets
    insert_idx: int | None = None
    if insert_after is not None:
        insert_idx = wb.sheetnames.index(insert_after) + 1
    # Copy all sheets (appended at end)
    for name in new_names:
        copied = wb.copy_worksheet(wb[source_sheet])
        copied.title = name
    # Reorder if insert_after was requested
    if insert_idx is not None:
        for i, name in enumerate(new_names):
            current_idx = wb.sheetnames.index(name)
            target_idx = insert_idx + i
            if current_idx != target_idx:
                wb.move_sheet(name, offset=target_idx - current_idx)
    _save_and_evict(wb, resolved)
    _audit_log("bulk_copy_sheet", str(resolved), source_sheet, None, len(new_names))
    return {
        "ok": True,
        "source_sheet": source_sheet,
        "sheets_created": len(new_names),
        "names": new_names,
    }


def delete_rows(
    path: str,
    sheet: str,
    start_row: int,
    end_row: int | None = None,
    confirm: bool = False,
) -> dict:
    """Delete rows *start_row* through *end_row* (inclusive, 1-based).

    Returns ``{"ok": True, "rows_deleted": int}``.
    """
    _check_write()
    _check_confirm(confirm)
    if start_row < 1:
        raise ValidationError("start_row must be >= 1")
    if end_row is None:
        end_row = start_row
    if end_row < start_row:
        raise ValidationError("end_row must be >= start_row")
    resolved = _check_path(path)
    _check_file_not_open_in_excel(resolved)   # LC-2: block if file is open in Excel
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)
    if sheet not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb[sheet]
    amount = end_row - start_row + 1
    ws.delete_rows(start_row, amount)
    _adjust_table_refs_after_delete(ws, start_row, amount)  # T2a: fix stale tbl.ref
    _sync_autofilter_to_table(ws)       # T2: keep AF aligned with tbl.ref
    _save_and_evict(wb, resolved)
    _audit_log("delete_rows", str(resolved), sheet, f"{start_row}:{end_row}", amount)
    return {"ok": True, "rows_deleted": amount}


def insert_rows(
    path: str,
    sheet: str,
    row: int,
    count: int = 1,
    confirm: bool = False,
) -> dict:
    """Insert one or more blank rows before the given row number.

    Args:
        path: Path to the workbook file.
        sheet: Sheet name.
        row: Row number (1-based) at which to insert. Existing rows are shifted down.
        count: Number of rows to insert (default 1).
        confirm: Must be True to allow writes.

    Returns:
        {"ok": True, "sheet": ..., "row": ..., "count": ..., "message": ...}
    """
    _check_write()
    _check_confirm(confirm)
    if row < 1:
        raise ValidationError(f"row must be >= 1, got {row!r}")
    if row > 1_048_576:
        raise ValidationError(f"row must be <= 1,048,576 (Excel row limit), got {row}")
    if count < 1:
        raise ValidationError(f"count must be >= 1, got {count!r}")
    if count > 1000:
        raise ValidationError(f"count must be <= 1000, got {count}. Use multiple calls for large inserts.")
    resolved = _check_path(path)
    _check_file_not_open_in_excel(resolved)   # LC-2: block if file is open in Excel
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)
    if sheet not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb[sheet]
    try:
        ws.insert_rows(row, amount=count)
    except (ValueError, TypeError) as exc:
        raise ValidationError(f"insert_rows failed: {exc}") from exc
    _adjust_table_refs_after_insert(ws, row, count)  # T2a: fix stale tbl.ref
    _sync_autofilter_to_table(ws)       # T2: keep AF aligned with tbl.ref
    _save_and_evict(wb, resolved)
    _audit_log("insert_rows", str(resolved), sheet, f"row{row}", count)
    return {
        "ok": True,
        "sheet": sheet,
        "row": row,
        "count": count,
        "message": f"Inserted {count} row(s) before row {row}",
    }


def delete_cols(
    path: str,
    sheet: str,
    start_col: int,
    end_col: int | None = None,
    confirm: bool = False,
) -> dict:
    """Delete columns *start_col* through *end_col* (inclusive, 1-based).

    Returns ``{"ok": True, "cols_deleted": int}``.
    """
    _check_write()
    _check_confirm(confirm)
    if start_col < 1:
        raise ValidationError("start_col must be >= 1")
    if end_col is None:
        end_col = start_col
    if end_col < start_col:
        raise ValidationError("end_col must be >= start_col")
    resolved = _check_path(path)
    _check_file_not_open_in_excel(resolved)   # LC-2: block if file is open in Excel
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)
    if sheet not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb[sheet]
    amount = end_col - start_col + 1
    ws.delete_cols(start_col, amount)
    _save_and_evict(wb, resolved)
    _audit_log(
        "delete_cols", str(resolved), sheet, f"col{start_col}:col{end_col}", amount
    )
    return {"ok": True, "cols_deleted": amount}


def insert_cols(
    path: str,
    sheet: str,
    col: int,
    count: int = 1,
    confirm: bool = False,
) -> dict:
    """Insert one or more blank columns before the given column number.

    Args:
        path: Path to the workbook file.
        sheet: Sheet name.
        col: Column number (1-based) at which to insert. Existing columns are shifted right.
        count: Number of columns to insert (default 1).
        confirm: Must be True to allow writes.

    Returns:
        {"ok": True, "sheet": ..., "col": ..., "count": ..., "message": ...}
    """
    _check_write()
    _check_confirm(confirm)
    if col < 1:
        raise ValidationError(f"col must be >= 1, got {col!r}")
    if col > 16_384:
        raise ValidationError(f"col must be <= 16,384 (Excel column limit), got {col}")
    if count < 1:
        raise ValidationError(f"count must be >= 1, got {count!r}")
    if count > 500:
        raise ValidationError(f"count must be <= 500, got {count}. Use multiple calls for large inserts.")
    resolved = _check_path(path)
    _check_file_not_open_in_excel(resolved)   # LC-2: block if file is open in Excel
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)
    if sheet not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb[sheet]
    try:
        ws.insert_cols(col, amount=count)
    except (ValueError, TypeError) as exc:
        raise ValidationError(f"insert_cols failed: {exc}") from exc
    _save_and_evict(wb, resolved)
    _audit_log("insert_cols", str(resolved), sheet, f"col{col}", count)
    return {
        "ok": True,
        "sheet": sheet,
        "col": col,
        "count": count,
        "message": f"Inserted {count} column(s) before column {col}",
    }


def import_csv_to_sheet(
    path: str,
    csv_path: str,
    sheet_name: str,
    delimiter: str | None = None,
    skip_header: bool = False,
    overwrite: bool = False,
    confirm: bool = False,
) -> dict:
    """Import a CSV/TSV file into *sheet_name*.

    All rows (including the header row) are imported by default so that column
    names are preserved in the sheet.  Set *skip_header=True* to discard the
    first row of the CSV (e.g. when it is a duplicate of an existing header).

    Returns ``{"ok": True, "rows_imported": int, "sheet_name": sheet_name}``.
    ``rows_imported`` counts only the rows actually written to the sheet.

    Note: this operation saves and closes the workbook immediately.
    """
    _check_write()
    _check_confirm(confirm)
    resolved = _check_path(path)
    _check_file_not_open_in_excel(resolved)   # LC-2: block if file is open in Excel
    csv_resolved = _check_csv_path(csv_path)
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)

    if sheet_name in wb.sheetnames:
        if overwrite:
            ws = wb[sheet_name]
            if ws.max_row is not None and ws.max_row > 0:
                ws.delete_rows(1, ws.max_row)
        else:
            raise ValidationError(
                f"Sheet {sheet_name!r} already exists; pass overwrite=True to replace"
            )
    else:
        ws = wb.create_sheet(title=sheet_name)

    rows_imported = 0
    with open(str(csv_resolved), newline="", encoding="utf-8-sig") as fh:
        if delimiter is None:
            reader = _csv_module.reader(fh)
        else:
            reader = _csv_module.reader(fh, delimiter=delimiter)
        for i, row in enumerate(reader):
            if skip_header and i == 0:
                continue  # explicitly skip the first row only when requested
            # Formula injection guard: check every cell before writing.
            # Numeric values (including negatives like "-500" or "+7.0") are safe.
            # Only block strings that START with a formula prefix AND are not numeric.
            for col_idx, cell_val in enumerate(row):
                if isinstance(cell_val, str):
                    stripped = cell_val.lstrip()
                    try:
                        float(stripped)  # "-500", "-1.5", "+7" are numeric — safe
                    except (ValueError, TypeError):
                        if stripped[:1] in ("=", "+", "-", "@"):
                            raise ValidationError(
                                f"Formula injection blocked at row {i + 1}, col {col_idx + 1}: "
                                "CSV values must not start with formula-prefix characters."
                            )
            ws.append(row)
            rows_imported += 1

    _save_and_evict(wb, resolved)
    _audit_log("import_csv_to_sheet", str(resolved), sheet_name, "A1:*", rows_imported)
    return {"ok": True, "rows_imported": rows_imported, "sheet_name": sheet_name}
