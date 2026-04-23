"""
Sheet-level operations: thin orchestration layer.

read_sheet_all and find_replace live here; all other sheet operations are
delegated to _sheet_crud and _sheet_props and re-exported for backward
compatibility.
"""
from __future__ import annotations

from types import ModuleType

from ._active_wb import _check_file_not_open_in_excel
from ._core import (
    ValidationError,
    _audit_log,
    _check_confirm,
    _check_path,
    _check_range_size,
    _check_write,
    _flush_if_dirty,
    _get_used_range,
    _load_wb,
    _read_range_cells,
    _save_and_evict,
)
from . import _sheet_crud

# Re-export everything from the two sub-modules so that any code doing
# ``from ._sheet import <name>`` or ``from ._sheet import *`` continues to work.
from ._sheet_props import (  # noqa: F401
    _VALID_CALC_MODES,
    create_table,
    get_sheet_properties,
    protect_workbook,
    set_sheet_properties,
    set_workbook_calc_mode,
)


def _sheet_crud_with_local_guard() -> ModuleType:
    """Keep the historical monkeypatch seam on excelmcp._sheet intact.

    Decomposition moved CRUD implementations into _sheet_crud, but older tests
    and callers patch excelmcp._sheet._check_file_not_open_in_excel. Mirror the
    current _sheet-local guard into _sheet_crud at call time so exception
    propagation stays backward-compatible.
    """
    _sheet_crud._check_file_not_open_in_excel = _check_file_not_open_in_excel
    return _sheet_crud


def add_sheet(
    path: str,
    sheet_name: str,
    position: int | None = None,
    confirm: bool = False,
) -> dict:
    return _sheet_crud_with_local_guard().add_sheet(
        path, sheet_name, position=position, confirm=confirm
    )


def rename_sheet(
    path: str,
    old_name: str,
    new_name: str,
    confirm: bool = False,
) -> dict:
    return _sheet_crud_with_local_guard().rename_sheet(
        path, old_name, new_name, confirm=confirm
    )


def delete_sheet(path: str, sheet_name: str, confirm: bool = False) -> dict:
    return _sheet_crud_with_local_guard().delete_sheet(
        path, sheet_name, confirm=confirm
    )


def copy_sheet(
    path: str,
    source_sheet: str,
    new_name: str,
    confirm: bool = False,
) -> dict:
    return _sheet_crud_with_local_guard().copy_sheet(
        path, source_sheet, new_name, confirm=confirm
    )


def bulk_copy_sheet(
    path: str,
    source_sheet: str,
    new_names: list[str],
    insert_after: str | None = None,
    confirm: bool = False,
) -> dict:
    return _sheet_crud_with_local_guard().bulk_copy_sheet(
        path,
        source_sheet,
        new_names,
        insert_after=insert_after,
        confirm=confirm,
    )


def delete_rows(
    path: str,
    sheet: str,
    start_row: int,
    end_row: int | None = None,
    confirm: bool = False,
) -> dict:
    return _sheet_crud_with_local_guard().delete_rows(
        path, sheet, start_row, end_row=end_row, confirm=confirm
    )


def insert_rows(
    path: str,
    sheet: str,
    row: int,
    count: int = 1,
    confirm: bool = False,
) -> dict:
    return _sheet_crud_with_local_guard().insert_rows(
        path, sheet, row, count=count, confirm=confirm
    )


def delete_cols(
    path: str,
    sheet: str,
    start_col: int,
    end_col: int | None = None,
    confirm: bool = False,
) -> dict:
    return _sheet_crud_with_local_guard().delete_cols(
        path, sheet, start_col, end_col=end_col, confirm=confirm
    )


def insert_cols(
    path: str,
    sheet: str,
    col: int,
    count: int = 1,
    confirm: bool = False,
) -> dict:
    return _sheet_crud_with_local_guard().insert_cols(
        path, sheet, col, count=count, confirm=confirm
    )


def import_csv_to_sheet(
    path: str | None = None,
    csv_path: str | None = None,
    sheet: str | None = None,
    delimiter: str | None = None,
    skip_header: bool = False,
    overwrite: bool = False,
    confirm: bool = False,
    *,
    workbook_path: str | None = None,
    sheet_name: str | None = None,
) -> dict:
    """Backward-compatible CSV import shim for direct ``excelmcp._sheet`` callers."""
    resolved_path = path if path is not None else workbook_path
    if path is not None and workbook_path is not None and path != workbook_path:
        raise ValidationError("path and workbook_path must match when both are provided")
    if resolved_path is None:
        raise TypeError("import_csv_to_sheet() missing required argument: 'path'")

    resolved_sheet = sheet if sheet is not None else sheet_name
    if sheet is not None and sheet_name is not None and sheet != sheet_name:
        raise ValidationError("sheet and sheet_name must match when both are provided")
    if resolved_sheet is None:
        raise TypeError("import_csv_to_sheet() missing required argument: 'sheet'")
    if csv_path is None:
        raise TypeError("import_csv_to_sheet() missing required argument: 'csv_path'")

    return _sheet_crud_with_local_guard().import_csv_to_sheet(
        resolved_path,
        csv_path,
        resolved_sheet,
        delimiter=delimiter,
        skip_header=skip_header,
        overwrite=overwrite,
        confirm=confirm,
    )


def read_sheet_all(path: str, sheet: str) -> dict:
    """Read the entire used range of *sheet*.

    Returns ``{"sheet": sheet, "data": list[list[dict]], "rows": int, "cols": int}``.
    Subject to the MAX_RANGE_CELLS guard.
    """
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    if sheet not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb[sheet]
    used = _get_used_range(ws)
    min_col = used["min_col"]
    min_row = used["min_row"]
    max_col = used["max_col"]
    max_row = used["max_row"]
    n_cells = (max_row - min_row + 1) * (max_col - min_col + 1)
    _check_range_size(n_cells)
    data = _read_range_cells(ws, min_row, min_col, max_row, max_col)
    return {
        "sheet": sheet,
        "data": data,
        "rows": max_row - min_row + 1,
        "cols": max_col - min_col + 1,
    }


def find_replace(
    path: str,
    sheet: str,
    find_value: str,
    replace_value: str,
    match_case: bool = False,
    whole_cell: bool = True,
    confirm: bool = False,
) -> dict:
    """Find and replace string values in *sheet*.

    Returns ``{"ok": True, "replacements_made": int, "cells_modified": [...]}``.
    """
    import re as _re

    _check_write()
    _check_confirm(confirm)
    if isinstance(replace_value, str):
        _stripped_rv = replace_value.lstrip()
        try:
            float(_stripped_rv)  # numeric replacement (e.g. "-500") is safe
        except (ValueError, TypeError):
            if _stripped_rv[:1] in ("=", "+", "-", "@"):
                raise ValidationError(
                    "replace_value must not start with formula-prefix characters (=, +, -, @)"
                )
    resolved = _check_path(path)
    _check_file_not_open_in_excel(resolved)   # LC-2: block if file is open in Excel
    _flush_if_dirty(str(resolved))
    wb = _load_wb(resolved, for_write=True)
    if sheet not in wb.sheetnames:
        raise ValidationError(f"Sheet not found: {sheet!r}")
    ws = wb[sheet]

    find_key = find_value if match_case else find_value.lower()
    replacements_made = 0
    cells_modified: list[str] = []

    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            cell_val = cell.value if match_case else cell.value.lower()
            if whole_cell:
                if cell_val == find_key:
                    cell.value = replace_value
                    replacements_made += 1
                    cells_modified.append(cell.coordinate)
            else:
                if find_key in cell_val:
                    if match_case:
                        cell.value = cell.value.replace(find_value, replace_value)
                    else:
                        cell.value = _re.sub(
                            _re.escape(find_value),
                            replace_value,
                            cell.value,
                            flags=_re.IGNORECASE,
                        )
                    replacements_made += 1
                    cells_modified.append(cell.coordinate)

    _save_and_evict(wb, resolved)
    _audit_log(
        op="find_replace",
        path=str(resolved),
        sheet=sheet,
        range_ref="*",
        rows_affected=replacements_made,
    )
    return {
        "ok": True,
        "replacements_made": replacements_made,
        "cells_modified": cells_modified,
    }
