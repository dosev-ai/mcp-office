"""Active-workbook COM helper for excelmcp.

Provides utilities to interact with a workbook *already open* in a running
Excel instance (as opposed to launching a new, invisible Excel process via
``_com_excel_app()``).

Module layout:
  _ensure_com_available()          lazy win32com/pythoncom import
  _com_active_app()               context manager — attaches to live Excel session
  get_active_workbook_com()       returns dict | None for the ActiveWorkbook
  _is_in_allowlist()              checks EXCEL_ALLOWLIST_ROOTS membership
  _check_file_not_open_in_excel() pre-write conflict guard (LC-2 protection)
"""
from __future__ import annotations

import logging
import os
import re
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Generator

from excelmcp._core import OfficeCOMError, ValidationError, _audit_log

def _is_valid_a1(address: str) -> bool:
    """Return True iff address is a valid Excel A1 cell reference.

    Validates both syntax (letter columns, numeric rows) and Excel limits:
    column ≤ 16384 (XFD) and row ≤ 1048576.
    """
    m = re.match(r'^([A-Za-z]{1,3})(\d{1,7})$', address.strip())
    if not m:
        return False
    col_str, row_str = m.group(1), m.group(2)
    try:
        from openpyxl.utils.cell import column_index_from_string  # noqa: PLC0415
        col_idx = column_index_from_string(col_str.upper())
    except Exception:
        return False
    row_idx = int(row_str)
    return 1 <= col_idx <= 16384 and 1 <= row_idx <= 1048576


logger = logging.getLogger(__name__)


def _ensure_com_available() -> tuple[Any, Any]:
    """Lazily import pythoncom and win32com.client.

    Returns (pythoncom, win32com.client).
    Raises ImportError if pywin32 is not installed.
    """
    import pythoncom  # noqa: PLC0415
    import win32com.client  # noqa: PLC0415

    return pythoncom, win32com.client


@contextmanager
def _com_active_app() -> Generator[Any, None, None]:
    """Attach to the user's running Excel session via GetActiveObject.

    Yields the Excel.Application COM handle. The caller is responsible for
    reading workbook/sheet data; this context manager MUST NOT:
      - Set xl_app.AutomationSecurity  (would alter user's macro security)
      - Call xl_app.Workbooks.Open()   (would run macros at user security level)
      - Call xl_app.Quit()             (would terminate the user's Excel session)

    SECURITY CONSTRAINT: the above three operations are permanently prohibited
    inside this context manager and any code that uses it.
    """
    pythoncom, win32c = _ensure_com_available()
    pythoncom.CoInitialize()
    try:
        try:
            xl_app = win32c.GetActiveObject("Excel.Application")
        except Exception:
            raise OfficeCOMError(
                "No running Excel instance found or COM connection failed"
            ) from None
        try:
            yield xl_app
        finally:
            pass  # NEVER call xl_app.Quit() here
    finally:
        xl_app = None  # null proxy BEFORE CoUninitialize — prevents GC Release crash
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _is_in_allowlist(full_name: str) -> bool:
    """Return True only if *full_name* resolves inside EXCEL_ALLOWLIST_ROOTS.

    Default-deny (returns False) when EXCEL_ALLOWLIST_ROOTS is unset or empty.
    This intentionally differs from _check_path() which raises on missing
    allowlist: read-only session tools silently exclude out-of-scope workbooks
    rather than raising, so a partial COM configuration does not break callers.

    Note: UNC paths (\\\\server\\share\\...) are excluded by design because
    Path.resolve() on Windows cannot equate UNC with a local drive-letter mount.
    """
    roots_env = os.getenv("EXCEL_ALLOWLIST_ROOTS", "").strip()
    if not roots_env:
        return False
    try:
        resolved = Path(full_name).resolve()
    except (ValueError, OSError):
        return False
    for root in (r.strip() for r in roots_env.split(",") if r.strip()):
        try:
            if resolved.is_relative_to(Path(root).resolve()):
                return True
        except (ValueError, TypeError):
            continue
    return False


def get_active_workbook_com() -> dict[str, Any] | None:
    """Return metadata for the workbook currently active in Excel, or None.

    Returns None if COM is disabled, Excel is not running, or no workbook
    is active. Never raises — callers always get a clean None on failure.
    """
    if os.getenv("EXCEL_ENABLE_COM", "").strip().lower() != "true":
        return None
    try:
        with _com_active_app() as xl_app:
            wb = xl_app.ActiveWorkbook
            if wb is None:
                return None
            return {
                "name": wb.Name,
                "path": wb.FullName,
                "sheets": [ws.Name for ws in wb.Worksheets],
            }
    except OfficeCOMError:
        logger.debug("get_active_workbook_com: COM unavailable, returning None")
        return None
    except Exception as exc:
        logger.debug("get_active_workbook_com: unexpected error: %s", exc)
        return None


def _check_file_not_open_in_excel(resolved: Path) -> None:
    """Guard: raise before an openpyxl write if the target file is open in Excel.

    Behaviour by case:
    - EXCEL_ENABLE_COM != "true"          → return silently (guard inapplicable)
    - win32com not importable / non-Win   → return silently (cannot detect)
    - GetActiveObject raises (no Excel)   → return silently (file cannot be open)
    - FullName matches resolved           → raise ValidationError (block write)
    - COM error while iterating Workbooks → raise OfficeCOMError (fail-closed)

    Accepted residual risks (documented):
    - TOCTOU: file may be opened between this check and the openpyxl write.
      The window is sub-second; accepted equivalent to symlink TOCTOU in _check_path().
    - UNC/mapped-drive: Path.resolve() cannot equate \\\\server\\share\\file.xlsx with
      a drive-letter mount of the same share. UNC workbooks may not be detected.
    """
    if os.getenv("EXCEL_ENABLE_COM", "").strip().lower() != "true":
        return
    if sys.platform != "win32":
        return
    try:
        pythoncom, win32c = _ensure_com_available()
    except ImportError:
        return  # pywin32 not installed — guard inapplicable

    pythoncom.CoInitialize()
    try:
        try:
            xl_app = win32c.GetActiveObject("Excel.Application")
        except Exception:
            return  # Excel not running — file cannot be open
        try:
            for wb in xl_app.Workbooks:
                try:
                    if str(Path(wb.FullName).resolve()).lower() == str(resolved.resolve()).lower():
                        raise ValidationError(
                            f"Workbook '{resolved.name}' is currently open in Excel. "
                            "Close it in Excel first, or use a copy of the file."
                        )
                except ValidationError:
                    raise
                except Exception as exc:
                    raise OfficeCOMError(
                        "Could not verify file lock state via COM: "
                        "error while inspecting open workbook properties"
                    ) from exc
        except ValidationError:
            raise
        except OfficeCOMError:
            raise   # preserve detail from inner raise
        except Exception:
            raise OfficeCOMError(
                "Could not verify file lock state via COM"
            ) from None  # COM error during iteration — fail-closed; suppress detail
    finally:
        xl_app = None  # null proxy BEFORE CoUninitialize — prevents GC Release crash
        wb = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _write_cell_via_com(
    resolved: Path,
    sheet: str,
    address: str,
    value: Any,
) -> dict:
    """Write a single cell value to the active Excel workbook via COM.

    Pre-conditions (enforced by caller _io.write_cell before this is called):
    - EXCEL_ENABLE_WRITE=true (_check_write already called)
    - confirm=True (_check_confirm already called)
    - formula injection already blocked
    - resolved path already allowlist-validated (_check_path already called)

    This function adds:
    - EXCEL_ENABLE_COM=true check
    - xl_app.Ready check
    - Active workbook path verification (resolved must match wb.FullName)
    - Sheet exists check
    - ActiveWorkbook.ReadOnly check

    Does NOT call xl_app.Quit().
    Returns {"ok": True, "previous_value": <prior_cell_value>}.
    """
    if os.getenv("EXCEL_ENABLE_COM", "").strip().lower() != "true":
        raise OfficeCOMError("EXCEL_ENABLE_COM is not 'true' \u2014 COM session write unavailable")

    try:
        pythoncom, win32c = _ensure_com_available()
    except ImportError as exc:
        raise OfficeCOMError("pywin32 is not installed \u2014 COM session write unavailable") from exc

    # Advisory-1: defense-in-depth formula guard (before CoInitialize)
    if isinstance(value, str) and value[:1] in ('=', '+', '-', '@'):
        raise ValidationError(f"Formula values rejected: {value[:20]!r}")
    # Advisory-2: validate address before CoInitialize
    if not _is_valid_a1(address):
        raise ValidationError(f"Invalid cell address: {address!r}")

    pythoncom.CoInitialize()
    try:
        try:
            xl_app = win32c.GetActiveObject("Excel.Application")
        except Exception:
            raise OfficeCOMError("No running Excel instance found \u2014 cannot write via COM session") from None

        if not xl_app.Ready:
            raise OfficeCOMError("Excel is busy (xl_app.Ready=False) \u2014 retry after Excel becomes ready")

        wb = xl_app.ActiveWorkbook
        if wb is None:
            raise OfficeCOMError("No active workbook found in the running Excel instance")

        try:
            active_path = str(Path(wb.FullName).resolve()).lower()
        except Exception as exc:
            raise OfficeCOMError("Could not read active workbook path via COM") from exc
        if active_path != str(resolved.resolve()).lower():
            raise ValidationError(
                f"Active workbook '{wb.Name}' does not match requested path '{resolved.name}'. "
                "Ensure the correct workbook is active in Excel, or omit session_mode to use the openpyxl path."
            )

        if wb.ReadOnly:
            raise ValidationError(
                f"Workbook '{wb.Name}' is open as ReadOnly in Excel \u2014 cannot write via COM session"
            )

        try:
            ws = wb.Worksheets(sheet)
        except Exception:
            raise ValidationError(f"Sheet not found in active workbook: {sheet!r}") from None

        try:
            cell_obj = ws.Range(address.upper())
            previous_value = cell_obj.Value
            cell_obj.Value = value
        except OfficeCOMError:
            raise
        except Exception as exc:
            logger.error("COM write failed (write_cell_com): %s", exc)
            raise OfficeCOMError("COM write failed") from exc
        _audit_log("write_cell_com", resolved, sheet, address.upper(), 1)
        return {"ok": True, "previous_value": previous_value}
    finally:
        cell_obj = None  # null proxies in reverse-creation order — prevents GC Release crash
        ws = None
        wb = None
        xl_app = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _write_range_via_com(
    resolved: Path,
    sheet: str,
    start_cell: str,
    values: list[list[Any]],
) -> dict:
    """Write a 2-D values list to the active Excel workbook via COM.

    Pre-conditions (enforced by caller _io.write_range before this is called):
    - EXCEL_ENABLE_WRITE=true, confirm=True, formula injection blocked
    - resolved path already allowlist-validated

    Same COM guards as _write_cell_via_com:
    - EXCEL_ENABLE_COM, xl_app.Ready, path match, ReadOnly, sheet exists

    Writes via ws.Range(start_cell).Resize(n_rows, n_cols).Value = tuple_of_tuples.
    Returns {"ok": True, "cells_written": N}.
    Does NOT call xl_app.Quit().
    """
    if os.getenv("EXCEL_ENABLE_COM", "").strip().lower() != "true":
        raise OfficeCOMError("EXCEL_ENABLE_COM is not 'true' \u2014 COM session write unavailable")

    try:
        pythoncom, win32c = _ensure_com_available()
    except ImportError as exc:
        raise OfficeCOMError("pywin32 is not installed \u2014 COM session write unavailable") from exc

    n_rows = len(values)
    n_cols = max(len(row) for row in values) if values else 0
    if n_cols == 0:
        raise ValidationError("values must contain at least one non-empty row")
    if not _is_valid_a1(start_cell):
        raise ValidationError(f"Invalid start_cell address: {start_cell!r}")
    # Advisory-1: defense-in-depth formula guard (check original values, before CoInitialize)
    for _row in values:
        for _cell_val in _row:
            if isinstance(_cell_val, str) and _cell_val[:1] in ('=', '+', '-', '@'):
                raise ValidationError(f"Formula values rejected in range: {_cell_val[:20]!r}")
    cells_written = sum(len(r) for r in values)

    pythoncom.CoInitialize()
    try:
        try:
            xl_app = win32c.GetActiveObject("Excel.Application")
        except Exception:
            raise OfficeCOMError("No running Excel instance found \u2014 cannot write via COM session") from None

        if not xl_app.Ready:
            raise OfficeCOMError("Excel is busy (xl_app.Ready=False) \u2014 retry after Excel becomes ready")

        wb = xl_app.ActiveWorkbook
        if wb is None:
            raise OfficeCOMError("No active workbook found in the running Excel instance")

        try:
            active_path = str(Path(wb.FullName).resolve()).lower()
        except Exception as exc:
            raise OfficeCOMError("Could not read active workbook path via COM") from exc
        if active_path != str(resolved.resolve()).lower():
            raise ValidationError(
                f"Active workbook '{wb.Name}' does not match requested path '{resolved.name}'. "
                "Ensure the correct workbook is active in Excel, or omit session_mode to use the openpyxl path."
            )

        if wb.ReadOnly:
            raise ValidationError(
                f"Workbook '{wb.Name}' is open as ReadOnly in Excel \u2014 cannot write via COM session"
            )

        try:
            ws = wb.Worksheets(sheet)
        except Exception:
            raise ValidationError(f"Sheet not found in active workbook: {sheet!r}") from None

        try:
            padded = tuple(tuple(row) + (None,) * (n_cols - len(row)) for row in values)
            target_range = ws.Range(start_cell.upper()).Resize(n_rows, n_cols)
            target_range.Value = padded
        except OfficeCOMError:
            raise
        except Exception as exc:
            logger.error("COM write failed (write_range_com): %s", exc)
            raise OfficeCOMError("COM write failed") from exc
        _audit_log("write_range_com", resolved, sheet, start_cell.upper(), cells_written)
        return {"ok": True, "cells_written": cells_written}
    finally:
        target_range = None  # null proxies in reverse-creation order — prevents GC Release crash
        ws = None
        wb = None
        xl_app = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
