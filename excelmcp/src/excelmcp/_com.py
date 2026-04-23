"""Shared COM infrastructure for excelmcp Sprint G COM Phase 2.

No MCP imports. No openpyxl imports. No module-level COM globals.

This module is the thin facade that:
  1. Hosts the COM session context managers (_com_excel_app, _com_excel_app_macros).
  2. Hosts the pivot table COM functions (_create_pivot_table, _refresh_pivot_tables,
     _read_pivot_table) and workbook helpers (_open_workbook, _close_workbook).
  3. Re-exports all public names from _com_gates, _com_exports, and _com_template
     so that existing importers of ``excelmcp._com`` continue to work unchanged.

Public API (all prefixed _ because this is an internal module):
    _ensure_com_gate()        — from _com_gates: checks EXCEL_ENABLE_COM
    _ensure_macros_gate()     — from _com_gates: checks EXCEL_ENABLE_COM + EXCEL_ENABLE_MACROS
    _ensure_com_available()   — from _com_gates: lazy win32com/pythoncom import guard
    _guard_cell_value()       — from _com_gates: rejects formula-injection strings
    _check_export_path()      — from _com_exports: validates PDF output path
    _check_csv_export_path()  — from _com_exports: validates CSV output path
    _sanitize_csv_value()     — from _com_exports: escapes CSV-injection prefixes
    _evaluate_formula_live()  — from _com_template: live COM formula/value evaluation
    _create_pivot_table()     — from _com_template: create a pivot table via Excel COM
    _refresh_pivot_tables()   — from _com_template: refresh all pivot tables via Excel COM
    _read_pivot_table()       — from _com_template: read pivot table structure/data via Excel COM
    _com_excel_app()          — context manager for COM Excel.Application lifecycle
    _com_excel_app_macros()   — context manager with macro execution enabled
    _open_workbook()          — opens a workbook via COM (inside _com_excel_app)
    _close_workbook()         — closes a COM workbook safely

Security requirements:
    - win32com/pythoncom are NEVER imported at module scope.
    - _com_excel_app() sets AutomationSecurity=3 (W2 B1) before yielding.
    - _check_export_path() enforces EXCEL_EXPORT_ROOTS + .pdf extension.
    - _guard_cell_value() mirrors _io.py write_cell formula-injection guard.

COM lifecycle rule (from _da_com.py — do NOT deviate):
    CoInitialize → Dispatch → ops → Quit → excel=None → CoUninitialize
    The excel=None BEFORE CoUninitialize is mandatory to prevent GC Release crash.
"""
from __future__ import annotations

import logging
from contextlib import contextmanager
from typing import Any, Generator

from excelmcp._com_exports import (
    _CSV_FORMULA_PREFIXES,  # noqa: F401
    _check_csv_export_path,  # noqa: F401
    _check_export_path,  # noqa: F401
    _sanitize_csv_value,  # noqa: F401
)
from excelmcp._com_gates import (
    _COMErrFallback,
    _FORMULA_PREFIXES,  # noqa: F401
    _XL_COL_FIELD,  # noqa: F401
    _XL_COUNT,  # noqa: F401
    _XL_DATABASE,  # noqa: F401
    _XL_DATA_FIELD,  # noqa: F401
    _XL_ROW_FIELD,  # noqa: F401
    _VALID_PIVOT_FUNCTIONS,  # noqa: F401
    _VALID_PIVOT_ORIENTATIONS,  # noqa: F401
    _ensure_com_available,
    _ensure_com_gate,
    _ensure_macros_gate,
    _guard_cell_value,  # noqa: F401
)
from excelmcp._com_template import (
    _create_pivot_table,  # noqa: F401
    _evaluate_formula_live,  # noqa: F401
    _read_pivot_table,  # noqa: F401
    _refresh_pivot_tables,  # noqa: F401
)
from excelmcp._core import OfficeCOMError

_log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Shared COM startup helper (DRY for _com_excel_app and _com_excel_app_macros)
# ---------------------------------------------------------------------------

def _excel_startup(
    pythoncom: Any,
    win32c: Any,
    startup_label: str,
) -> tuple[Any, bool]:
    """Initialise COM STA and create an Excel.Application proxy.

    Returns (excel, coinitialized).
    Raises OfficeCOMError on any startup failure.
    Does NOT set Visible/DisplayAlerts/AutomationSecurity — callers do that.
    """
    try:
        pythoncom.CoInitialize()
    except Exception:
        raise OfficeCOMError(f"{startup_label} failed during CoInitialize()") from None

    try:
        excel = win32c.Dispatch("Excel.Application")
    except Exception:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        raise OfficeCOMError(f"{startup_label} failed creating Excel.Application") from None

    import time as _time  # noqa: PLC0415
    _ready_deadline = _time.monotonic() + 15.0
    while True:
        try:
            if excel.Ready:
                break
        except Exception:
            pass  # COM proxy not yet stable — keep polling
        if _time.monotonic() >= _ready_deadline:
            try:
                excel.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            raise OfficeCOMError(
                f"{startup_label}: Excel.Application not ready after 15 s — "
                "timed out waiting for xl_app.Ready"
            )
        _time.sleep(0.25)

    return excel, True


def _excel_configure(excel: Any, startup_label: str, **attrs: Any) -> None:
    """Set attributes on an Excel.Application proxy; raise OfficeCOMError on failure."""
    for attr, value in attrs.items():
        try:
            setattr(excel, attr, value)
        except Exception:
            try:
                excel.Quit()
            except Exception:
                pass
            raise OfficeCOMError(f"{startup_label} failed setting Excel.{attr}") from None


# ---------------------------------------------------------------------------
# COM session context managers
# ---------------------------------------------------------------------------

@contextmanager
def _com_excel_app() -> Generator[Any, None, None]:
    """Context manager for Excel COM application lifecycle.

    Handles STA initialisation, macro security hardening, and cleanup.

    Yields:
        A live win32com Excel.Application handle.

    Raises:
        NotAllowedError: if EXCEL_ENABLE_COM env var is not "true".
        OfficeCOMError: if pywin32 is unavailable or a COM error occurs.
    """
    _ensure_com_gate()
    pythoncom, win32c = _ensure_com_available()

    try:
        import pywintypes as _pywintypes  # noqa: PLC0415
        _com_err: type[BaseException] = _pywintypes.com_error
    except ImportError:
        _com_err = _COMErrFallback

    coinitialized = False
    excel = None
    startup_error_active = False

    try:
        excel, coinitialized = _excel_startup(pythoncom, win32c, "Excel COM startup")
        try:
            _excel_configure(
                excel,
                "Excel COM startup",
                Visible=False,
                DisplayAlerts=False,
                AutomationSecurity=3,  # msoAutomationSecurityForceDisable — W2 B1
            )
        except OfficeCOMError:
            startup_error_active = True
            excel = None
            raise

        try:
            yield excel
        except _com_err as exc:
            _log.error("Excel COM error: %s", exc)
            raise OfficeCOMError(f"Excel COM error: {exc}") from None
    except OfficeCOMError:
        startup_error_active = True
        raise
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None  # null proxy BEFORE CoUninitialize — prevents GC Release crash
        if coinitialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                if not startup_error_active:
                    raise


@contextmanager
def _com_excel_app_macros() -> Generator[Any, None, None]:
    """Context manager for Excel COM application lifecycle with macro execution enabled.

    Identical to _com_excel_app() except:
      1. Calls _ensure_macros_gate() (checks EXCEL_ENABLE_COM + EXCEL_ENABLE_MACROS).
      2. Sets excel.AutomationSecurity = 2 (msoAutomationSecurityByUI) — allows macros.
      3. Saves and restores the original AutomationSecurity value on exit.

    Yields:
        A live win32com Excel.Application handle with macro execution permitted.

    Raises:
        NotAllowedError: if EXCEL_ENABLE_COM or EXCEL_ENABLE_MACROS is not 'true'.
        OfficeCOMError:  if pywin32 unavailable or a COM error occurs.
    """
    _ensure_macros_gate()
    pythoncom, win32c = _ensure_com_available()

    try:
        import pywintypes as _pywintypes  # noqa: PLC0415
        _com_err: type[BaseException] = _pywintypes.com_error
    except ImportError:
        _com_err = _COMErrFallback

    coinitialized = False
    excel = None
    startup_error_active = False
    original_security: int | None = None

    try:
        excel, coinitialized = _excel_startup(
            pythoncom, win32c, "Excel COM macro startup"
        )
        try:
            _excel_configure(
                excel,
                "Excel COM macro startup",
                Visible=False,
                DisplayAlerts=False,
            )
            # Save original security, then set to ByUI (allows macros)
            try:
                original_security = int(excel.AutomationSecurity)
                excel.AutomationSecurity = 2  # msoAutomationSecurityByUI
            except Exception:
                try:
                    excel.Quit()
                except Exception:
                    pass
                raise OfficeCOMError(
                    "Excel COM macro startup failed setting excel.AutomationSecurity"
                ) from None
        except OfficeCOMError:
            startup_error_active = True
            excel = None
            raise

        try:
            yield excel
        except _com_err as exc:
            _log.error("Excel COM (macros) error: %s", exc)
            raise OfficeCOMError(f"Excel COM error during macro execution: {exc}") from None
    except OfficeCOMError:
        startup_error_active = True
        raise
    finally:
        if excel is not None:
            if original_security is not None:
                try:
                    excel.AutomationSecurity = original_security
                except Exception:
                    pass  # best-effort; Excel about to Quit
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None  # null proxy BEFORE CoUninitialize — prevents GC Release crash
        if coinitialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                if not startup_error_active:
                    raise


# ---------------------------------------------------------------------------
# Workbook helpers
# ---------------------------------------------------------------------------

def _open_workbook(excel: Any, path: Any, read_only: bool = False) -> Any:
    """Open a workbook via COM. Must be called inside _com_excel_app()."""
    return excel.Workbooks.Open(str(path), ReadOnly=read_only)


def _close_workbook(wb: Any, save: bool = False) -> None:
    """Close a COM workbook. Call del wb afterwards to release COM reference."""
    try:
        wb.Close(SaveChanges=save)
    except Exception:
        pass
