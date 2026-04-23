"""COM probe helpers for Excel MCP dynamic-array finalization.

All COM / win32com logic for the DA finalization pipeline lives here.
Zero OOXML string-patching logic — layer-separation rules apply.
"""
from __future__ import annotations

import logging
import os
import tempfile

__all__ = ["probe_excel_reopen"]

_log = logging.getLogger(__name__)


def probe_excel_reopen(path: str) -> dict:
    """Open *path* in Excel COM, recalculate, save, and reopen to verify no repair dialog.

    Returns {"ok": True} on clean open+calc+save+reopen cycle, or
    {"ok": False, "error": "..."} on any failure or if win32com is unavailable.

    Raises NotAllowedError if EXCEL_ENABLE_COM != "true" (COM gate blocked).

    This is a Windows-only integration helper -- safe to call from tests marked
    @pytest.mark.integration, and from _da_finalize.finalize_workbook_da_artifact
    when EXCEL_DA_COM_PROBE=auto|required.
    """
    from ._com import _ensure_com_gate  # noqa: PLC0415
    _ensure_com_gate()
    try:
        import pythoncom  # noqa: PLC0415
        import win32com.client  # noqa: PLC0415
    except ImportError:
        return {"ok": False, "error": "win32com not available"}

    abs_path = os.path.abspath(path)
    # --- allowlist guard (reuses _core._check_path; temp paths always allowed) ---
    _tmp = os.path.normcase(os.path.realpath(tempfile.gettempdir())) + os.sep
    if not os.path.normcase(os.path.realpath(abs_path)).startswith(_tmp):
        roots_env = os.getenv("EXCEL_ALLOWLIST_ROOTS", "").strip()
        if not roots_env:
            return {"ok": False, "error": "EXCEL_ALLOWLIST_ROOTS not configured; non-temp paths are denied by default"}
        from ._core import _check_path  # noqa: PLC0415
        try:
            _check_path(abs_path)
        except Exception as exc:
            return {"ok": False, "error": f"Path validation failed: {exc}"}
    # -------------------------------------------------------------------------
    _fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(_fd)

    # Initialise COM STA for this thread.  Safe to call multiple times -- COM
    # uses a refcount, and CoUninitialize in the finally decrements it again.
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable — W2 B1

        wb = excel.Workbooks.Open(abs_path, CorruptLoad=0)
        wb.Application.CalculateFull()
        wb.SaveAs(temp_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook
        wb.Close(SaveChanges=False)
        del wb

        wb2 = excel.Workbooks.Open(temp_path, CorruptLoad=0)
        wb2.Close(SaveChanges=False)
        del wb2

        return {"ok": True, "path": path}
    except Exception as exc:
        _log.warning("[excelmcp] probe_excel_reopen failed for %s: %s", abs_path, exc)
        return {"ok": False, "error": repr(exc), "path": path}
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None  # Null proxy before CoUninitialize to prevent GC Release crash
        if os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except Exception:
                pass
        pythoncom.CoUninitialize()
