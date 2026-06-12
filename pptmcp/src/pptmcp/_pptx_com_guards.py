"""
COM availability guard and application lifecycle helpers for pptmcp.

Private module — not part of the public API.
Imported by: presentation_com, _pptx_com_export.
"""
from __future__ import annotations

import contextlib
import logging
import os
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Generator

from pptmcp.presentation_pptx import NotAllowedError, OfficeCOMError

__all__: list[str] = []

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# COM availability guard
# ---------------------------------------------------------------------------

def _check_com_enabled() -> None:
    """Raise NotAllowedError if COM automation is not available or not enabled.

    Checks, in order:
      1. PPT_ENABLE_COM env var must be exactly "true".
      2. Must be running on win32 platform.
      3. win32com.client must be importable (pywin32 installed).
    """
    if os.getenv("PPT_ENABLE_COM", "").strip().lower() != "true":
        raise NotAllowedError(
            "COM automation requires PPT_ENABLE_COM=true"
        )
    if sys.platform != "win32":
        raise NotAllowedError(
            "COM automation is only supported on Windows (sys.platform='win32')"
        )
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        raise NotAllowedError(
            "pywin32 is not installed. Run: pip install pywin32"
        )


# ---------------------------------------------------------------------------
# COM application lifecycle context manager
# ---------------------------------------------------------------------------

@contextmanager
def _ppt_app_context(visible: bool = False) -> Generator[Any, None, None]:
    """Context manager that opens a transient PowerPoint COM application.

    Guarantees app.Quit() in the finally block so no POWERPNT.EXE process
    is left behind after export/recalculate operations.

    Usage:
        with _ppt_app_context(visible=False) as app:
            prs = app.Presentations.Open(...)

    Addresses B-10 (visibility set immediately; inherent COM flash limitation noted).
    BUG-2: com_error raised by Dispatch() during __enter__ is caught here and
    re-raised as OfficeCOMError so callers always see a clean error, not a raw
    pywintypes.com_error that bypasses FastMCP error handling.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = None
    try:
        try:
            app = win32com.client.Dispatch("PowerPoint.Application")
        except pythoncom.com_error as e:
            hresult = hex(e.hresult) if hasattr(e, "hresult") else "0x0"
            logger.error("Failed to launch PowerPoint.Application. HRESULT: %s", hresult)
            raise OfficeCOMError(
                "Failed to launch PowerPoint.Application — check that PowerPoint is "
                "installed, PPT_ENABLE_COM=true is set, and the process has permission "
                "to launch COM servers."
            ) from None
        try:
            app.DisplayAlerts = False
        except Exception:
            logger.warning("Could not set DisplayAlerts=False; continuing", exc_info=True)
        yield app
    finally:
        if app is not None:
            with contextlib.suppress(Exception):
                app.Quit()
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# Slide-show running check
# ---------------------------------------------------------------------------

def _is_slideshow_running(resolved: Path) -> bool:
    """Return True if a slide show is already running for *resolved*.

    Uses GetActiveObject to attach to an existing PowerPoint instance without
    spawning a new process. Returns False on any error (no existing instance,
    COM unavailable, etc.).

    Addresses B-4 (DoS / multiple-instance guard).
    """
    try:
        import win32com.client

        existing = win32com.client.GetActiveObject("PowerPoint.Application")
        target = str(resolved).lower()
        count = existing.Presentations.Count
        for i in range(1, count + 1):
            try:
                prs = existing.Presentations.Item(i)
                if str(Path(prs.FullName).resolve()).lower() == target:
                    return True
            except Exception:
                continue
        return False
    except Exception:
        return False
