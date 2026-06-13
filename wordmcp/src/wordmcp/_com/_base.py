"""
Base COM utilities: COM imports, availability flag, error helpers,
COM availability guard, and the _word_app_context lifecycle manager.

No FastMCP/MCP imports. No intra-_com imports.
"""
from __future__ import annotations

import contextlib
import os
import sys

try:
    import pythoncom
    import win32com.client

    _COM_AVAILABLE = True
except ImportError:
    _COM_AVAILABLE = False

from wordmcp.document_docx import (  # noqa: F401
    NotAllowedError,
    OfficeCOMError,
    ValidationError,
    _check_confirm,
    _check_path,
    _check_write,
)

# ---------------------------------------------------------------------------
# COM error type helper
# ---------------------------------------------------------------------------


def _is_com_error(exc: BaseException) -> bool:
    """Return True if exc is a native pythoncom COM error."""
    return (
        _COM_AVAILABLE
        and hasattr(pythoncom, "com_error")
        and isinstance(exc, pythoncom.com_error)
    )


# ---------------------------------------------------------------------------
# COM error message helper — H-007
# ---------------------------------------------------------------------------


def _com_err_msg(exc: Exception) -> str:
    """Extract human-readable message from a COM error, stripping HRESULT tuple.

    pywintypes.com_error args layout: (hresult, msg, excepinfo, argErr).
    Falls back to a generic safe message if the expected format is not present.
    """
    if hasattr(exc, "args") and exc.args:
        args = exc.args
        # com_error: args[1] is the human-readable description string
        if len(args) >= 2 and isinstance(args[1], str):
            return args[1]
    return "COM error (check stderr log for details)"


# ---------------------------------------------------------------------------
# COM availability guard
# ---------------------------------------------------------------------------


def _check_com_enabled() -> None:
    """Raise NotAllowedError if COM automation is not available or not enabled.

    Checks, in order:
      1. WORD_ENABLE_COM env var must be exactly "true".
      2. Must be running on win32 platform.
      3. win32com.client must be importable (pywin32 installed).
    """
    if os.environ.get("WORD_ENABLE_COM", "").lower() != "true":
        raise NotAllowedError("COM automation requires WORD_ENABLE_COM=true")
    if sys.platform != "win32":
        raise NotAllowedError(
            "COM automation is only supported on Windows (sys.platform='win32')"
        )
    if not _COM_AVAILABLE:
        raise NotAllowedError("pywin32 is not installed. Run: pip install pywin32")


# ---------------------------------------------------------------------------
# COM application lifecycle context manager
# ---------------------------------------------------------------------------


@contextlib.contextmanager
def _word_app_context():  # type: ignore[return]
    """Open a transient Word COM application; quit in the finally block.

    If Dispatch fails with CO_E_SERVER_EXEC_FAILURE (Word not running / not
    registered in the COM ROT), spawn WINWORD.EXE from within this process so
    it registers in the same COM session, then retry Dispatch once.
    """
    if not _COM_AVAILABLE:
        raise OfficeCOMError("pywin32 is not installed")
    pythoncom.CoInitialize()
    app = None
    try:
        try:
            app = win32com.client.Dispatch("Word.Application")
        except Exception as exc:
            # CO_E_SERVER_EXEC_FAILURE (-2146959355) means Word isn't running /
            # can't be reached.  Try spawning it from this process, then retry.
            _WORD_EXES = [
                r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
                r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
            ]
            import subprocess as _sp
            import time as _t
            _launched = False
            for _exe in _WORD_EXES:
                if _COM_AVAILABLE:  # we already checked win32com; also check path
                    import pathlib as _pl
                    if _pl.Path(_exe).exists():
                        # H-008: wrap Popen so PermissionError becomes OfficeCOMError
                        try:
                            _sp.Popen([_exe, "/q"])
                        except OSError as spawn_err:
                            raise OfficeCOMError(
                                f"Failed to start Word: {spawn_err}"
                            ) from spawn_err
                        _t.sleep(5)
                        _launched = True
                        break
            if not _launched:
                if hasattr(pythoncom, "com_error") and isinstance(exc, pythoncom.com_error):
                    raise OfficeCOMError(f"COM error (could not start Word): {_com_err_msg(exc)}") from exc
                raise
            # Retry after launch
            try:
                app = win32com.client.Dispatch("Word.Application")
            except Exception as exc2:
                if hasattr(pythoncom, "com_error") and isinstance(exc2, pythoncom.com_error):
                    raise OfficeCOMError(f"COM error after Word launch: {_com_err_msg(exc2)}") from exc2
                raise
        # Setting Visible/DisplayAlerts may raise a COM error when connecting to
        # an existing Word session whose UI restrictions prevent it.  Suppress
        # both — they are cosmetic guards, not functional requirements.
        with contextlib.suppress(Exception):
            app.Visible = False
        with contextlib.suppress(Exception):
            app.DisplayAlerts = 0
        yield app
    except Exception as exc:
        if hasattr(pythoncom, "com_error") and isinstance(exc, pythoncom.com_error):
            raise OfficeCOMError(f"COM error: {_com_err_msg(exc)}") from exc
        raise
    finally:
        if app is not None:
            with contextlib.suppress(Exception):
                app.Quit()
        pythoncom.CoUninitialize()
