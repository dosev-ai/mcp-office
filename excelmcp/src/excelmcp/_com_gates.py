"""Gate helpers and guards for excelmcp COM infrastructure.

No MCP imports. No win32com imports at module scope.

Public API:
    _COMErrFallback       -- sentinel exception class (non-Windows compat)
    _ensure_com_gate()    -- checks sys.platform + EXCEL_ENABLE_COM
    _ensure_macros_gate() -- checks _ensure_com_gate + EXCEL_ENABLE_MACROS
    _ensure_com_available() -- lazy win32com/pythoncom import guard
    _guard_cell_value()   -- rejects formula-injection strings (OWASP A03)
"""
from __future__ import annotations

import os
import sys
from typing import Any

from excelmcp._core import NotAllowedError, OfficeCOMError, ValidationError

_FORMULA_PREFIXES: frozenset[str] = frozenset({"=", "+", "-", "@"})

_SINGLE_CELL_RE_STR: str = r"^\$?[A-Z]{1,3}\$?\d+$"

_VALID_PIVOT_ORIENTATIONS: frozenset = frozenset({1, 2, 3, 4})
_VALID_PIVOT_FUNCTIONS: frozenset = frozenset({-4157, -4112, -4106, -4148, -4136, -4111, -4150})
_XL_ROW_FIELD: int = 1
_XL_COL_FIELD: int = 2
_XL_DATA_FIELD: int = 4
_XL_COUNT: int = -4112
_XL_DATABASE: int = 1


class _COMErrFallback(Exception):
    """Sentinel exception class used when pywintypes is not available.

    Never raised in practice — exists only so that ``except _com_err`` in
    ``_com_excel_app`` remains syntactically valid on non-Windows runners
    (where ``pywintypes`` is absent and ``type(None)`` would cause
    ``TypeError: catching classes that do not inherit from BaseException``).
    """


def _ensure_com_gate() -> None:
    """Raise NotAllowedError unless on Windows with EXCEL_ENABLE_COM=true."""
    if sys.platform != "win32":
        raise NotAllowedError(
            "COM tools require Windows (sys.platform='win32'). "
            "Set live=False to use the cross-platform openpyxl path."
        )
    if os.getenv("EXCEL_ENABLE_COM", "").strip().lower() != "true":
        raise NotAllowedError(
            "COM operations require EXCEL_ENABLE_COM=true in the environment"
        )


def _ensure_macros_gate() -> None:
    """Raise NotAllowedError unless on Windows with EXCEL_ENABLE_COM=true AND EXCEL_ENABLE_MACROS=true.

    Call order in run_macro: call this once for early rejection before expensive
    path-check work. _com_excel_app_macros() also calls this internally so
    callers do NOT need to call _ensure_com_gate() separately.
    """
    # Step 1: delegate base COM gate check (platform + EXCEL_ENABLE_COM)
    _ensure_com_gate()
    # Step 2: require explicit macro opt-in
    if os.getenv("EXCEL_ENABLE_MACROS", "").strip().lower() != "true":
        raise NotAllowedError(
            "Macro execution requires EXCEL_ENABLE_MACROS=true in the environment"
        )


def _ensure_com_available() -> tuple[Any, Any]:
    """Lazily import pywin32 modules; raise OfficeCOMError if unavailable."""
    try:
        import pythoncom  # noqa: PLC0415
        import win32com.client  # noqa: PLC0415
    except ImportError as exc:
        raise OfficeCOMError(
            "win32com not available — Excel COM requires pywin32 to be installed"
        ) from exc
    return pythoncom, win32com.client


def _guard_cell_value(value: Any) -> None:
    """Raise ValidationError if value is a formula-injection string.

    Checks strings starting with =, +, -, @ (after stripping leading whitespace).
    Mirrors the guard in _io.py write_cell and write_range.

    Args:
        value: A cell value that is about to be written via COM.

    Raises:
        ValidationError: If value is a string starting with a formula prefix.
    """
    if isinstance(value, str) and value.lstrip()[:1] in _FORMULA_PREFIXES:
        raise ValidationError(
            f"Formula injection blocked in cell value: {value[:20]!r}"
        )
