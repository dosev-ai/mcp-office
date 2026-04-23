"""VBA macro inspection and execution tools for excelmcp (US-E-029).

Tools registered on shared mcp instance at import time (side-effect import).
Do not instantiate FastMCP here.

Security:
    list_macros: _ensure_com_gate() + _check_path(path) + _com_excel_app()
                 VBProject access error -> ValidationError (VBA trust message)
    run_macro:   _MACRO_NAME_RE validation (OWASP A03 injection guard)
                 + _check_path(path)
                 + single-quote filename guard (BLOCKER-3)
                 + .xlsm-only guard (BLOCKER-4)
                 + _check_write() + _check_confirm(confirm)
                 + _ensure_macros_gate() + _com_excel_app_macros()

Env vars consumed:
    EXCEL_ENABLE_COM      -- gate for all tools in this module
    EXCEL_ENABLE_WRITE    -- gate for run_macro
    EXCEL_ENABLE_MACROS   -- gate for run_macro (in addition to EXCEL_ENABLE_COM)
    EXCEL_ALLOWLIST_ROOTS -- input workbook path validation
"""
from __future__ import annotations

import logging
import re
import time
from typing import Any

from fastmcp.exceptions import ToolError

from excelmcp._com import (
    _close_workbook,
    _com_excel_app,
    _com_excel_app_macros,
    _ensure_com_gate,
    _ensure_macros_gate,
    _open_workbook,
)
from excelmcp._core import (
    NotAllowedError,
    OfficeCOMError,
    ValidationError,
    _check_confirm,
    _check_path,
    _check_write,
)
from excelmcp._server_instance import mcp

logger = logging.getLogger(__name__)

# OWASP A03: validate macro names before interpolating into excel.Run() string.
# Allows simple identifiers ("RunSomething") and module-qualified names
# ("Module1.MyMacro"). Double-dot, spaces, and special chars are rejected.
# DEFECT-2 correction: use strict module.procedure pattern (not just [A-Za-z0-9_.]).
_MACRO_NAME_RE = re.compile(
    r'^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*)?$'
)


# ---------------------------------------------------------------------------
# Tool: list_macros
# ---------------------------------------------------------------------------

@mcp.tool()
def list_macros(
    path: str,
) -> dict[str, Any]:
    """List all VBA components in a workbook via the VBA project object model.

    Enumerates VBProject.VBComponents and returns component names and types.
    Only .xlsm files have a VBA project; passing an .xlsx path raises
    ValidationError immediately.

    Args:
        path: Absolute path to .xlsm workbook. Must be within EXCEL_ALLOWLIST_ROOTS.

    Returns:
        {"ok": True, "path": str, "components": [{"name": str, "type": int}],
         "count": int}

    Requires: EXCEL_ENABLE_COM=true.
    Does NOT require EXCEL_ENABLE_WRITE or EXCEL_ENABLE_MACROS.
    """
    try:
        # COM gate first (early rejection before path validation or COM startup)
        _ensure_com_gate()

        # Validate path against allowlist
        resolved = _check_path(path)

        # .xlsm-only guard — .xlsx files have no executable VBA project
        if resolved.suffix.lower() != ".xlsm":
            raise ValidationError(
                f"list_macros requires an .xlsm file (macro-enabled workbook). "
                f"Got: {resolved.suffix!r}. Only .xlsm files contain a VBA project."
            )

        # Open workbook via COM and enumerate VBProject components
        components: list[dict] = []
        with _com_excel_app() as excel:
            wb = _open_workbook(excel, resolved, read_only=True)
            try:
                # Access VBProject — may raise if Trust access to VBA is disabled
                try:
                    vbp = wb.VBProject
                    vb_components = vbp.VBComponents
                    comp_count = vb_components.Count
                except Exception as vba_exc:
                    # Translate VBProject access failure to a user-friendly ValidationError.
                    # Most likely cause: "Trust access to VBA project object model" disabled.
                    raise ValidationError(
                        "VBA project access denied -- enable 'Trust access to the VBA "
                        "project object model' in Excel > Trust Center > Macro Settings, "
                        f"then retry. (Underlying error: {vba_exc})"
                    ) from None

                # Enumerate components (1-based index is COM convention)
                for i in range(1, comp_count + 1):
                    try:
                        comp = vb_components.Item(i)
                        components.append({
                            "name": comp.Name,
                            "type": comp.Type,
                        })
                    except Exception:
                        # Skip components that raise errors during enumeration
                        continue

            finally:
                _close_workbook(wb, save=False)
                del wb

        return {
            "ok": True,
            "path": str(resolved),
            "components": components,
            "count": len(components),
        }

    except (NotAllowedError, ValidationError):
        raise
    except OfficeCOMError as exc:
        logger.error("[excelmcp] list_macros COM error: %s", exc, exc_info=True)
        raise ToolError(str(exc)) from exc
    except Exception as exc:
        logger.error("[excelmcp] list_macros unexpected error: %s", exc, exc_info=True)
        raise ToolError("list_macros failed -- check server logs for details.") from exc


# ---------------------------------------------------------------------------
# Tool: run_macro
# ---------------------------------------------------------------------------

@mcp.tool()
def run_macro(
    path: str,
    macro_name: str,
    args: list[Any] | None = None,
    confirm: bool = False,
) -> dict[str, Any]:
    """Execute a VBA macro in an .xlsm workbook via Excel COM.

    Args:
        path:       Absolute path to .xlsm workbook. Must be within EXCEL_ALLOWLIST_ROOTS.
        macro_name: VBA Sub or Function name. Valid formats: "MyMacro" or "Module1.MyMacro".
                    Must match ^[A-Za-z_][A-Za-z0-9_]*(\\.[A-Za-z_][A-Za-z0-9_]*)?$ .
        args:       Optional list of positional arguments to pass to the macro.
        confirm:    Must be True (write gate).

    Returns:
        {"ok": True, "path": str, "macro_name": str, "elapsed_ms": int}

    Requires: EXCEL_ENABLE_WRITE=true, EXCEL_ENABLE_COM=true,
              EXCEL_ENABLE_MACROS=true, confirm=True.
    """
    try:
        # DEFECT-2 / OWASP A03: validate macro_name FIRST before any path/COM work
        if not _MACRO_NAME_RE.match(macro_name):
            raise ValidationError(
                "macro_name must be a valid VBA identifier "
                "(letters, digits, underscore only; optionally 'Module.Procedure'). "
                f"Got: {macro_name!r}"
            )

        # Validate path against allowlist
        resolved = _check_path(path)

        # BLOCKER-3: single-quote injection guard on workbook filename
        if "'" in resolved.name:
            raise ValidationError(
                "Workbook filename must not contain single quotes for macro dispatch. "
                f"Rename the file to remove the apostrophe: {resolved.name!r}"
            )

        # BLOCKER-4: .xlsm-only guard
        if resolved.suffix.lower() != ".xlsm":
            raise ValidationError(
                f"run_macro requires an .xlsm file (macro-enabled workbook). "
                f"Got: {resolved.suffix!r}. Convert the workbook to .xlsm format."
            )

        # Write gate + confirm gate
        _check_write()
        _check_confirm(confirm)

        # BLOCKER-5: macro execution gate (EXCEL_ENABLE_MACROS + EXCEL_ENABLE_COM)
        # _ensure_macros_gate() is also called inside _com_excel_app_macros() but
        # calling it here provides early rejection before Excel.Application is launched.
        _ensure_macros_gate()

        # Open workbook, run macro, save, close -- all within macro-enabled COM context
        t0 = time.monotonic()
        with _com_excel_app_macros() as excel:
            wb = _open_workbook(excel, resolved, read_only=False)
            try:
                # Format: "'WorkbookName.xlsm'!MacroName" or "'...xlsm'!Module.Proc"
                macro_ref = f"'{resolved.name}'!{macro_name}"
                try:
                    if args:
                        excel.Run(macro_ref, *args)
                    else:
                        excel.Run(macro_ref)
                    wb.Save()
                except OfficeCOMError:
                    raise
                except Exception:
                    raise OfficeCOMError("COM error during macro execution") from None
            finally:
                _close_workbook(wb, save=False)
                del wb

        elapsed = round((time.monotonic() - t0) * 1000)
        return {
            "ok": True,
            "path": str(resolved),
            "macro_name": macro_name,
            "elapsed_ms": elapsed,
        }

    except (NotAllowedError, ValidationError):
        raise
    except OfficeCOMError:
        raise  # propagate OfficeCOMError (FastMCP / test catches it)
    except Exception as exc:
        logger.error(
            "[excelmcp] run_macro unexpected error for %s!%s: %s",
            path, macro_name, exc, exc_info=True,
        )
        raise ToolError("run_macro failed -- check server logs for details.") from exc
