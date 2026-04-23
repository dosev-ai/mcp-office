"""Workbook recalculation and template hydration tools.

Tools registered on shared mcp instance at import time (side-effect import).
Do not instantiate FastMCP here.

Split from server_com_pivot.py — pivot table ops live in server_com_pivot.py.

SECURITY: _validate_cell_writes lives HERE AND ONLY HERE.

Env vars consumed:
    EXCEL_ENABLE_WRITE    -- gate for recalculate_workbook, hydrate_template
    EXCEL_ENABLE_COM      -- gate for all tools
    EXCEL_ALLOWLIST_ROOTS -- input workbook and .xlsx/.xlsm output path root
"""
from __future__ import annotations

import logging
import time
from typing import Any

from fastmcp.exceptions import ToolError

from excelmcp._com import (
    _close_workbook,
    _com_excel_app,
    _ensure_com_gate,
    _guard_cell_value,
    _open_workbook,
)
from excelmcp._core import (
    _NAMED_RANGE_RE,
    _RANGE_ADDRESS_RE,
    NotAllowedError,
    OfficeCOMError,
    ValidationError,
    _check_confirm,
    _check_path,
    _check_write,
)
from excelmcp._server_instance import mcp

logger = logging.getLogger(__name__)


def _classify_com_error(exc: OfficeCOMError) -> str:
    """Return a user-visible hint classifying a COM error.

    Pure function — no COM calls. Maps exact OfficeCOMError message substrings
    emitted by _com.py and _com_template.py to human-readable hint strings.
    """
    msg = str(exc)
    if "CoInitialize" in msg:
        return "COM initialisation failed — is Excel installed?"
    if "creating Excel.Application" in msg:
        return "Excel.Application could not be dispatched — is Excel installed and licensed?"
    if "not ready after 15 s" in msg:
        return "Excel did not become ready in time — close other Excel instances and retry"
    if "Excel is busy" in msg:
        return "Excel is currently busy — retry after the active operation completes"
    if "PivotCache" in msg:
        return "PivotCache creation failed — check source_address range and data"
    if "PivotTable" in msg and "create" in msg.lower():
        return "PivotTable creation failed — check dest_sheet and dest_address"
    if "SourceType" in msg:
        return "Cannot read pivot SourceType — workbook may be corrupt or pivot is stale"
    return "Excel COM error — check server logs for details"


# ---------------------------------------------------------------------------
# SECURITY: _validate_cell_writes lives HERE AND ONLY HERE
# ---------------------------------------------------------------------------

def _validate_cell_writes(cell_writes: list[dict]) -> None:
    """Validate all cell_writes items BEFORE entering a COM context.

    Checks:
      1. Each item has "ref" and "value" keys.
      2. "ref" matches _RANGE_ADDRESS_RE or _NAMED_RANGE_RE.
      3. "value" (scalar or 2D array) contains no formula-injection strings.

    Raises ValidationError on first violation.
    Must be called before _com_excel_app() so invalid input is rejected
    without spawning an Excel process.
    """
    for i, write in enumerate(cell_writes):
        if "ref" not in write or "value" not in write:
            raise ValidationError(
                f"cell_writes[{i}] must have 'ref' and 'value' keys"
            )
        ref = write["ref"]
        if not ref or not (
            _RANGE_ADDRESS_RE.match(ref) or _NAMED_RANGE_RE.match(ref)
        ):
            raise ValidationError(
                f"cell_writes[{i}]: invalid ref {ref!r} -- must be a cell address "
                "(e.g. 'A1', 'A1:C3') or named range (e.g. 'MyRange')"
            )
        value = write["value"]
        if isinstance(value, list):
            for row in value:
                if isinstance(row, list):
                    for cell_val in row:
                        _guard_cell_value(cell_val)
                else:
                    _guard_cell_value(row)
        else:
            _guard_cell_value(value)


# ---------------------------------------------------------------------------
# Tool T1: recalculate_workbook
# ---------------------------------------------------------------------------

@mcp.tool()
def recalculate_workbook(
    path: str,
    full_calc: bool = False,
    confirm: bool = False,
) -> dict[str, Any]:
    """Force recalculation of all formulas in a workbook using Excel COM.

    Opens the workbook in Excel, triggers recalculation, saves in-place, and
    closes. The source workbook is modified (formulas are evaluated and their
    cached values updated in the file).

    Args:
        path: Absolute path to the .xlsx or .xlsm workbook.
              Must be within EXCEL_ALLOWLIST_ROOTS.
        full_calc: If True, calls Application.CalculateFull() which rebuilds
                   the entire dependency tree across all open workbooks.
                   If False (default), calls wb.Calculate() which recalculates
                   only dirty cells in this workbook (faster).
        confirm: Must be True. Required because this operation saves the
                 source workbook in-place (permanent mutation).

    Returns:
        {"ok": True, "path": <resolved>, "calc_mode": "full_recalc"|"recalc",
         "elapsed_ms": <int>}

    Requires: EXCEL_ENABLE_WRITE=true, EXCEL_ENABLE_COM=true, confirm=True.
    """
    try:
        _ensure_com_gate()
        _check_write()
        _check_confirm(confirm)
        resolved = _check_path(path)

        t0 = time.monotonic()
        with _com_excel_app() as excel:
            wb = _open_workbook(excel, resolved, read_only=False)
            try:
                wb.Activate()
            except Exception as _act_exc:
                logger.error("wb.Activate() failed: %s", _act_exc)
                raise OfficeCOMError(f"Failed to activate workbook: {_act_exc}") from _act_exc
            try:
                if full_calc:
                    wb.Application.CalculateFull()
                else:
                    wb.Calculate()
                wb.Save()
            finally:
                _close_workbook(wb, save=False)
                del wb

        elapsed = round((time.monotonic() - t0) * 1000)
        calc_mode = "full_recalc" if full_calc else "recalc"
        return {
            "ok": True,
            "path": str(resolved),
            "calc_mode": calc_mode,
            "elapsed_ms": elapsed,
        }
    except (NotAllowedError, ValidationError) as exc:
        raise ToolError(f"Operation failed ({type(exc).__name__})") from exc
    except OfficeCOMError as exc:
        logger.error("[excelmcp] recalculate_workbook COM error: %s", exc, exc_info=True)
        raise ToolError(f"Recalculation failed — {_classify_com_error(exc)}") from exc
    except Exception as exc:
        logger.error("[excelmcp] recalculate_workbook unexpected error: %s", exc, exc_info=True)
        raise ToolError("An unexpected error occurred -- check server logs for details.") from exc


# ---------------------------------------------------------------------------
# Tool T7a: hydrate_template
# ---------------------------------------------------------------------------

@mcp.tool()
def hydrate_template(
    template_path: str,
    cell_writes: list[dict] | None = None,
    output_path: str = "",
    sheet: str | None = None,
    recalculate: bool = True,
    confirm: bool = False,
) -> dict[str, Any]:
    """Open a workbook template, write cell values, and save as a new file.

    Writes data to named ranges or cell addresses in the template workbook
    and saves the result to output_path. The template is not modified
    (SaveAs creates a new file at output_path).

    cell_writes format (list of dicts with 'ref' and 'value' keys)::

        [
            {"ref": "A1",          "value": "scalar"},
            {"ref": "A1:C2",       "value": [[1, 2, 3], [4, 5, 6]]},
            {"ref": "named_range", "value": "ProjectX"},
            {"ref": "Sheet2!B3",   "value": 42.0},
        ]

    Args:
        template_path: Absolute path to source .xlsx/.xlsm template.
        cell_writes:   List of {ref, value} write instructions.
        output_path:   Absolute path for output .xlsx/.xlsm file.
        sheet:         Optional default sheet for Range() lookups.
        recalculate:   If True (default), calls Application.CalculateFull() before SaveAs.
        confirm:       Must be True.

    Returns:
        {"ok": True, "output_path": str, "writes_applied": int, "recalculated": bool}

    Requires: EXCEL_ENABLE_WRITE=true, EXCEL_ENABLE_COM=true, confirm=True.
    """
    try:
        _ensure_com_gate()                      # COM gate -- FIRST
        _check_write()                          # C2a gate
        _check_confirm(confirm)
        resolved_tmpl = _check_path(template_path)
        resolved_out = _check_path(output_path)
        if cell_writes is None:
            cell_writes = []
        _validate_cell_writes(cell_writes)      # injection guard -- before COM

        file_format = 52 if resolved_out.suffix.lower() == ".xlsm" else 51

        with _com_excel_app() as excel:
            wb = _open_workbook(excel, resolved_tmpl, read_only=False)
            try:
                wb.Activate()
            except Exception as _act_exc:
                logger.error("wb.Activate() failed: %s", _act_exc)
                raise OfficeCOMError(f"Failed to activate workbook: {_act_exc}") from _act_exc
            try:
                if sheet is not None:
                    sheet_names = [
                        wb.Sheets(i + 1).Name for i in range(wb.Sheets.Count)
                    ]
                    if sheet not in sheet_names:
                        raise ValidationError(
                            f"Sheet {sheet!r} not found in template workbook "
                            f"(available: {sheet_names})"
                        )
                    target_ws = wb.Sheets(sheet)
                else:
                    target_ws = wb.Sheets(1)

                name_list: list[str] = [
                    wb.Names(i + 1).Name for i in range(wb.Names.Count)
                ]

                for write in cell_writes:
                    ref: str = write["ref"]
                    value: Any = write["value"]
                    if ref in name_list:
                        wb.Names(ref).RefersToRange.Value = value
                    else:
                        try:
                            target_ws.Range(ref).Value = value
                        except Exception as rng_exc:
                            raise ValidationError(
                                f"Range lookup failed for ref {ref!r} ({type(rng_exc).__name__})"
                            ) from rng_exc

                if recalculate:
                    wb.Application.CalculateFull()

                wb.SaveAs(str(resolved_out), FileFormat=file_format)
            finally:
                _close_workbook(wb, save=False)
                del wb

        return {
            "ok": True,
            "output_path": str(resolved_out),
            "writes_applied": len(cell_writes),
            "recalculated": recalculate,
        }
    except (NotAllowedError, ValidationError) as exc:
        raise ToolError(f"Operation failed ({type(exc).__name__})") from exc
    except OfficeCOMError as exc:
        logger.error("[excelmcp] hydrate_template COM error: %s", exc, exc_info=True)
        raise ToolError(f"Template hydration failed — {_classify_com_error(exc)}") from exc
    except Exception as exc:
        logger.error("[excelmcp] hydrate_template unexpected error: %s", exc, exc_info=True)
        raise ToolError("Template hydration failed -- check server logs for details.") from exc
