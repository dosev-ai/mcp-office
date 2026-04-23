"""MCP session tools — read-only active-workbook bridge.

Registers two tools on the shared ``mcp`` instance:
  get_active_workbook_context   metadata for the currently active workbook
  get_all_open_workbooks        list all workbooks open in the running Excel

Both tools require EXCEL_ENABLE_COM=true and filter via EXCEL_ALLOWLIST_ROOTS.
Workbooks outside the allowlist are silently excluded from all responses.
"""
from __future__ import annotations

import logging

from fastmcp.exceptions import ToolError

from excelmcp._active_wb import _com_active_app, _is_in_allowlist
from excelmcp._com import _ensure_com_gate
from excelmcp._core import NotAllowedError, OfficeCOMError, ValidationError
from excelmcp._server_instance import mcp

logger = logging.getLogger(__name__)


@mcp.tool()
def get_active_workbook_context() -> dict:
    """Return name, path, and sheet list for the workbook active in Excel.

    Requires EXCEL_ENABLE_COM=true. Returns only workbooks within
    EXCEL_ALLOWLIST_ROOTS; workbooks outside the allowlist are treated as
    unavailable (status: outside_allowlist).
    """
    try:
        _ensure_com_gate()
        with _com_active_app() as xl_app:
            wb = xl_app.ActiveWorkbook
            if wb is None:
                logger.debug("get_active_workbook_context: no active workbook")
                return {"status": "no_active_workbook", "workbook": None}
            try:
                full_name = wb.FullName
            except Exception as exc:
                logger.warning("[excelmcp] get_active_workbook_context: FullName COM error: %s", exc)
                return {"status": "com_error", "workbook": None, "error": "Could not read workbook metadata"}
            if not _is_in_allowlist(full_name):
                logger.debug(
                    "get_active_workbook_context: active workbook outside allowlist: %s",
                    full_name,
                )
                return {"status": "outside_allowlist", "workbook": None}
            try:
                sheets = [ws.Name for ws in wb.Worksheets]
            except Exception:
                sheets = []
            return {
                "status": "ok",
                "workbook": {
                    "name": wb.Name,
                    "path": full_name,
                    "sheets": sheets,
                },
            }
    except (NotAllowedError, ValidationError):
        raise
    except ToolError:
        raise
    except OfficeCOMError as exc:
        logger.error("[excelmcp] get_active_workbook_context: COM error: %s", exc, exc_info=True)
        raise ToolError(str(exc)) from exc
    except Exception as exc:
        logger.error("[excelmcp] get_active_workbook_context: unexpected error: %s", exc, exc_info=True)
        raise ToolError("An unexpected error occurred -- check server logs for details.") from exc


@mcp.tool()
def get_all_open_workbooks() -> dict:
    """List all workbooks open in the running Excel instance.

    Requires EXCEL_ENABLE_COM=true. Filters to EXCEL_ALLOWLIST_ROOTS only;
    workbooks outside the allowlist are silently excluded (their paths and
    names are not exposed — not even as a count or redacted entry).
    """
    try:
        _ensure_com_gate()
        results: list[dict] = []
        with _com_active_app() as xl_app:
            for wb in xl_app.Workbooks:
                try:
                    full_name = wb.FullName
                except Exception:
                    continue
                if not _is_in_allowlist(full_name):
                    continue  # outside allowlist — silently exclude
                try:
                    sheets = [ws.Name for ws in wb.Worksheets]
                except Exception:
                    sheets = []
                results.append(
                    {
                        "name": wb.Name,
                        "path": full_name,
                        "sheets": sheets,
                    }
                )
        logger.debug("get_all_open_workbooks: %d allowlisted workbooks found", len(results))
        return {"status": "ok", "workbooks": results, "count": len(results)}
    except (NotAllowedError, ValidationError):
        raise
    except ToolError:
        raise
    except OfficeCOMError as exc:
        logger.error("[excelmcp] get_all_open_workbooks: COM error: %s", exc, exc_info=True)
        raise ToolError(str(exc)) from exc
    except Exception as exc:
        logger.error("[excelmcp] get_all_open_workbooks: unexpected error: %s", exc, exc_info=True)
        raise ToolError("An unexpected error occurred -- check server logs for details.") from exc
