"""Pivot table tools.

Tools registered on shared mcp instance at import time (side-effect import).
Do not instantiate FastMCP here.

Split from server_com_pivot.py — recalculate/hydrate ops live in server_com_recalc.py.

Env vars consumed:
    EXCEL_ENABLE_WRITE    -- gate for pivot create/refresh
    EXCEL_ENABLE_COM      -- gate for all tools
    EXCEL_ALLOWLIST_ROOTS -- input workbook path root
"""
from __future__ import annotations

import logging
from typing import Literal, cast

from fastmcp.exceptions import ToolError

from excelmcp._com import (
    _create_pivot_table,
    _evaluate_formula_live,
    _read_pivot_table,
    _refresh_pivot_tables,
)
from excelmcp._core import (
    NotAllowedError,
    OfficeCOMError,
    ValidationError,
)
from excelmcp._server_instance import mcp
from excelmcp.server_com_recalc import (
    _classify_com_error,
    hydrate_template,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Tool: pivot_table (unified T2/T3/T4)
# ---------------------------------------------------------------------------

@mcp.tool()
def pivot_table(
    operation: Literal["create", "read", "refresh"],
    path: str,
    source_address: str | None = None,
    dest_sheet: str | None = None,
    dest_address: str | None = None,
    row_fields: list | None = None,
    col_fields: list | None = None,
    value_fields: list | None = None,
    table_name: str = "PivotTable1",
    sheet: str | None = None,
    pivot_name: str | None = None,
    confirm: bool = False,
) -> dict:
    """Create, read, or refresh Excel pivot tables via Excel COM.

    operation='create': Build a new pivot table from a local worksheet range.
    operation='read':   Return structure and data of a named pivot table.
    operation='refresh': Refresh all pivot caches (xlDatabase sources only).

    Args:
        operation:      'create', 'read', or 'refresh'.
        path:           Absolute path to .xlsx/.xlsm workbook.
        source_address: Sheet-qualified range e.g. 'Sheet1!A1:D100' (create only).
        dest_sheet:     Existing sheet for pivot output (create only).
        dest_address:   Upper-left cell for pivot e.g. 'F1' (create only).
        row_fields:     Field names for the row axis (create only).
        col_fields:     Field names for the column axis (create only).
        value_fields:   Field names to aggregate (create only).
        table_name:     PivotTable name (create only; default 'PivotTable1').
        sheet:          Worksheet name (read: required; refresh: optional filter).
        pivot_name:     PivotTable name to read (read only).
        confirm:        Must be True for create and refresh operations.

    Returns:
        create:  {"ok": True, "path": str, "table_name": str, ...}
        read:    {"ok": True, "path": str, "sheet": str, "pivot_name": str, ...}
        refresh: {"ok": True, "path": str, "refreshed": list, "count": int}
    """
    if operation == "create":
        missing = [
            name for name, val in [
                ("source_address", source_address),
                ("dest_sheet", dest_sheet),
                ("dest_address", dest_address),
                ("row_fields", row_fields),
                ("value_fields", value_fields),
            ] if val is None
        ]
        if missing:
            raise ToolError(
                f"pivot_table(operation='create') missing required params: {missing}"
            )
        try:
            return _create_pivot_table(
                path=path,
                source_range=cast(str, source_address),
                dest_sheet=cast(str, dest_sheet),
                dest_cell=cast(str, dest_address),
                row_fields=cast(list, row_fields),
                col_fields=col_fields or [],
                value_fields=cast(list, value_fields),
                table_name=table_name,
                confirm=confirm,
            )
        except (NotAllowedError, ValidationError) as exc:
            raise ToolError(f"Operation failed ({type(exc).__name__})") from exc
        except OfficeCOMError as exc:
            logger.error("[excelmcp] pivot_table create COM error: %s", exc, exc_info=True)
            raise ToolError(f"Pivot table creation failed — {_classify_com_error(exc)}") from exc
        except Exception as exc:
            logger.error(
                "[excelmcp] pivot_table create unexpected error: %s", exc, exc_info=True
            )
            raise ToolError(f"Pivot table creation failed ({type(exc).__name__})") from exc

    elif operation == "read":
        if sheet is None:
            raise ToolError("pivot_table(operation='read') requires sheet")
        if pivot_name is None:
            raise ToolError("pivot_table(operation='read') requires pivot_name")
        try:
            return _read_pivot_table(path=path, sheet=sheet, pivot_name=pivot_name)
        except (NotAllowedError, ValidationError) as exc:
            raise ToolError(f"Operation failed ({type(exc).__name__})") from exc
        except OfficeCOMError as exc:
            logger.error("[excelmcp] pivot_table read COM error: %s", exc, exc_info=True)
            raise ToolError(f"Pivot table read failed — {_classify_com_error(exc)}") from exc
        except Exception as exc:
            logger.error(
                "[excelmcp] pivot_table read unexpected error: %s", exc, exc_info=True
            )
            raise ToolError(f"Pivot table read failed ({type(exc).__name__})") from exc

    elif operation == "refresh":
        try:
            return _refresh_pivot_tables(path=path, sheet=sheet, confirm=confirm)
        except (NotAllowedError, ValidationError) as exc:
            raise ToolError(f"Operation failed ({type(exc).__name__})") from exc
        except OfficeCOMError as exc:
            logger.error("[excelmcp] pivot_table refresh COM error: %s", exc, exc_info=True)
            raise ToolError(f"Pivot table refresh failed — {_classify_com_error(exc)}") from exc
        except Exception as exc:
            logger.error(
                "[excelmcp] pivot_table refresh unexpected error: %s", exc, exc_info=True
            )
            raise ToolError(f"Pivot table refresh failed ({type(exc).__name__})") from exc

    else:
        raise ToolError(
            f"pivot_table operation must be 'create', 'read', or 'refresh'; got {operation!r}"
        )


# ---------------------------------------------------------------------------
# Backward-compat shims -- NOT MCP-registered
# ---------------------------------------------------------------------------

def open_template_and_recalculate(
    template_path: str,
    output_path: str,
    confirm: bool = False,
) -> dict:
    """Backward compat. Use hydrate_template(cell_writes=None) in new code."""
    result = hydrate_template(
        template_path=template_path,
        cell_writes=[],
        output_path=output_path,
        recalculate=True,
        confirm=confirm,
    )
    result.setdefault("template", template_path)
    return result


def evaluate_formula_live(
    path: str,
    sheet: str,
    cell_address: str,
) -> dict:
    """Backward compat. Use evaluate_formula(live=True) in new code."""
    from excelmcp._core import NotAllowedError, ValidationError, OfficeCOMError  # noqa: PLC0415

    try:
        return _evaluate_formula_live(path=path, sheet=sheet, cell_address=cell_address)
    except (NotAllowedError, ValidationError, OfficeCOMError) as exc:
        logger.error("[excelmcp] evaluate_formula_live error: %s", exc, exc_info=True)
        raise ToolError(f"Operation failed ({type(exc).__name__})") from exc
    except Exception as exc:
        logger.error("[excelmcp] evaluate_formula_live error: %s", exc, exc_info=True)
        raise ToolError(f"Formula evaluation failed ({type(exc).__name__})") from exc


def create_pivot_table(
    path: str,
    source_range: str,
    dest_sheet: str,
    dest_cell: str,
    row_fields: list,
    col_fields: list,
    value_fields: list,
    table_name: str = "PivotTable1",
    confirm: bool = False,
) -> dict:
    """Backward compat. Use pivot_table(operation='create') in new code."""
    try:
        return _create_pivot_table(
            path=path,
            source_range=source_range,
            dest_sheet=dest_sheet,
            dest_cell=dest_cell,
            row_fields=row_fields,
            col_fields=col_fields,
            value_fields=value_fields,
            table_name=table_name,
            confirm=confirm,
        )
    except (NotAllowedError, ValidationError, OfficeCOMError) as exc:
        logger.error("[excelmcp] create_pivot_table error: %s", exc, exc_info=True)
        raise ToolError(f"Operation failed ({type(exc).__name__})") from exc
    except Exception as exc:
        logger.error("[excelmcp] create_pivot_table error: %s", exc, exc_info=True)
        raise ToolError(f"Pivot table creation failed ({type(exc).__name__})") from exc


def refresh_pivot_tables(
    path: str,
    sheet: str | None = None,
    confirm: bool = False,
) -> dict:
    """Backward compat. Use pivot_table(operation='refresh') in new code."""
    try:
        return _refresh_pivot_tables(path=path, sheet=sheet, confirm=confirm)
    except (NotAllowedError, ValidationError, OfficeCOMError) as exc:
        logger.error("[excelmcp] refresh_pivot_tables error: %s", exc, exc_info=True)
        raise ToolError(f"Operation failed ({type(exc).__name__})") from exc
    except Exception as exc:
        logger.error("[excelmcp] refresh_pivot_tables error: %s", exc, exc_info=True)
        raise ToolError(f"Pivot table refresh failed ({type(exc).__name__})") from exc


def read_pivot_table(
    path: str,
    sheet: str,
    pivot_name: str,
) -> dict:
    """Backward compat. Use pivot_table(operation='read') in new code."""
    try:
        return _read_pivot_table(path=path, sheet=sheet, pivot_name=pivot_name)
    except (NotAllowedError, ValidationError, OfficeCOMError) as exc:
        logger.error("[excelmcp] read_pivot_table error: %s", exc, exc_info=True)
        raise ToolError(f"Operation failed ({type(exc).__name__})") from exc
    except Exception as exc:
        logger.error("[excelmcp] read_pivot_table error: %s", exc, exc_info=True)
        raise ToolError(f"Pivot table read failed ({type(exc).__name__})") from exc
