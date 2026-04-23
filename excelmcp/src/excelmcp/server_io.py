"""Core I/O MCP tools for excelmcp.

Extracted from server.py. Tools cover workbook creation, sheet management,
cell/range read-write, named ranges, tables, formula evaluation, CSV import,
and save.

Import side-effect: importing this module registers all tools on the shared
mcp instance from excelmcp._server_instance.
"""
from __future__ import annotations

import asyncio
import json
import logging
import threading
from typing import Any, Literal
from urllib.parse import unquote

from fastmcp.exceptions import ToolError

from excelmcp import workbook_openpyxl as wb
from excelmcp._com import _evaluate_formula_live
from excelmcp._core import NotAllowedError, ValidationError
from excelmcp._metadata_contract import build_workbook_metadata
from excelmcp._server_instance import mcp
from excelmcp.runtime_config import get_effective_policy
from excelmcp.workbook_openpyxl import ExcelMCPError

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Discovery / metadata
# ---------------------------------------------------------------------------


def _run_async(awaitable_factory):
    """Run a small async FastMCP API call from synchronous tool handlers."""
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return asyncio.run(awaitable_factory())

    result: Any | None = None
    error: BaseException | None = None

    def _runner() -> None:
        nonlocal result, error
        try:
            result = asyncio.run(awaitable_factory())
        except BaseException as exc:  # pragma: no cover - defensive relay
            error = exc

    thread = threading.Thread(target=_runner, daemon=True)
    thread.start()
    thread.join()
    if error is not None:
        raise error
    return result


def _get_live_prompt_names(mcp_server: Any = mcp) -> list[str]:
    """Return prompt names from the shared live MCP registry."""

    async def _load() -> list[str]:
        prompts = await mcp_server.list_prompts()
        return [prompt.name for prompt in prompts]

    return _run_async(_load)

def get_workbook_metadata(path: str) -> dict:
    """Build structural WorkbookMetadata for *path* and inject the active policy.

    This is the shared implementation for both the /meta and /metadata resources.
    Raises ToolError (delegated from build_workbook_metadata) on any field failure.
    """
    meta = dict(build_workbook_metadata(path))          # WorkbookMetadata -> plain dict
    meta["policy"] = get_effective_policy()             # inject runtime policy
    return meta


@mcp.tool()
def capabilities() -> dict:
    """Return metadata about this MCP phase and its available tools."""
    try:
        prompt_names = _get_live_prompt_names(mcp)
        return wb.capabilities(prompt_names=prompt_names)
    except ToolError:
        raise
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e
    except Exception as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# Structural write tools  (require EXCEL_ENABLE_WRITE=true + confirm=True)
# ---------------------------------------------------------------------------

@mcp.tool()
def create_workbook(
    path: str,
    sheets: list[str] | None = None,
    overwrite: bool = False,
    confirm: bool = False,
    session_mode: bool = False,
) -> dict:
    """Create a new .xlsx workbook at the specified path.

    sheets: worksheet names to create (default: ["Sheet1"]).
    overwrite: if False (default), raises ExcelMCPError if the file exists.
    confirm: must be True (write gate — prevents accidental creation).
    Returns: {"ok": true, "path": "...", "sheets": [...], "created": true}
    session_mode: Reserved for future active-workbook COM routing (Child 4).
        Pass False (default) to use the standard openpyxl path.
        Raises ToolError if True — routing not yet implemented.
    """
    if session_mode:
        raise ToolError(
            "session_mode=True routing is not yet implemented. "
            "Omit session_mode or set session_mode=False to use openpyxl path."
        )
    try:
        return wb.create_workbook(
            path, sheets=sheets, overwrite=overwrite, confirm=confirm
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# Sheet navigation
# ---------------------------------------------------------------------------

@mcp.tool()
def sheet(
    operation: Literal["list", "add", "rename", "delete"],
    path: str,
    sheet_name: str | None = None,
    new_name: str | None = None,
    position: int | None = None,
    confirm: bool = False,
) -> dict:
    """Manage workbook sheets: list, add, rename, or delete.

    operation='list': returns {"sheets": [str, ...], "count": int} — no confirm needed.
    operation='add': adds a new sheet named sheet_name at optional position.
    operation='rename': renames sheet_name to new_name.
    operation='delete': deletes sheet named sheet_name.
    Write operations require EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "list":
        try:
            result = wb.list_sheets(path)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
        if isinstance(result, list):
            return {"sheets": result, "count": len(result)}
        return result
    elif operation == "add":
        if sheet_name is None:
            raise ToolError("sheet_name is required for operation='add'")
        try:
            return wb.add_sheet(path, sheet_name, position=position, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "rename":
        if sheet_name is None:
            raise ToolError("sheet_name is required for operation='rename'")
        if new_name is None:
            raise ToolError("new_name is required for operation='rename'")
        try:
            return wb.rename_sheet(path, sheet_name, new_name, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "delete":
        if sheet_name is None:
            raise ToolError("sheet_name is required for operation='delete'")
        try:
            return wb.delete_sheet(path, sheet_name, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'list', 'add', 'rename', or 'delete'.")


@mcp.tool()
def get_used_range(path: str, sheet: str) -> dict:
    """Return the bounding box of used cells in *sheet* of the workbook at *path*."""
    try:
        return wb.get_used_range(path, sheet)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# Read tools
# ---------------------------------------------------------------------------

@mcp.tool()
def cell(
    operation: Literal["read", "write"],
    path: str,
    sheet: str,
    address: str,
    value: object = None,
    confirm: bool = False,
    session_mode: bool = False,
) -> dict:
    """Read or write a single cell value.

    operation='read': returns cell value, type, and formula (read-only).
    operation='write': writes value to the cell. Requires EXCEL_ENABLE_WRITE=true and confirm=True.
    value=None is not allowed for 'write' (use range_io with values=[[None]] to clear a cell).
    session_mode: when True, writes via COM to the active Excel workbook (requires
        EXCEL_ENABLE_COM=true). The workbook at `path` must be the currently active
        workbook in Excel. Guards: xl_app.Ready, ReadOnly, path match.
        Omit or set False to use the openpyxl file path (default).
    """
    if operation == "read":
        try:
            return wb.read_cell(path, sheet, address)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "write":
        if value is None:
            raise ToolError("value is required for operation='write' (use range_io with values=[[None]] to clear a cell)")
        try:
            return wb.write_cell(path, sheet, address, value, confirm=confirm, session_mode=session_mode)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'read' or 'write'.")


@mcp.tool()
def range_io(
    operation: Literal["read", "write"],
    path: str,
    sheet: str,
    address: str,
    values: list[list] | None = None,
    confirm: bool = False,
    session_mode: bool = False,
) -> dict:
    """Read or write a rectangular range of cells.

    operation='read': returns {"data": [[{...},...], ...], "rows": int, "cols": int}.
    operation='write': writes values 2-D array starting from address cell.
    Requires EXCEL_ENABLE_WRITE=true and confirm=True for 'write'.
    session_mode: when True, writes via COM to the active Excel workbook (requires
        EXCEL_ENABLE_COM=true). The workbook at `path` must be the currently active
        workbook in Excel. Guards: xl_app.Ready, ReadOnly, path match.
        Omit or set False to use the openpyxl file path (default).
    """
    if operation == "read":
        try:
            raw = wb.read_range(path, sheet, address)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
        rows_count = len(raw)
        cols_count = len(raw[0]) if rows_count > 0 else 0
        return {"data": raw, "rows": rows_count, "cols": cols_count}
    elif operation == "write":
        if values is None:
            raise ToolError("values is required for operation='write'")
        try:
            return wb.write_range(path, sheet, address, values, confirm=confirm, session_mode=session_mode)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'read' or 'write'.")


@mcp.tool()
def read_sheet_all(path: str, sheet: str) -> dict:
    """Read the entire used range of *sheet* from the workbook at *path*.

    Returns all cell data, row count, and column count.
    Subject to EXCEL_MAX_RANGE_CELLS guard.
    """
    try:
        return wb.read_sheet_all(path, sheet)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def named_range(
    operation: Literal["list", "read", "write"],
    path: str,
    name: str | None = None,
    sheet: str | None = None,
    address: str | None = None,
    confirm: bool = False,
) -> dict:
    """List, read, or write Excel named ranges (defined names).

    operation='list': returns {"named_ranges": [{name, refers_to, scope}, ...], "count": N}.
    operation='read': reads the data values at the named range. name is required.
    operation='write': creates or updates a named range. name, sheet, address, confirm all required.
    'write' requires EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "list":
        try:
            result = wb.get_named_ranges(path)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
        if isinstance(result, list):
            return {"named_ranges": result, "count": len(result)}
        return result
    elif operation == "read":
        if name is None:
            raise ToolError("name is required for operation='read'")
        try:
            return wb.read_named_range(path, name)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    elif operation == "write":
        if name is None:
            raise ToolError("name is required for operation='write'")
        if sheet is None:
            raise ToolError("sheet is required for operation='write'")
        if address is None:
            raise ToolError("address is required for operation='write'")
        try:
            return wb.write_named_range(path, name=name, sheet=sheet, range_address=address, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    else:
        raise ToolError(f"Unknown operation: {operation}. Use 'list', 'read', or 'write'.")


@mcp.tool()
def get_cell_metadata(path: str, sheet: str, address: str) -> dict:
    """Return formatting and metadata for a single cell.

    Includes font, fill, number format, comment, and hyperlink information.
    *address* must be in A1 notation (e.g. "B3").
    """
    try:
        return wb.get_cell_metadata(path, sheet, address)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def evaluate_formula(path: str, sheet: str, address: str, live: bool = False) -> dict:
    """Return the formula string and value for *address*.

    When live=False (default): openpyxl returns the last-saved cached value.
    When live=True: opens workbook via Excel COM for live evaluation.
    Useful for volatile functions (TODAY, NOW, RAND, OFFSET).
    Requires EXCEL_ENABLE_COM=true when live=True.
    """
    if live:
        try:
            return _evaluate_formula_live(path=path, sheet=sheet, cell_address=address)
        except Exception as e:
            raise ToolError(str(e)) from e
    try:
        return wb.evaluate_formula(path, sheet, address)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def list_tables(path: str, sheet: str | None = None) -> list[dict]:
    """List all Excel tables (ListObjects) in *sheet*, or across all sheets when
    *sheet* is omitted.

    Each entry contains sheet, name, display_name, ref, header_row, and totals_row.
    Omitting *sheet* enables workbook-wide table discovery.
    """
    try:
        return wb.list_tables(path, sheet)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def read_table(path: str, sheet: str, table_name: str) -> dict:
    """Read all rows from a named Excel table in *sheet*.

    Returns table metadata, headers, and all data rows.
    """
    try:
        return wb.read_table(path, sheet, table_name)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------------------
# Write tools  (require EXCEL_ENABLE_WRITE=true + confirm=True)
# ---------------------------------------------------------------------------

@mcp.tool()
def append_rows(
    path: str,
    sheet: str,
    rows: list[list[Any]],
    confirm: bool = False,
    session_mode: bool = False,
) -> dict:
    """Append *rows* after the last used row in *sheet*.

    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    session_mode: Reserved for future active-workbook COM routing (Child 4).
        Pass False (default) to use the standard openpyxl path.
        Raises ToolError if True — routing not yet implemented.
    """
    if session_mode:
        raise ToolError(
            "session_mode=True routing is not yet implemented. "
            "Omit session_mode or set session_mode=False to use openpyxl path."
        )
    try:
        return wb.append_rows(path, sheet, rows, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def find_replace(
    path: str,
    sheet: str,
    find_value: str,
    replace_value: str,
    match_case: bool = False,
    whole_cell: bool = True,
    confirm: bool = False,
) -> dict:
    """Find and replace string values across all cells in *sheet*.

    Set ``whole_cell=False`` for substring replacement.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.find_replace(
            path, sheet, find_value, replace_value,
            match_case=match_case, whole_cell=whole_cell, confirm=confirm
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def import_csv_to_sheet(
    path: str,
    csv_path: str,
    sheet: str,
    skip_header: bool = False,
    overwrite: bool = False,
    confirm: bool = False,
) -> dict:
    """Import a CSV or TSV file into *sheet*.

    All rows including the header row are imported by default so column names
    are preserved.  Set ``skip_header=True`` to discard the first row.
    Set ``overwrite=True`` to replace an existing sheet with the same name.
    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    """
    try:
        return wb.import_csv_to_sheet(
            path, csv_path, sheet,
            skip_header=skip_header, overwrite=overwrite, confirm=confirm
        )
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def save(path: str, confirm: bool = False) -> dict:
    """Flush the in-memory workbook at *path* to disk.

    Requires ``EXCEL_ENABLE_WRITE=true`` and ``confirm=True``.
    Evicts the workbook from the cache after saving.
    """
    try:
        return wb.save(path, confirm=confirm)
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.tool()
def workbook_metadata(
    operation: Literal["read", "write"],
    path: str,
    payload: dict[str, Any] | None = None,
    confirm: bool = False,
) -> dict:
    """Read or write workbook-level metadata stored on a hidden worksheet.

    operation='read': returns whether metadata is present plus stored payload details.
    operation='write': stores payload on hidden sheet _MCP_META. Requires
    EXCEL_ENABLE_WRITE=true and confirm=True.
    """
    if operation == "read":
        try:
            return wb.read_workbook_metadata(path)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    if operation == "write":
        if payload is None:
            raise ToolError("payload is required for operation='write'")
        try:
            return wb.write_workbook_metadata(path, payload, confirm=confirm)
        except ExcelMCPError as e:
            raise ToolError(str(e)) from e
    raise ToolError("Unknown operation: use 'read' or 'write'.")


@mcp.resource("excelmcp://workbook/{path}/metadata")
def resource_workbook_metadata(path: str) -> str:
    """Return canonical structural WorkbookMetadata for the workbook at *path*.

    Fields: path, workbook_name, sheet_names, active_sheet, sheet_count,
    has_formulas, last_modified, size_bytes, format, contract_version, policy.
    """
    decoded_path = unquote(path)
    try:
        return json.dumps(get_workbook_metadata(decoded_path))
    except ToolError:
        raise
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.resource("excelmcp://workbook/{path}/meta")
def resource_workbook_meta_legacy(path: str) -> str:
    """Legacy alias for excelmcp://workbook/{path}/metadata — same canonical shape.

    Preserved for backward compatibility. Do not remove.
    Both /meta and /metadata call get_workbook_metadata() — no divergence.
    """
    decoded_path = unquote(path)
    try:
        return json.dumps(get_workbook_metadata(decoded_path))
    except ToolError:
        raise
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e



@mcp.tool()
def consolidate_ranges(
    ranges: list[dict],
    source_label_key: str = "source",
    max_rows: int = 500,
) -> dict[str, Any]:
    """Read multiple ranges from workbooks and merge into one table.

    Each entry in ``ranges`` must have keys: ``path``, ``sheet``, ``range``,
    and optionally ``label``. Enforces EXCEL_ALLOWLIST_ROOTS via wb.read_range
    (BLOCKER-2 compliance -- no direct openpyxl.load_workbook calls here).

    Does NOT require EXCEL_ENABLE_WRITE. Does NOT require confirm.
    """
    try:
        if len(ranges) == 0:
            raise ValidationError("consolidate_ranges: ranges must not be empty")
        if len(ranges) > 20:
            raise ValidationError(
                f"consolidate_ranges: too many source ranges ({len(ranges)}); maximum is 20"
            )
        _RESERVED_ROW_KEYS = frozenset({"row_index", "data"})
        if source_label_key in _RESERVED_ROW_KEYS:
            raise ValidationError(
                f"source_label_key {source_label_key!r} conflicts with reserved output keys "
                f"{sorted(_RESERVED_ROW_KEYS)}; choose a different key name"
            )
        for i, t in enumerate(ranges):
            for required_key in ("path", "sheet", "range"):
                if required_key not in t:
                    raise ValidationError(
                        f"ranges[{i}] missing required key: {required_key!r}"
                    )
        all_rows: list[dict] = []
        hard_cap = min(max_rows, 500)
        truncated = False
        for t in ranges:
            path_str = t["path"]
            sheet_name = t["sheet"]
            address = t["range"]
            label = t.get("label", f"{path_str}:{sheet_name}!{address}")
            raw: list[list[dict]] = wb.read_range(
                path=path_str,
                sheet=sheet_name,
                a1_range=address,
            )
            for row_data in raw:
                if len(all_rows) >= hard_cap:
                    truncated = True
                    break
                all_rows.append({
                    source_label_key: label,
                    "row_index": len(all_rows),
                    "data": row_data,
                })
            if truncated:
                break
        return {
            "ok": True,
            "source_count": len(ranges),
            "total_rows": len(all_rows),
            "truncated": truncated,
            "rows": all_rows,
        }
    except (ValidationError, NotAllowedError):
        raise
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e
    except Exception as e:
        logger.error("consolidate_ranges failed: %s", e, exc_info=True)
        raise ToolError("consolidate_ranges failed -- check server logs") from e
