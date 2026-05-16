"""Core I/O MCP tools for excelmcp.

Retained tools: capabilities (discovery) and workbook metadata resources.
All other tools have been extracted to focused modules:
  - server_range.py    — get_used_range, cell, range_io, read_sheet_all,
                         get_cell_metadata, consolidate_ranges
  - server_sheet.py    — sheet
  - server_workbook.py — create_workbook, save, workbook_metadata, evaluate_formula
  - server_table_io.py — list_tables, read_table, append_rows, find_replace,
                         import_csv_to_sheet, named_range

Import side-effect: importing this module registers the capabilities tool and
workbook metadata resources on the shared mcp instance from
excelmcp._server_instance.
"""
from __future__ import annotations

import asyncio
import json
import logging
import threading
from typing import Any
from urllib.parse import unquote

from fastmcp.exceptions import ToolError

from excelmcp import workbook_openpyxl as wb
from excelmcp._metadata_contract import build_workbook_metadata
from excelmcp._server_instance import mcp
from excelmcp.runtime_config import get_effective_policy
from excelmcp.workbook_openpyxl import ExcelMCPError

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Backward-compatible re-exports for tests that import directly from server_io
# ---------------------------------------------------------------------------
from excelmcp.server_range import (  # noqa: E402, F401
    cell,
    consolidate_ranges,
    get_cell_metadata,
    get_used_range,
    range_io,
    read_sheet_all,
)
from excelmcp.server_sheet import sheet  # noqa: E402, F401
from excelmcp.server_table_io import (  # noqa: E402, F401
    append_rows,
    find_replace,
    import_csv_to_sheet,
    list_tables,
    named_range,
    read_table,
)
from excelmcp.server_workbook import (  # noqa: E402, F401
    create_workbook,
    evaluate_formula,
    save,
    workbook_metadata,
)

# ---------------------------------------------------------------------------
# Discovery / metadata helpers
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


# ---------------------------------------------------------------------------
# Discovery tool
# ---------------------------------------------------------------------------

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
# Workbook metadata resources
# ---------------------------------------------------------------------------

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
