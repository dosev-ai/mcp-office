"""
FastMCP tool wiring for Excel MCP Phase 2.0.
No openpyxl imports ��� delegates all I/O to workbook_openpyxl.
"""
from __future__ import annotations

import json
import logging
import os
import sys
from urllib.parse import unquote

# ---------------------------------------------------------------------------
# Namespace-package shadow fix — must run before any excelmcp.* imports.
# When `python -m excelmcp.server` is run from the repo root, Python prepends
# '' (CWD) to sys.path and discovers excelmcp/ as a namespace package,
# shadowing the editable-install package at src/excelmcp/.  This causes
# `from excelmcp import __version__` (and relative equivalents) to fail with
# ImportError because the namespace package has no __init__.py.
# ---------------------------------------------------------------------------
def _fix_namespace_shadow() -> None:
    import importlib as _il
    pkg = sys.modules.get("excelmcp")
    if pkg is not None and getattr(pkg, "__file__", None) is None:
        # Namespace package detected — find the real src/ dir and reimport.
        _src = os.path.normpath(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir)
        )
        del sys.modules["excelmcp"]
        if _src not in sys.path:
            sys.path.insert(0, _src)
        _il.import_module("excelmcp")

_fix_namespace_shadow()
del _fix_namespace_shadow
# ---------------------------------------------------------------------------

# Suppress FastMCP startup banner, upgrade nag, and docket.worker noise
# before FastMCP is imported so the setting takes effect at construction time.
os.environ.setdefault("FASTMCP_SHOW_CLI_BANNER", "false")

from fastmcp.exceptions import ToolError

from excelmcp._server_instance import mcp
from excelmcp.workbook_openpyxl import (
    ExcelMCPError,
    resolve_charts_resource,
    resolve_named_ranges_resource,
    resolve_sheet_resource,
    resolve_sheets_resource,
    resolve_tables_resource,
    resolve_validations_resource,
)

logging.basicConfig(stream=sys.stderr, level=logging.INFO)
logging.getLogger("docket").setLevel(logging.WARNING)
logging.getLogger("docket.worker").setLevel(logging.WARNING)
logging.getLogger("mcp.server.lowlevel.server").setLevel(logging.WARNING)
logger = logging.getLogger(__name__)

import excelmcp.server_acp  # noqa: F401, E402  -- registers ACP tool
import excelmcp.server_batch  # noqa: F401, E402 � side-effect: registers batch tools
import excelmcp.server_chart  # noqa: F401, E402 � side-effect: registers chart tools
import excelmcp.server_com  # noqa: F401, E402 � side-effect: registers COM tools
import excelmcp.server_format  # noqa: F401, E402 � side-effect: registers format tools
import excelmcp.server_io  # noqa: F401, E402 � side-effect: registers I/O tools
import excelmcp.server_metadata_prompts  # noqa: F401, E402  – registers metadata prompts & contract resource
import excelmcp.server_ops  # noqa: F401, E402 � side-effect: registers ops tools
import excelmcp.server_prompts_analysis  # noqa: F401, E402 -- registers analysis prompts
import excelmcp.server_prompts_format  # noqa: F401, E402 -- registers format prompts
import excelmcp.server_prompts_mgmt  # noqa: F401, E402 -- registers mgmt prompts
import excelmcp.server_prompts_workflow  # noqa: F401, E402 -- registers workflow prompts
import excelmcp.server_review  # noqa: F401, E402  – registers review tools
import excelmcp.server_snapshot  # noqa: F401, E402  -- side-effect: registers snapshot/diff tools
from excelmcp.server_format import apply_style  # noqa: E402, F401

# ---------------------------------------------------------------------------
# Backward-compatible re-exports (for tests that import symbols from server)
# ---------------------------------------------------------------------------
from excelmcp.server_io import (  # noqa: E402, F401
    append_rows,
    cell,
    evaluate_formula,
    named_range,
    range_io,
    sheet,
)
from excelmcp.server_ops import (  # noqa: E402, F401
    cell_comment,
    cols,
    data_validation,
    hyperlink,
    merge,
    protect,
    rows,
    sheet_properties,
)

# ---------------------------------------------------------------------------
# MCP Resources (Phase 2.0) ��� read-only structural views
# ---------------------------------------------------------------------------

@mcp.resource("excelmcp://workbook/{path}/sheets")
def resource_sheets(path: str) -> str:
    """List all worksheets in the workbook with used-range dimensions."""
    decoded_path = unquote(path)
    try:
        return json.dumps(resolve_sheets_resource(decoded_path))
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.resource("excelmcp://workbook/{path}/sheet/{name}")
def resource_sheet(path: str, name: str) -> str:
    """Return structural info (dims, headers, tables) for a specific sheet."""
    decoded_path = unquote(path)
    decoded_name = unquote(name)
    try:
        return json.dumps(resolve_sheet_resource(decoded_path, decoded_name))
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.resource("excelmcp://workbook/{path}/named_ranges")
def resource_named_ranges(path: str) -> str:
    """Return all named ranges defined in the workbook."""
    decoded_path = unquote(path)
    try:
        return json.dumps(resolve_named_ranges_resource(decoded_path))
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.resource("excelmcp://workbook/{path}/charts")
def resource_charts(path: str) -> str:
    """Chart inventory for a workbook � returns all embedded charts with type, anchor, and data address."""
    decoded_path = unquote(path)
    try:
        return json.dumps(resolve_charts_resource(decoded_path))
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.resource("excelmcp://workbook/{path}/tables")
def resource_tables(path: str) -> str:
    """Table inventory for a workbook � returns all Excel Tables with column names and references."""
    decoded_path = unquote(path)
    try:
        return json.dumps(resolve_tables_resource(decoded_path))
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


@mcp.resource("excelmcp://workbook/{path}/validations")
def resource_validations(path: str) -> str:
    """Data validation inventory � returns all validation rules with address range, type, and formula."""
    decoded_path = unquote(path)
    try:
        return json.dumps(resolve_validations_resource(decoded_path))
    except ExcelMCPError as e:
        raise ToolError(str(e)) from e


def main() -> None:
    logger.info("excelmcp starting (stdio transport)")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
