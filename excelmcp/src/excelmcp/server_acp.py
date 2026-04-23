"""FastMCP tool wiring for the ExcelMCP ACP adapter.

Side-effect import — registers get_workbook_context on the shared mcp instance.
Add to server.py:
    import excelmcp.server_acp  # noqa: F401, E402  — registers ACP tool
"""
from __future__ import annotations

import logging
from typing import Literal

from fastmcp.exceptions import ToolError

from excelmcp._acp_adapter import build_workbook_context
from excelmcp._core import ExcelMCPError
from excelmcp._server_instance import mcp

logger = logging.getLogger(__name__)


@mcp.tool()
def get_workbook_context(
    path: str,
    level: Literal["index", "focused", "deep"] = "focused",
    annotations: list[dict] | None = None,
) -> dict:
    """Return an Artifact Context Packet (ACP) for an Excel workbook.

    Progressive disclosure via the ``level`` parameter:

    - **index**   — identity + summary only (fast; single stat + openpyxl open)
    - **focused** — + sheet_names, sheet_count, table_names
    - **deep**    — + metadata_policy, named_ranges (names only), lineage_summary,
                    and optional annotations

    Security guarantees:
    - Path must be within EXCEL_ALLOWLIST_ROOTS (allowlist is enforced).
    - ACPDetail never contains cell values (S-001).
    - Annotations are validated with validate_acp_annotations() before storage.

    Returns a JSON-serialisable dict matching the ACP contract v1.0 shape.
    """
    try:
        return build_workbook_context(path, level=level, annotations=annotations)
    except ValueError as exc:
        raise ToolError(str(exc)) from exc
    except ExcelMCPError as exc:
        raise ToolError(str(exc)) from exc
