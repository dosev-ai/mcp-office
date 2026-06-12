"""ACP tool registration for pptmcp.

Called from server.py:
    from pptmcp import _server_handlers_acp
    ...
    _server_handlers_acp.register_acp_tools(mcp)
"""
from __future__ import annotations

import logging
from typing import Literal

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

from pptmcp._pptx_acp_adapter import build_presentation_context
from pptmcp.presentation_pptx import PPTMCPError

logger = logging.getLogger(__name__)


def register_acp_tools(mcp: FastMCP) -> None:
    """Register get_presentation_context on the given FastMCP instance."""

    @mcp.tool()
    def get_presentation_context(
        path: str,
        level: Literal["index", "focused", "deep"] = "focused",
        annotations: list[dict] | None = None,
    ) -> dict:
        """Return an Artifact Context Packet (ACP) for a PowerPoint presentation.

        Progressive disclosure via the ``level`` parameter:

        - **index**   — identity + summary only (fast)
        - **focused** — + slide_count, slide_titles (one per slide, '' if no title)
        - **deep**    — + png_paths (if a rendered-export directory exists),
                        review_findings (empty; populate from produce_evidence_bundle),
                        and optional annotations

        Security guarantees:
        - Path must be within PPT_ALLOWLIST_ROOTS (allowlist is enforced).
        - Annotations are validated with validate_acp_annotations() before storage.

        Returns a JSON-serialisable dict matching the ACP contract v1.0 shape.
        """
        try:
            return build_presentation_context(
                path, level=level, annotations=annotations
            )
        except ValueError as exc:
            raise ToolError(str(exc)) from exc
        except PPTMCPError as exc:
            raise ToolError(str(exc)) from exc
