"""
FastMCP tool wiring for PowerPoint MCP.
Thin aggregator — delegates all tool registrations to handler modules.
"""
from __future__ import annotations

import logging
import sys

from fastmcp import FastMCP

try:
    from fastmcp.server.middleware.error_handling import ErrorHandlingMiddleware as _ErrorHandlingMiddleware
    _HAS_ERROR_MIDDLEWARE = True
except ImportError:
    _HAS_ERROR_MIDDLEWARE = False

from pptmcp import (
    _server_compact,
    _server_handlers_acp,
    _server_handlers_com,
    _server_handlers_export,
    _server_handlers_read,
    _server_handlers_review,
    _server_handlers_write,
)
from pptmcp import contract_pptx as contract
from pptmcp._server_handlers_comments import register_comments_tools as _register_comments_tools
from pptmcp._server_handlers_batch_text import register_batch_text_tools as _register_batch_text_tools
from pptmcp._server_handlers_table import register_table_tools as _register_table_tools

# Re-export COM symbols for backwards compatibility (test_security_guards_core patches these)
from pptmcp._server_handlers_com import (  # noqa: F401
    _COM_AVAILABLE,
    _pcom,
    export_changed_slides_only,
    export_deck,
    export_slide,
    export_slides_to_stamped_dir,
)

# Re-export dispatcher tools for backwards compatibility (test_dispatcher_tools imports these)
from pptmcp._server_handlers_write import add_content, set_format, set_table_style, slide, shape  # noqa: F401

logging.basicConfig(stream=sys.stderr, level=logging.INFO)
logger = logging.getLogger(__name__)

mcp = FastMCP("pptmcp")

if _HAS_ERROR_MIDDLEWARE:
    mcp.add_middleware(_ErrorHandlingMiddleware())

# Register either the full backward-compatible surface or the compact
# dispatcher-first surface intended for agent deployments.
if _server_compact.compact_surface_enabled():
    _server_compact.register_compact_tools(mcp)
else:
    _server_handlers_read.register_read_tools(mcp)
    _server_handlers_write.register_write_tools(mcp)
    _server_handlers_export.register_export_tools(mcp)
    _server_handlers_review.register_review_tools(mcp)
    _server_handlers_com.register_com_tools(mcp)
    _server_handlers_acp.register_acp_tools(mcp)

_register_comments_tools(mcp)
_register_batch_text_tools(mcp)
_register_table_tools(mcp)

# ---------------------------------------------------------------------------
# MCP Resources
# ---------------------------------------------------------------------------

@mcp.resource(
    "resource://contract-schema",
    name="contract-schema",
    description="JSON Schema v1.0 for pptmcp Output Contract files",
    mime_type="application/json",
)
def contract_schema() -> str:
    """Return the JSON Schema for a pptmcp Output Contract v1.0 file."""
    return contract.get_contract_schema()


@mcp.resource("ppt://prompts/ppt_render_check_iterate_v1")
def ppt_render_check_iterate_resource() -> str:
    """Recommended workflow prompt: Spec -> Build -> Render -> Review -> Iterate."""
    return (
        "# PPT Render-Check-Iterate Workflow v1\n\n"
        "## Step 1 - Declare Contract\n"
        "call: validate_contract(contract_path)\n\n"
        "## Step 2 - Build with Quality Gates\n"
        "call: add_shape(...) - check overlap_warning and overflow_risk in response\n\n"
        "## Step 3 - Export PNG\n"
        "call: export_slides_to_stamped_dir(path, base_dir, confirm=True)\n\n"
        "## Step 4 - Review\n"
        "call: review_slide_export(path, slide_index, png_path, contract_path)\n"
        "  -> if passed: proceed to Step 6\n"
        "  -> if failed: proceed to Step 5\n\n"
        "## Step 5 - Iterate (fix + re-export changed slides only)\n"
        "call: export_changed_slides_only(path, base_dir, previous_hashes, confirm=True)\n"
        "  -> repeat Step 4 for changed slides\n\n"
        "## Step 6 - Produce Evidence Bundle\n"
        "call: produce_evidence_bundle(path, review_results, contract_path, exports_dir)\n"
    )


# ---------------------------------------------------------------------------
# MCP Prompts
# ---------------------------------------------------------------------------
from pptmcp.server_prompts import (  # noqa: E402, F401
    ppt_render_check_iterate_v1,
    register_prompts,
)

register_prompts(mcp)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
