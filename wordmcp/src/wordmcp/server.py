"""
FastMCP tool wiring for Word MCP Phase 1.
No python-docx imports — delegates all I/O to document_docx.

STDOUT DISCIPLINE (CRITICAL — MCP PROTOCOL INTEGRITY):
  - NO print() calls anywhere in this module.
  - logging is configured to sys.stderr BEFORE FastMCP initialisation.
  - All diagnostic output via logger = logging.getLogger(__name__)
  - Violation corrupts the JSON-RPC framing on stdout and breaks MCP clients.
"""
from __future__ import annotations

import logging
import sys
from collections.abc import Callable
from typing import Any

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

from wordmcp import (
    _server_handlers_com,
    _server_handlers_content,
    _server_handlers_context,
    _server_handlers_docx_advanced,
    _server_handlers_docx_core,
    _server_handlers_docx_edit,
    _server_handlers_docx_export,
    _server_handlers_document,
    _server_handlers_export,
    _server_handlers_meta,
    _server_handlers_paragraph,
    _server_handlers_read_document,
    _server_handlers_review,
    _server_handlers_review_dispatcher,
    _server_handlers_style,
    _server_handlers_table,
)
from wordmcp import _server_handlers_assembly
from wordmcp import document_docx as doc
from wordmcp._server_runtime import ServerRuntime
from wordmcp.document_docx import WordMCPError

try:
    from wordmcp import document_com as _com
    _COM_LOADED = True
except ImportError:
    _com = None  # type: ignore[assignment]
    _COM_LOADED = False

# BLOCKER-4: stderr logging MUST be configured before FastMCP is instantiated.
logging.basicConfig(stream=sys.stderr, level=logging.INFO)
logger = logging.getLogger(__name__)

mcp = FastMCP("wordmcp")

# Declared for static analyzers; actual bindings are populated below via the
# shared registration loop to preserve the public wordmcp.server surface.
capabilities: Any
read_document: Any
get_document_metadata: Any
list_paragraphs: Any
read_paragraph: Any
list_tables: Any
read_table: Any
add_paragraph: Any
add_heading: Any
add_page_break: Any
add_table: Any
insert_image: Any
find_replace: Any
save: Any
manage_comments: Any
create_document: Any
list_headings: Any
read_section: Any
search_text: Any
list_styles: Any
manage_tracked_changes: Any
export_document: Any
apply_style: Any
update_paragraph: Any
delete_paragraph: Any
update_table_cell: Any
set_document_properties: Any
add_list: Any
insert_paragraph: Any
bulk_add_paragraphs: Any
bulk_update_paragraphs: Any
bulk_update_table_cells: Any
get_headers_footers: Any
add_hyperlink: Any
add_footnote: Any
get_document_outline: Any
review_document: Any
write_review_findings: Any
export_review_evidence: Any
get_document_context: Any
paragraph: Any
set_paragraph_format: Any
export: Any
table: Any
style: Any
content: Any
review: Any
document: Any
bulk_find_replace: Any
manage_hyperlinks: Any
verify_no_placeholders: Any


def _require_com() -> None:
    """Raise ToolError if the COM layer is not loaded."""
    if not _COM_LOADED:
        raise ToolError("COM tools require pywin32: pip install wordmcp[com]")


_runtime = ServerRuntime(
    get_doc=lambda: doc,
    get_word_error=lambda: WordMCPError,
    get_com=lambda: _com,
    get_com_loaded=lambda: _COM_LOADED,
    require_com=_require_com,
)


def _bind_tool(fn: Callable[..., Any], public_name: str):
    fn.__name__ = public_name
    fn.__qualname__ = public_name
    return mcp.tool()(fn)


for register in (
    _server_handlers_meta.register,
    _server_handlers_context.register,
    _server_handlers_read_document.register,
    _server_handlers_docx_core.register,
    _server_handlers_docx_edit.register,
    _server_handlers_docx_advanced.register,
    _server_handlers_com.register,
    _server_handlers_docx_export.register,
    _server_handlers_review.register,
    _server_handlers_paragraph.register,
    _server_handlers_export.register,
    _server_handlers_table.register,
    _server_handlers_style.register,
    _server_handlers_content.register,
    _server_handlers_review_dispatcher.register,
    _server_handlers_document.register,
    _server_handlers_assembly.register,
):
    globals().update(register(_bind_tool, _runtime))


@mcp.resource("word://prompts/word_review_check_iterate_v1")
def _word_review_prompt_resource() -> str:
    """Word document review checklist and iteration guide."""
    return _server_handlers_review.REVIEW_PROMPT_TEXT


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
