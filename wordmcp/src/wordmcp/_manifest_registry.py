"""Tool registry for wordmcp server manifest."""
from __future__ import annotations

from wordmcp._manifest_enums import (
    CONTENT_OPERATIONS,
    DOCUMENT_OPERATIONS,
    EXPORT_SCOPES,
    PARAGRAPH_OPERATIONS,
    READ_DOCUMENT_SCOPES,
    REVIEW_OPERATIONS,
    STYLE_OPERATIONS,
    TABLE_OPERATIONS,
)

# ---------------------------------------------------------------------------
# Deprecation policy
# ---------------------------------------------------------------------------

DEPRECATION_POLICY: dict = {
    "window_releases": 2,
    "telemetry_field": "deprecated_alias_called",
    "removal_date_iso": None,
}

# ---------------------------------------------------------------------------
# Per-tool registry (Option B — explicit, no docstring parsing)
#
# Each entry is a plain dict with these keys:
#   kind: "dispatcher" | "standalone" | "deprecated_alias"
#   replacement_tool: str | None          — only for deprecated_alias
#   replacement_operation_or_scope: str | None  — only for deprecated_alias
#   write_gated: bool
#   operations_or_scopes: list[str] | None     — only for dispatcher
# ---------------------------------------------------------------------------

_WRITE_GATE = {"env_var": "WORD_ENABLE_WRITE", "requires_confirm": True}

# TOOL_REGISTRY is a dict: tool_name -> metadata dict.
# Total: 10 primary (2 standalone + 8 dispatcher) + 37 deprecated aliases = 47 entries.
TOOL_REGISTRY: dict[str, dict] = {
    # ------------------------------------------------------------------
    # Standalone primary tools (2)
    # ------------------------------------------------------------------
    "capabilities": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "create_document": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    # ------------------------------------------------------------------
    # Dispatcher primary tools (8)
    # ------------------------------------------------------------------
    "read_document": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": sorted(READ_DOCUMENT_SCOPES),
    },
    "paragraph": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(PARAGRAPH_OPERATIONS),
    },
    "export": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(EXPORT_SCOPES),
    },
    "table": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(TABLE_OPERATIONS),
    },
    "style": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(STYLE_OPERATIONS),
    },
    "content": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(CONTENT_OPERATIONS),
    },
    "review": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(REVIEW_OPERATIONS),
    },
    "document": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(DOCUMENT_OPERATIONS),
    },
    # ------------------------------------------------------------------
    # Deprecated aliases — read-only (13)
    # ------------------------------------------------------------------
    "get_document_metadata": {
        "kind": "deprecated_alias",
        "replacement_tool": "read_document",
        "replacement_operation_or_scope": "metadata",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "get_document_outline": {
        "kind": "deprecated_alias",
        "replacement_tool": "read_document",
        "replacement_operation_or_scope": "outline",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "get_document_context": {
        "kind": "deprecated_alias",
        "replacement_tool": "read_document",
        "replacement_operation_or_scope": "context",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "read_section": {
        "kind": "deprecated_alias",
        "replacement_tool": "read_document",
        "replacement_operation_or_scope": "section",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "list_headings": {
        "kind": "deprecated_alias",
        "replacement_tool": "read_document",
        "replacement_operation_or_scope": "headings",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "list_paragraphs": {
        "kind": "deprecated_alias",
        "replacement_tool": "read_document",
        "replacement_operation_or_scope": "paragraphs",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "read_paragraph": {
        "kind": "deprecated_alias",
        "replacement_tool": "read_document",
        "replacement_operation_or_scope": "paragraphs",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "list_tables": {
        "kind": "deprecated_alias",
        "replacement_tool": "table",
        "replacement_operation_or_scope": "list",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "read_table": {
        "kind": "deprecated_alias",
        "replacement_tool": "table",
        "replacement_operation_or_scope": "read",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "list_styles": {
        "kind": "deprecated_alias",
        "replacement_tool": "style",
        "replacement_operation_or_scope": "list",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "search_text": {
        "kind": "deprecated_alias",
        "replacement_tool": "document",
        "replacement_operation_or_scope": "search_text",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "get_headers_footers": {
        "kind": "deprecated_alias",
        "replacement_tool": "document",
        "replacement_operation_or_scope": "get_headers_footers",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "review_document": {
        "kind": "deprecated_alias",
        "replacement_tool": "review",
        "replacement_operation_or_scope": "review",
        "write_gated": False,
        "operations_or_scopes": None,
    },
    # ------------------------------------------------------------------
    # Deprecated aliases — write-gated (24)
    # ------------------------------------------------------------------
    "add_paragraph": {
        "kind": "deprecated_alias",
        "replacement_tool": "paragraph",
        "replacement_operation_or_scope": "add",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "add_heading": {
        "kind": "deprecated_alias",
        "replacement_tool": "paragraph",
        "replacement_operation_or_scope": "add_heading",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "insert_paragraph": {
        "kind": "deprecated_alias",
        "replacement_tool": "paragraph",
        "replacement_operation_or_scope": "insert",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "update_paragraph": {
        "kind": "deprecated_alias",
        "replacement_tool": "paragraph",
        "replacement_operation_or_scope": "update",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "delete_paragraph": {
        "kind": "deprecated_alias",
        "replacement_tool": "paragraph",
        "replacement_operation_or_scope": "delete",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "bulk_add_paragraphs": {
        "kind": "deprecated_alias",
        "replacement_tool": "paragraph",
        "replacement_operation_or_scope": "bulk_add",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "bulk_update_paragraphs": {
        "kind": "deprecated_alias",
        "replacement_tool": "paragraph",
        "replacement_operation_or_scope": "bulk_update",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "set_paragraph_format": {
        "kind": "deprecated_alias",
        "replacement_tool": "paragraph",
        "replacement_operation_or_scope": "set_format",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "add_table": {
        "kind": "deprecated_alias",
        "replacement_tool": "table",
        "replacement_operation_or_scope": "add",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "update_table_cell": {
        "kind": "deprecated_alias",
        "replacement_tool": "table",
        "replacement_operation_or_scope": "update_cell",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "bulk_update_table_cells": {
        "kind": "deprecated_alias",
        "replacement_tool": "table",
        "replacement_operation_or_scope": "bulk_update_cells",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "apply_style": {
        "kind": "deprecated_alias",
        "replacement_tool": "style",
        "replacement_operation_or_scope": "apply",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "add_list": {
        "kind": "deprecated_alias",
        "replacement_tool": "content",
        "replacement_operation_or_scope": "add_list",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "add_page_break": {
        "kind": "deprecated_alias",
        "replacement_tool": "content",
        "replacement_operation_or_scope": "add_page_break",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "insert_image": {
        "kind": "deprecated_alias",
        "replacement_tool": "content",
        "replacement_operation_or_scope": "insert_image",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "add_hyperlink": {
        "kind": "deprecated_alias",
        "replacement_tool": "content",
        "replacement_operation_or_scope": "add_hyperlink",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "add_footnote": {
        "kind": "deprecated_alias",
        "replacement_tool": "content",
        "replacement_operation_or_scope": "add_footnote",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "find_replace": {
        "kind": "deprecated_alias",
        "replacement_tool": "content",
        "replacement_operation_or_scope": "find_replace",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "set_document_properties": {
        "kind": "deprecated_alias",
        "replacement_tool": "content",
        "replacement_operation_or_scope": "set_properties",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "save": {
        "kind": "deprecated_alias",
        "replacement_tool": "document",
        "replacement_operation_or_scope": "save",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "export_document": {
        "kind": "deprecated_alias",
        "replacement_tool": "export",
        "replacement_operation_or_scope": "pdf",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "write_review_findings": {
        "kind": "deprecated_alias",
        "replacement_tool": "review",
        "replacement_operation_or_scope": "write_findings",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "export_review_evidence": {
        "kind": "deprecated_alias",
        "replacement_tool": "review",
        "replacement_operation_or_scope": "export_evidence",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "manage_comments": {
        "kind": "deprecated_alias",
        "replacement_tool": "review",
        "replacement_operation_or_scope": "manage_comments",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    # ------------------------------------------------------------------
    # Assembly tools — standalone primaries (3)
    # ------------------------------------------------------------------
    "bulk_find_replace": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "manage_hyperlinks": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "verify_no_placeholders": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
}
