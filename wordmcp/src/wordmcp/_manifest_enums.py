"""Enumeration constants for wordmcp tool manifests."""
from __future__ import annotations

COM_TOOLS = [
    "manage_tracked_changes",
    "export_document",
]

CONTEXT_TOOLS = [
    "get_document_context",
]

EXTRA_TOOLS = [
    "manage_comments",
    "export_document",
    "create_document",
    "get_document_context",
    "list_headings",
    "read_section",
    "search_text",
    "list_styles",
    "apply_style",
    "update_paragraph",
    "delete_paragraph",
    "update_table_cell",
    "set_document_properties",
    "insert_paragraph",
    "add_list",
    "bulk_add_paragraphs",
    "bulk_update_paragraphs",
    "bulk_update_table_cells",
    "get_headers_footers",
    "add_hyperlink",
    "add_footnote",
    # Review / evidence tools
    "get_document_outline",
    "review_document",
    "write_review_findings",
    "export_review_evidence",
    # Paragraph dispatcher + set_format backend
    "paragraph",
    "set_paragraph_format",
    # Export scope dispatcher
    "export",
    # Table dispatcher
    "table",
    # Style dispatcher
    "style",
    # Content dispatcher
    "content",
    # Review dispatcher
    "review",
    # Document dispatcher
    "document",
    # Assembly tools
    "bulk_find_replace",
    "manage_hyperlinks",
    "verify_no_placeholders",
]

READ_DOCUMENT_SCOPES = [
    "full",
    "metadata",
    "headings",
    "outline",
    "section",
    "paragraphs",
    "context",
]

EXPORT_SCOPES = [
    "pdf",
    "txt",
    "markdown",
]

TABLE_OPERATIONS = [
    "list",
    "read",
    "add",
    "update_cell",
    "bulk_update_cells",
]

STYLE_OPERATIONS = [
    "list",
    "apply",
]

CONTENT_OPERATIONS = [
    "add_list",
    "add_page_break",
    "insert_image",
    "add_hyperlink",
    "add_footnote",
    "find_replace",
    "set_properties",
]

REVIEW_OPERATIONS = [
    "review",
    "write_findings",
    "export_evidence",
    "manage_comments",
    "manage_tracked_changes",
]

DOCUMENT_OPERATIONS = [
    "save",
    "search_text",
    "get_headers_footers",
]

PARAGRAPH_OPERATIONS = [
    "list",
    "read",
    "add",
    "add_heading",
    "insert",
    "update",
    "delete",
    "set_format",
    "bulk_add",
    "bulk_update",
]
