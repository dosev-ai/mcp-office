"""
read_document scope dispatcher for wordmcp.

Provides a unified `read_document` tool that routes to the appropriate
read handler based on the `scope` parameter.  All individual read tools
remain registered (with deprecation notes) — this dispatcher is an
additive convenience wrapper that replaces the previous scope-less
read_document registered in _server_handlers_docx_core.

Supported scopes
----------------
full        — Document summary (paragraph count, table count, etc.)
metadata    — Core document properties (title, author, dates, …)
headings    — Flat list of all headings with level and paragraph index
outline     — Hierarchical heading tree (via get_document_outline)
section     — Content of a specific section by heading_index
paragraphs  — Flat list of all paragraphs (same as list_paragraphs)
context     — Token-budget-friendly summary overview
"""
from __future__ import annotations

from collections.abc import Callable
from typing import Any

import wordmcp._docx.review_doc as _review_doc
from wordmcp import _server_context_tools, _server_docx_tools
from wordmcp._server_runtime import ServerRuntime

_VALID_SCOPES = (
    "full",
    "metadata",
    "headings",
    "outline",
    "section",
    "paragraphs",
    "context",
)


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _read_document(
        path: str,
        scope: str = "full",
        section_index: int | None = None,
    ) -> Any:
        """Read a Word document, routing to the appropriate handler via `scope`.

        path (required): absolute path to the .docx file.

        scope (optional, default 'full'): determines what is returned.

          full       — High-level summary: paragraph count, table count, image
                       count, section count, word count, metadata title/author.
          metadata   — Full document properties: title, author, subject,
                       keywords, created, modified, revision. Replaces the
                       deprecated get_document_metadata tool.
          headings   — Flat list of all headings with outline level, paragraph
                       index, and style name. Replaces list_headings.
          outline    — Hierarchical heading tree with paragraph count and table
                       count. Replaces get_document_outline.
          section    — Content of one section identified by its heading
                       paragraph index (section_index required). Replaces
                       read_section.
          paragraphs — Flat list of all body paragraphs with index, style, and
                       text_preview. Equivalent to list_paragraphs.
          context    — Token-budget-friendly overview: metadata + aggregate
                       counts + heading text only. Replaces get_document_context.

        section_index (int): required when scope='section'. Must be the
          paragraph index of a heading paragraph (use scope='headings' to
          discover valid indices).

        Unknown scope returns a structured error dict:
          {"error": "unknown_scope", "scope": s, "valid_scopes": [...]}
        """
        s = scope.strip().lower()

        if s == "full":
            return _server_docx_tools.read_document(runtime, path)

        if s == "metadata":
            return _server_docx_tools.get_document_metadata(runtime, path)

        if s == "headings":
            return _server_docx_tools.list_headings(runtime, path)

        if s == "outline":
            # get_document_outline uses the review_doc backend directly
            # (does its own allowlist check internally via _check_path)
            return _review_doc.get_document_outline(path)

        if s == "section":
            if section_index is None:
                from fastmcp.exceptions import ToolError
                raise ToolError("section_index is required when scope='section'")
            return _server_docx_tools.read_section(
                runtime, path=path, heading_index=section_index
            )

        if s == "paragraphs":
            return _server_docx_tools.list_paragraphs(runtime, path)

        if s == "context":
            return _server_context_tools.get_document_context(runtime, path)

        # Unknown scope — return structured error (no exception, caller can inspect)
        return {
            "error": "unknown_scope",
            "scope": scope,
            "valid_scopes": list(_VALID_SCOPES),
        }

    return {
        "read_document": bind_tool(_read_document, "read_document"),
    }
