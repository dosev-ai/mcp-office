"""
Export scope dispatcher for wordmcp.

Provides a unified `export` tool that routes to the appropriate export backend
based on the `scope` parameter.  The existing `export_document` and
`export_review_evidence` tools remain registered (their docstrings have been
updated to note deprecation in favour of this dispatcher).

Supported scopes
----------------
pdf       — Export document to PDF via Word COM (COM required).
txt       — Extract plain text and write to output_path (python-docx).
markdown  — Render Markdown and write to output_path (python-docx).

All scopes require WORD_ENABLE_WRITE=true and confirm=True (mutations gate).
"""
from __future__ import annotations

import os
from collections.abc import Callable
from typing import Any

from fastmcp.exceptions import ToolError

from wordmcp._server_runtime import ServerRuntime

_VALID_SCOPES = ("pdf", "txt", "markdown")


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _export(
        path: str,
        scope: str,
        output_path: str,
        confirm: bool = False,
    ) -> Any:
        """Export a Word document using the specified scope.

        path (required): absolute path to the source .docx file.

        scope (required): determines the export format and backend.

          pdf      — Export to PDF via Word COM.  Requires pywin32 + Word installed.
                     output_path must end in .pdf.
          txt      — Extract all text and write to a UTF-8 .txt file (python-docx,
                     no COM required).  output_path must end in .txt.
          markdown — Render document as GitHub-Flavoured Markdown and write to a
                     UTF-8 .md file (python-docx, no COM required).
                     Headings → # / ## / …, bold/italic runs, GFM tables.
                     output_path must end in .md.

        output_path (required): absolute path where the exported file will be
          written.  The parent directory must already exist and must be within
          WORD_ALLOWLIST_ROOTS.

        confirm (required): must be True to perform the write.

        Mutation gate: WORD_ENABLE_WRITE=true AND confirm=True are both required
        before any file I/O occurs.

        Unknown scope returns a structured error dict (not an exception):
          {"error": "unknown_scope", "scope": s, "valid_scopes": ["pdf", "txt", "markdown"]}
        """
        # GATE 1: WORD_ENABLE_WRITE env var (checked first, before any I/O)
        if os.getenv("WORD_ENABLE_WRITE", "false").lower() != "true":
            raise ToolError(
                "NotAllowedError: WORD_ENABLE_WRITE is not set to 'true'. "
                "Set WORD_ENABLE_WRITE=true to enable file export operations."
            )

        # GATE 2: confirm parameter
        if not confirm:
            raise ToolError(
                "ValidationError: confirm=True is required for export operations. "
                "Pass confirm=True to proceed."
            )

        s = scope.strip().lower()

        if s == "txt":
            import wordmcp._docx.export as _exp
            try:
                return _exp.export_as_file_txt(path=path, output_path=output_path)
            except Exception as exc:
                _rethrow(exc)

        if s == "markdown":
            import wordmcp._docx.export as _exp
            try:
                return _exp.export_markdown_to_file(path=path, output_path=output_path)
            except Exception as exc:
                _rethrow(exc)

        if s == "pdf":
            # Route to COM backend
            from wordmcp import _server_com_tools
            try:
                return _server_com_tools.export_document(
                    runtime,
                    path=path,
                    output_path=output_path,
                    format=s,
                    confirm=confirm,
                )
            except ToolError:
                raise
            except Exception as exc:
                _rethrow(exc)

        # Unknown scope — structured error, not an exception
        return {
            "error": "unknown_scope",
            "scope": scope,
            "valid_scopes": list(_VALID_SCOPES),
        }

    return {
        "export": bind_tool(_export, "export"),
    }


def _rethrow(exc: Exception) -> None:
    """Translate domain errors to ToolError; re-raise ToolError unchanged."""
    if isinstance(exc, ToolError):
        raise
    raise ToolError(f"{type(exc).__name__}: {exc}") from exc
