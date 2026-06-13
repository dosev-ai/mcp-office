from __future__ import annotations

from collections.abc import Callable
from typing import Any, Literal

from fastmcp.exceptions import ToolError

from wordmcp import _server_com_tools, _server_docx_tools
from wordmcp._server_runtime import ServerRuntime

VALID_FORMATS = frozenset({"txt", "md", "pdf", "html"})


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _export_document(
        path: str,
        format: Literal["txt", "md", "pdf", "html"] = "txt",  # noqa: A002
        output_path: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """Export document to a specified format.

        DEPRECATED: Use the `export` tool instead.
          export(path, scope='pdf', output_path=..., confirm=True)      → replaces format='pdf'
          export(path, scope='html', output_path=..., confirm=True)     → replaces format='html'
          export(path, scope='txt', output_path=..., confirm=True)      → writes .txt file
          export(path, scope='markdown', output_path=..., confirm=True) → writes .md file

        This alias is kept for backward compatibility and continues to function identically.

        Formats (no COM required, no write gate):
          'txt' — plain text; capped at WORD_MAX_TEXT_CHARS.
          'md'  — Markdown with headings, lists, tables, bold/italic.

        Formats (COM required + WORD_ENABLE_WRITE=true + confirm=True):
          'pdf'  — PDF via Word COM; output_path required.
          'html' — Filtered HTML via Word COM SaveAs2; output_path required.

        Raises ToolError for unknown format values.
        """
        if format not in VALID_FORMATS:
            raise ToolError(
                f"Unknown format: {format!r}. Valid: {sorted(VALID_FORMATS)}"
            )
        if format == "txt":
            return _server_docx_tools.export_as_text(runtime, path=path)
        if format == "md":
            return _server_docx_tools.export_to_markdown(runtime, path=path)
        if output_path is None:
            raise ToolError(f"output_path is required for format={format!r}")
        return _server_com_tools.export_document(
            runtime,
            path=path,
            output_path=output_path,
            format=format,
            confirm=confirm,
        )

    return {
        "export_document": bind_tool(_export_document, "export_document"),
    }
