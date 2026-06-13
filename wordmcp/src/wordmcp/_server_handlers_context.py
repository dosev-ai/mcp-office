from __future__ import annotations

from collections.abc import Callable
from typing import Any

from wordmcp import _server_context_tools
from wordmcp._server_runtime import ServerRuntime


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _get_document_context(path: str) -> dict[str, Any]:
        """Return a summary-only context packet for a .docx file.

        The payload is read-only and limited to document metadata, aggregate counts,
        and heading text. It does not include paragraph bodies, run text, table cell
        contents, comments, or tracked changes.

        DEPRECATED: Use read_document(path, scope='context') instead.
        This alias is kept for backward compatibility and returns identical output.
        """
        return _server_context_tools.get_document_context(runtime, path)

    return {
        "get_document_context": bind_tool(
            _get_document_context,
            "get_document_context",
        ),
    }
