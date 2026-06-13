from __future__ import annotations

from typing import Any

from wordmcp._server_runtime import ServerRuntime


def get_document_context(runtime: ServerRuntime, path: str) -> dict[str, Any]:
    return runtime.call_doc("get_document_context", path)
