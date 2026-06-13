from __future__ import annotations

from collections.abc import Callable
from typing import Any

from wordmcp._server_manifest import build_capabilities
from wordmcp._server_runtime import ServerRuntime


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _capabilities() -> dict:
        """Return metadata about this MCP server phase, backend, and available tools."""
        return build_capabilities(runtime)

    return {
        "capabilities": bind_tool(_capabilities, "capabilities"),
    }
