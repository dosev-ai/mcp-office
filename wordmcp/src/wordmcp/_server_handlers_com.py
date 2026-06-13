from __future__ import annotations

from collections.abc import Callable
from typing import Any

from fastmcp.exceptions import ToolError

from wordmcp import _server_com_tools
from wordmcp._server_runtime import ServerRuntime

_VALID_TRACKED_CHANGE_OPS = frozenset({"list", "accept_all", "reject_all"})


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _manage_tracked_changes(
        path: str,
        operation: str,
        revision_index: int | None = None,  # per-revision ops intentionally excluded from MCP surface until Phase 3; revision_index forwarded but currently unreachable
        confirm: bool = False,
    ) -> dict:
        """Manage document tracked changes via COM (requires Word desktop).

        Operations: 'list' | 'accept_all' | 'reject_all'.
        Invalid operation raises ToolError immediately (before any COM call).
        Requires WORD_ENABLE_COM=true. Write operations also require WORD_ENABLE_WRITE=true
        and confirm=True.
        """
        if operation not in _VALID_TRACKED_CHANGE_OPS:
            raise ToolError(
                f"Unknown operation: {operation!r}. "
                f"Valid: {sorted(_VALID_TRACKED_CHANGE_OPS)}"
            )
        return _server_com_tools.manage_tracked_changes(
            runtime,
            path=path,
            operation=operation,
            revision_index=revision_index,
            confirm=confirm,
        )

    return {
        "manage_tracked_changes": bind_tool(
            _manage_tracked_changes,
            "manage_tracked_changes",
        ),
    }
