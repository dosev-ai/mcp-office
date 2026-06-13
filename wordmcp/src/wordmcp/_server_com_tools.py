from __future__ import annotations

from wordmcp._server_runtime import ServerRuntime


def manage_tracked_changes(
    runtime: ServerRuntime,
    path: str,
    operation: str,
    revision_index: int | None = None,
    confirm: bool = False,
) -> dict:
    return runtime.call_com(
        "manage_tracked_changes",
        path=path,
        operation=operation,
        revision_index=revision_index,
        confirm=confirm,
    )


def export_document(
    runtime: ServerRuntime,
    path: str,
    output_path: str,
    format: str = "pdf",
    confirm: bool = False,
) -> dict:
    return runtime.call_com(
        "export_document",
        path=path,
        output_path=output_path,
        format=format,
        confirm=confirm,
    )
