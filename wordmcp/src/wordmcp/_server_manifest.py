"""wordmcp server manifest — capabilities and tool registry."""
from __future__ import annotations

from typing import Any

from wordmcp._server_runtime import ServerRuntime

from wordmcp._manifest_enums import (  # noqa: F401  (re-exported for callers)
    COM_TOOLS,
    CONTENT_OPERATIONS,
    CONTEXT_TOOLS,
    DOCUMENT_OPERATIONS,
    EXPORT_SCOPES,
    EXTRA_TOOLS,
    PARAGRAPH_OPERATIONS,
    READ_DOCUMENT_SCOPES,
    REVIEW_OPERATIONS,
    STYLE_OPERATIONS,
    TABLE_OPERATIONS,
)
from wordmcp._manifest_registry import (  # noqa: F401  (re-exported for callers)
    DEPRECATION_POLICY,
    TOOL_REGISTRY,
    _WRITE_GATE,
)

# ---------------------------------------------------------------------------
# V2 capabilities builder
# ---------------------------------------------------------------------------

_DISPATCHERS = frozenset(
    name for name, meta in TOOL_REGISTRY.items() if meta["kind"] == "dispatcher"
)
_PRIMARY_KINDS = frozenset({"dispatcher", "standalone"})


def build_capabilities_v2(
    runtime: Any,
    active_tools: list[str] | None = None,
) -> dict:
    """Return v2 schema-conformant capabilities dict.

    Parameters
    ----------
    runtime:
        Passed for signature compatibility; not used in v2 derivation (the
        registry is the source of truth, not the live server).
    active_tools:
        Reserved for compact-surface filtering (e.g. PPTMCP). When ``None``
        all entries from ``TOOL_REGISTRY`` are used.
    """
    registry = TOOL_REGISTRY
    if active_tools is not None:
        registry = {k: v for k, v in TOOL_REGISTRY.items() if k in active_tools}

    primary_tools: list[str] = sorted(
        name for name, meta in registry.items() if meta["kind"] in _PRIMARY_KINDS
    )
    deprecated_aliases: list[str] = sorted(
        name for name, meta in registry.items() if meta["kind"] == "deprecated_alias"
    )

    replacement_tool: dict[str, str] = {
        name: meta["replacement_tool"]  # type: ignore[assignment]
        for name, meta in registry.items()
        if meta["kind"] == "deprecated_alias"
    }
    replacement_operation_or_scope: dict[str, str] = {
        name: meta["replacement_operation_or_scope"] or ""
        for name, meta in registry.items()
        if meta["kind"] == "deprecated_alias"
    }
    total_callable_endpoints: int = len(primary_tools) + len(deprecated_aliases)

    # Validate replacement_tool values reference existing primary tools
    primary_set = set(primary_tools)
    for alias, repl in replacement_tool.items():
        if repl not in primary_set:
            raise ValueError(
                f"Deprecated alias {alias!r} references replacement_tool={repl!r} "
                f"which is not in primary_tools {sorted(primary_set)}"
            )

    # operation_scope_enums: dispatcher -> sorted operations/scopes
    operation_scope_enums: dict[str, list[str]] = {
        name: sorted(meta["operations_or_scopes"])  # type: ignore[arg-type]
        for name, meta in registry.items()
        if meta["kind"] == "dispatcher" and meta["operations_or_scopes"]
    }

    # write_gate_metadata: every write-gated tool (primary + deprecated)
    write_gate_metadata: dict[str, dict] = {
        name: dict(_WRITE_GATE)
        for name, meta in registry.items()
        if meta["write_gated"]
    }

    return {
        "primary_tools": primary_tools,
        "deprecated_aliases": deprecated_aliases,
        "replacement_tool": replacement_tool,
        "replacement_operation_or_scope": replacement_operation_or_scope,
        "total_callable_endpoints": total_callable_endpoints,
        "operation_scope_enums": operation_scope_enums,
        "write_gate_metadata": write_gate_metadata,
        "deprecation_policy": dict(DEPRECATION_POLICY),
    }


def build_capabilities(runtime: ServerRuntime) -> dict:
    result = runtime.call_doc("capabilities")
    result["com_tools"] = {
        "loaded": runtime.is_com_loaded(),
        "tools": list(COM_TOOLS),
        "label": "com: Word COM automation tools",
    }
    result["context_tools"] = {
        "tools": list(CONTEXT_TOOLS),
        "label": "read-only context tools",
    }
    result["read_document_scopes"] = list(READ_DOCUMENT_SCOPES)
    result["export_scopes"] = list(EXPORT_SCOPES)
    result["table_operations"] = list(TABLE_OPERATIONS)
    result["style_operations"] = list(STYLE_OPERATIONS)
    result["content_operations"] = list(CONTENT_OPERATIONS)
    result["review_operations"] = list(REVIEW_OPERATIONS)
    result["document_operations"] = list(DOCUMENT_OPERATIONS)
    result["tools"] = result["tools"] + list(EXTRA_TOOLS)
    # v2 structured fields — additive, nested under capabilities_v2 key
    result["capabilities_v2"] = build_capabilities_v2(runtime)
    return result
