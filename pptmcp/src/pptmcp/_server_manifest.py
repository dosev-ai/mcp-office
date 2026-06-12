from __future__ import annotations

from typing import Any

# ---------------------------------------------------------------------------
# Vocabulary constants
# ---------------------------------------------------------------------------

SLIDE_OPERATIONS = ["add", "copy", "delete", "reorder"]

SHAPE_OPERATIONS = [
    "add_autoshape", "add_table", "add_text_box",
    "delete", "set_geometry", "set_properties", "set_table_style",
]

EXPORT_SCOPES = ["deck_pdf", "slide", "slide_images", "slide_text"]
# NOTE: "slide_png" is intentionally ABSENT — deprecated in favour of "slide"

SET_FORMAT_TARGETS = ["paragraph", "text"]

MANAGE_HYPERLINKS_OPERATIONS = ["add", "list", "remove"]

# ---------------------------------------------------------------------------
# Deprecation policy
# ---------------------------------------------------------------------------

DEPRECATION_POLICY: dict = {
    "window_releases": 2,
    "telemetry_field": "deprecated_alias_called",
    "removal_date_iso": None,
}

# ---------------------------------------------------------------------------
# Per-tool registry
#
# Each entry is a plain dict with these keys:
#   kind: "dispatcher" | "standalone" | "deprecated_alias"
#   replacement_tool: str | None          — only for deprecated_alias
#   replacement_operation_or_scope: str | None  — only for deprecated_alias
#   write_gated: bool
#   operations_or_scopes: list[str] | None     — only for dispatcher
# ---------------------------------------------------------------------------

_WRITE_GATE = {"env_var": "PPT_ENABLE_WRITE", "requires_confirm": True}

# TOOL_REGISTRY is a dict: tool_name -> metadata dict.
# Total: 5 dispatchers + 1 deprecated_alias + 40 standalones = 46 entries.
TOOL_REGISTRY: dict[str, dict] = {
    # ------------------------------------------------------------------
    # Dispatcher primary tools (5)
    # ------------------------------------------------------------------
    "export": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(EXPORT_SCOPES),
    },
    "manage_hyperlinks": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(MANAGE_HYPERLINKS_OPERATIONS),
    },
    "set_format": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(SET_FORMAT_TARGETS),
    },
    "shape": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(SHAPE_OPERATIONS),
    },
    "slide": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": sorted(SLIDE_OPERATIONS),
    },
    # ------------------------------------------------------------------
    # Deprecated alias (1)
    # ------------------------------------------------------------------
    "add_content": {
        "kind": "deprecated_alias",
        "replacement_tool": "shape",
        "replacement_operation_or_scope": "",
        "write_gated": True,
        "operations_or_scopes": None,
    },
    # ------------------------------------------------------------------
    # Standalone primary tools (40)
    # ------------------------------------------------------------------
    "add_hyperlink": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "add_slide": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "apply_slide_layout": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "capabilities": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "check_presentation_against_contract": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "clear_slide_content": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "copy_slide": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "create_presentation": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "declare_slide_contract": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "delete_shape": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "delete_slide": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "detect_overlapping_shapes": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "edit_text_placeholder": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "export_changed_slides_only": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "export_deck": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "export_slide": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "export_slide_as_text": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "export_slides_to_stamped_dir": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "extract_images": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "extract_presentation_text": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "extract_tables": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "get_presentation_context": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "get_presentation_metadata": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "get_shape": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "insert_image": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "list_layouts": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "list_shapes": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "list_slides": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "produce_evidence_bundle": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "read_presentation": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "read_slide": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "read_speaker_notes": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "remove_empty_placeholders": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "replace_slide_text": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "reorder_slides": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "review_slide_export": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
    "save": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "set_speaker_notes": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "set_table_style": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": True,
        "operations_or_scopes": None,
    },
    "validate_contract": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "write_gated": False,
        "operations_or_scopes": None,
    },
}


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
        Reserved for compact-surface filtering. When ``None``
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
