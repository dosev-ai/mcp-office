"""ExcelMCP server manifest — TOOL_REGISTRY, vocabulary constants, and build_capabilities_v2.

Single source of truth for:
  - Tool vocabulary (operation/scope enum lists)
  - Gate class definitions
  - Per-tool metadata (kind, mode, gates)
  - capabilities_v2 builder

No MCP/FastMCP imports. No COM/pywin32 imports. Pure stdlib only.
"""
from __future__ import annotations

# ---------------------------------------------------------------------------
# Vocabulary constants (sorted lists, lowercase strings)
# ---------------------------------------------------------------------------

FORMAT_OPERATIONS = [
    "alignment",
    "border",
    "conditional",
    "copy_format",
    "number_format",
    "sheet_list",
    "style",
    "write_with_format",
]

SHEET_OPERATIONS             = ["add", "delete", "list", "rename"]
CELL_OPERATIONS              = ["read", "write"]
RANGE_IO_OPERATIONS          = ["read", "write"]
NAMED_RANGE_OPERATIONS       = ["list", "read", "write"]
WORKBOOK_METADATA_OPERATIONS = ["read", "write"]
MERGE_OPERATIONS             = ["list", "merge", "unmerge"]
ROWS_OPERATIONS              = ["delete", "insert"]
COLS_OPERATIONS              = ["delete", "insert"]
PROTECT_SCOPES               = ["sheet", "workbook"]
CELL_COMMENT_OPERATIONS      = ["delete", "get"]
SHEET_PROPERTIES_OPERATIONS  = ["get", "set"]
DATA_VALIDATION_OPERATIONS   = ["add", "get"]
HYPERLINK_OPERATIONS         = ["add", "get", "remove"]
CHART_OPERATIONS             = ["create", "delete", "list", "update"]
EXPORT_AS_PDF_SCOPES         = ["sheet", "workbook"]
PIVOT_TABLE_OPERATIONS       = ["create", "read", "refresh"]
GET_WORKBOOK_CONTEXT_LEVELS  = ["deep", "focused", "index"]

# ---------------------------------------------------------------------------
# Gate class definitions
# ---------------------------------------------------------------------------

GATE_CLASSES = [
    {"kind": "write",     "env_var": "EXCEL_ENABLE_WRITE",    "requires_confirm": True,  "scope": "per-tool"},
    {"kind": "com",       "env_var": "EXCEL_ENABLE_COM",      "requires_confirm": False, "scope": "per-tool"},
    {"kind": "macros",    "env_var": "EXCEL_ENABLE_MACROS",   "requires_confirm": False, "scope": "per-tool"},
    {"kind": "allowlist", "env_var": "EXCEL_ALLOWLIST_ROOTS", "requires_confirm": False, "scope": "universal"},
]

# ---------------------------------------------------------------------------
# Deprecation policy and write-gate shorthand
# ---------------------------------------------------------------------------

DEPRECATION_POLICY = {
    "window_releases": 2,
    "telemetry_field": "deprecated_alias_called",
    "removal_date_iso": None,
}

_WRITE_GATE = {"env_var": "EXCEL_ENABLE_WRITE", "requires_confirm": True}

# ---------------------------------------------------------------------------
# TOOL_REGISTRY — 65 entries, exactly 6 keys per entry
# ---------------------------------------------------------------------------
# Entry shape:
#   kind:                        "dispatcher" | "standalone" | "deprecated_alias"
#   replacement_tool:            None for dispatchers/standalones, str for aliases
#   replacement_operation_or_scope: None or str pointing to dispatcher operation
#   operations_or_scopes:        list for dispatchers, None for standalones
#   mode:                        "openpyxl" | "com_conditional" | "com_session"
#   gates:                       list of strings from {"write", "com", "macros"}

TOOL_REGISTRY: dict[str, dict] = {
    # ------------------------------------------------------------------
    # 18 dispatchers
    # ------------------------------------------------------------------
    "format": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": FORMAT_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "sheet": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": SHEET_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "cell": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": CELL_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "range_io": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": RANGE_IO_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "named_range": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": NAMED_RANGE_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "workbook_metadata": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": WORKBOOK_METADATA_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "merge": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": MERGE_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "rows": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": ROWS_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "cols": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": COLS_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "protect": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": PROTECT_SCOPES,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "cell_comment": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": CELL_COMMENT_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "sheet_properties": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": SHEET_PROPERTIES_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "data_validation": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": DATA_VALIDATION_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "hyperlink": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": HYPERLINK_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "chart": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": CHART_OPERATIONS,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "export_as_pdf": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": EXPORT_AS_PDF_SCOPES,
        "mode": "com_conditional",
        "gates": ["com", "write"],
    },
    "pivot_table": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": PIVOT_TABLE_OPERATIONS,
        "mode": "com_conditional",
        "gates": ["com", "write"],
    },
    "get_workbook_context": {
        "kind": "dispatcher",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": GET_WORKBOOK_CONTEXT_LEVELS,
        "mode": "openpyxl",
        "gates": [],
    },
    # ------------------------------------------------------------------
    # 48 standalones — server_io.py (13)
    # ------------------------------------------------------------------
    "capabilities": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    "create_workbook": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "get_used_range": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    "read_sheet_all": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    "get_cell_metadata": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    "evaluate_formula": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    "list_tables": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    "read_table": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    "append_rows": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "find_replace": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "import_csv_to_sheet": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "save": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "consolidate_ranges": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    # ------------------------------------------------------------------
    # server_format.py deprecated aliases (8) — replaced by "format" dispatcher
    # ------------------------------------------------------------------
    "apply_number_format": {
        "kind": "deprecated_alias",
        "replacement_tool": "format",
        "replacement_operation_or_scope": "number_format",
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "apply_style": {
        "kind": "deprecated_alias",
        "replacement_tool": "format",
        "replacement_operation_or_scope": "style",
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "apply_alignment": {
        "kind": "deprecated_alias",
        "replacement_tool": "format",
        "replacement_operation_or_scope": "alignment",
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "add_border": {
        "kind": "deprecated_alias",
        "replacement_tool": "format",
        "replacement_operation_or_scope": "border",
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "apply_format_to_sheet_list": {
        "kind": "deprecated_alias",
        "replacement_tool": "format",
        "replacement_operation_or_scope": "sheet_list",
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "copy_range_format": {
        "kind": "deprecated_alias",
        "replacement_tool": "format",
        "replacement_operation_or_scope": "copy_format",
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "write_range_with_format": {
        "kind": "deprecated_alias",
        "replacement_tool": "format",
        "replacement_operation_or_scope": "write_with_format",
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "apply_conditional_format": {
        "kind": "deprecated_alias",
        "replacement_tool": "format",
        "replacement_operation_or_scope": "conditional",
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    # ------------------------------------------------------------------
    # server_ops.py standalones (8)
    # ------------------------------------------------------------------
    "freeze_panes": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "auto_filter": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "validate_formula_syntax": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    "copy_range": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "create_table": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "set_column_visibility": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "set_workbook_calc_mode": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "set_print_area": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    # ------------------------------------------------------------------
    # server_batch.py standalones (6) — all mode=openpyxl, gates=["write"]
    # ------------------------------------------------------------------
    "bulk_copy_sheet": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "set_row_heights": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "set_column_widths": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "bulk_add_comments": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "fill_formula_range": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "bulk_range_write": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    # ------------------------------------------------------------------
    # server_com_export.py standalones (1)
    # ------------------------------------------------------------------
    "export_range_as_csv": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    # ------------------------------------------------------------------
    # server_com_recalc.py standalones (2)
    # ------------------------------------------------------------------
    "recalculate_workbook": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "com_conditional",
        "gates": ["com", "write"],
    },
    "hydrate_template": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "com_conditional",
        "gates": ["com", "write"],
    },
    # ------------------------------------------------------------------
    # server_com_vba.py standalones (2)
    # ------------------------------------------------------------------
    "list_macros": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "com_conditional",
        "gates": ["com"],
    },
    "run_macro": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "com_conditional",
        "gates": ["com", "write", "macros"],
    },
    # ------------------------------------------------------------------
    # server_com_session.py standalones (2)
    # ------------------------------------------------------------------
    "get_active_workbook_context": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "com_session",
        "gates": ["com"],
    },
    "get_all_open_workbooks": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "com_session",
        "gates": ["com"],
    },
    # ------------------------------------------------------------------
    # server_review.py standalones (3)
    # ------------------------------------------------------------------
    "review_workbook_render": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    "produce_export_evidence_bundle": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
    "export_changed_sheets_only": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    # ------------------------------------------------------------------
    # server_snapshot.py standalones (2)
    # ------------------------------------------------------------------
    "snapshot_workbook": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": ["write"],
    },
    "diff_workbooks": {
        "kind": "standalone",
        "replacement_tool": None,
        "replacement_operation_or_scope": None,
        "operations_or_scopes": None,
        "mode": "openpyxl",
        "gates": [],
    },
}


# ---------------------------------------------------------------------------
# build_capabilities_v2
# ---------------------------------------------------------------------------

def build_capabilities_v2(runtime, active_tools=None):
    """Build the capabilities_v2 metadata dict from TOOL_REGISTRY."""
    primary_tools = sorted(
        name for name, e in TOOL_REGISTRY.items() if e["kind"] != "deprecated_alias"
    )
    deprecated_aliases = sorted(
        name for name, e in TOOL_REGISTRY.items() if e["kind"] == "deprecated_alias"
    )
    replacement_tool = {
        name: e["replacement_tool"]
        for name, e in TOOL_REGISTRY.items()
        if e["kind"] == "deprecated_alias"
    }
    replacement_operation_or_scope = {
        name: e["replacement_operation_or_scope"]
        for name, e in TOOL_REGISTRY.items()
        if e["kind"] == "deprecated_alias"
    }
    total_callable_endpoints = len(primary_tools) + len(deprecated_aliases)

    operation_scope_enums = {
        name: sorted(entry["operations_or_scopes"])
        for name, entry in TOOL_REGISTRY.items()
        if entry["operations_or_scopes"] is not None
    }

    write_gate_metadata = {
        name: {"env_var": "EXCEL_ENABLE_WRITE", "requires_confirm": True}
        for name, entry in TOOL_REGISTRY.items()
        if "write" in entry["gates"]
    }

    gate_metadata = {}
    for name, entry in TOOL_REGISTRY.items():
        if not entry["gates"]:
            continue
        gate_entries = []
        for g in entry["gates"]:
            gc = next((c for c in GATE_CLASSES if c["kind"] == g), None)
            if gc:
                gate_entries.append({
                    "kind": gc["kind"],
                    "env_var": gc["env_var"],
                    "requires_confirm": gc["requires_confirm"],
                })
        gate_metadata[name] = {"gates": gate_entries}

    tool_modes = {name: entry["mode"] for name, entry in TOOL_REGISTRY.items()}

    return {
        "primary_tools": primary_tools,
        "deprecated_aliases": deprecated_aliases,
        "replacement_tool": replacement_tool,
        "replacement_operation_or_scope": replacement_operation_or_scope,
        "total_callable_endpoints": total_callable_endpoints,
        "operation_scope_enums": operation_scope_enums,
        "write_gate_metadata": write_gate_metadata,
        "gate_classes": GATE_CLASSES,
        "gate_metadata": gate_metadata,
        "tool_modes": tool_modes,
        "deprecation_policy": DEPRECATION_POLICY,
    }
