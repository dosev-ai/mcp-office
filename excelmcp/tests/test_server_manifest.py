"""Phase 3D: TOOL_REGISTRY and capabilities_v2 test suite.

CI-safe tests covering:
- TOOL_REGISTRY structure and counts
- build_capabilities_v2() output shape and values
- session_mode removal from server_io.py tools

No COM required. All tests runnable on any OS.
"""
from __future__ import annotations

import inspect
import os
import sys


# ---------------------------------------------------------------------------
# Ensure excelmcp is importable; EXCEL_ALLOWLIST_ROOTS must be set before
# any server module imports to avoid ValidationError on path checks.
# ---------------------------------------------------------------------------
os.environ.setdefault("EXCEL_ALLOWLIST_ROOTS", "C:/Temp")

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from excelmcp._server_manifest import (
    CELL_COMMENT_OPERATIONS,
    CELL_OPERATIONS,
    CHART_OPERATIONS,
    COLS_OPERATIONS,
    DATA_VALIDATION_OPERATIONS,
    DEPRECATION_POLICY,
    EXPORT_AS_PDF_SCOPES,
    FORMAT_OPERATIONS,
    GET_WORKBOOK_CONTEXT_LEVELS,
    HYPERLINK_OPERATIONS,
    MERGE_OPERATIONS,
    NAMED_RANGE_OPERATIONS,
    PIVOT_TABLE_OPERATIONS,
    PROTECT_SCOPES,
    RANGE_IO_OPERATIONS,
    ROWS_OPERATIONS,
    SHEET_OPERATIONS,
    SHEET_PROPERTIES_OPERATIONS,
    TOOL_REGISTRY,
    WORKBOOK_METADATA_OPERATIONS,
    build_capabilities_v2,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _v2():
    """Return a fresh capabilities_v2 dict."""
    return build_capabilities_v2(None)


# ===========================================================================
# 1. test_registry_entry_count
# ===========================================================================
def test_registry_entry_count():
    assert len(TOOL_REGISTRY) == 65, f"Expected 65, got {len(TOOL_REGISTRY)}"


# ===========================================================================
# 2. test_registry_entry_structure
# ===========================================================================
def test_registry_entry_structure():
    valid_kinds = {"dispatcher", "standalone", "deprecated_alias"}
    valid_modes = {"openpyxl", "com_conditional", "com_session"}
    valid_gate_values = {"write", "com", "macros"}
    required_keys = {"kind", "replacement_tool", "replacement_operation_or_scope",
                     "operations_or_scopes", "mode", "gates"}

    for name, entry in TOOL_REGISTRY.items():
        assert set(entry.keys()) == required_keys, (
            f"{name}: expected 6 keys, got {sorted(entry.keys())}"
        )
        assert entry["kind"] in valid_kinds, f"{name}: invalid kind {entry['kind']!r}"
        assert entry["mode"] in valid_modes, f"{name}: invalid mode {entry['mode']!r}"
        assert isinstance(entry["gates"], list), f"{name}: gates must be a list"
        for g in entry["gates"]:
            assert g in valid_gate_values, f"{name}: invalid gate {g!r}"
        if entry["kind"] == "deprecated_alias":
            assert isinstance(entry["replacement_tool"], str), (
                f"{name}: deprecated_alias must have a replacement_tool str"
            )
            assert isinstance(entry["replacement_operation_or_scope"], str), (
                f"{name}: deprecated_alias must have a replacement_operation_or_scope str"
            )
        else:
            assert entry["replacement_tool"] is None, f"{name}: replacement_tool must be None"
            assert entry["replacement_operation_or_scope"] is None, (
                f"{name}: replacement_operation_or_scope must be None"
            )


# ===========================================================================
# 3. test_registry_dispatcher_count
# ===========================================================================
def test_registry_dispatcher_count():
    dispatchers = [
        name for name, e in TOOL_REGISTRY.items()
        if e["kind"] == "dispatcher"
    ]
    assert len(dispatchers) == 18, f"Expected 18 dispatchers, got {len(dispatchers)}: {sorted(dispatchers)}"
    assert "format" in dispatchers, "format dispatcher missing from TOOL_REGISTRY"
    for name in dispatchers:
        assert TOOL_REGISTRY[name]["operations_or_scopes"] is not None, (
            f"Dispatcher {name!r} must have non-None operations_or_scopes"
        )


# ===========================================================================
# 4. test_registry_standalone_count
# ===========================================================================
def test_registry_standalone_count():
    standalones = [
        name for name, e in TOOL_REGISTRY.items()
        if e["kind"] == "standalone"
    ]
    assert len(standalones) == 39, f"Expected 39 standalones, got {len(standalones)}"
    for name in standalones:
        assert TOOL_REGISTRY[name]["operations_or_scopes"] is None, (
            f"Standalone {name!r} must have operations_or_scopes=None"
        )


# ===========================================================================
# 5. test_registry_deprecated_aliases_count_and_replacements
# ===========================================================================
def test_registry_deprecated_aliases_count_and_replacements():
    deprecated = {
        name: e for name, e in TOOL_REGISTRY.items()
        if e["kind"] == "deprecated_alias"
    }
    assert len(deprecated) == 8, f"Expected 8 deprecated aliases, got {len(deprecated)}: {sorted(deprecated)}"
    expected = {
        "apply_number_format": ("format", "number_format"),
        "apply_style":         ("format", "style"),
        "apply_alignment":     ("format", "alignment"),
        "add_border":          ("format", "border"),
        "apply_format_to_sheet_list": ("format", "sheet_list"),
        "copy_range_format":   ("format", "copy_format"),
        "write_range_with_format": ("format", "write_with_format"),
        "apply_conditional_format": ("format", "conditional"),
    }
    for name, (rt, ros) in expected.items():
        assert name in deprecated, f"{name} missing from deprecated_alias entries"
        assert deprecated[name]["replacement_tool"] == rt, (
            f"{name}: replacement_tool should be {rt!r}"
        )
        assert deprecated[name]["replacement_operation_or_scope"] == ros, (
            f"{name}: replacement_operation_or_scope should be {ros!r}"
        )


# ===========================================================================
# 6. test_primary_tools_count_and_sort
# ===========================================================================
def test_primary_tools_count_and_sort():
    v2 = _v2()
    primary = v2["primary_tools"]
    assert len(primary) == 57, f"Expected 57 primary tools, got {len(primary)}"
    assert primary == sorted(primary), "primary_tools must be sorted"
    # spot checks: dispatchers and standalones present
    assert "capabilities" in primary
    assert "sheet" in primary
    assert "format" in primary
    assert "run_macro" in primary
    # deprecated aliases must NOT appear in primary_tools
    for alias in ("apply_number_format", "apply_style", "apply_alignment",
                  "add_border", "apply_format_to_sheet_list", "copy_range_format",
                  "write_range_with_format", "apply_conditional_format"):
        assert alias not in primary, f"{alias} should not be in primary_tools"


# ===========================================================================
# 7. test_deprecated_aliases_in_capabilities_v2
# ===========================================================================
def test_deprecated_aliases_in_capabilities_v2():
    v2 = _v2()
    aliases = v2["deprecated_aliases"]
    assert len(aliases) == 8, f"Expected 8 deprecated_aliases, got {len(aliases)}"
    assert aliases == sorted(aliases), "deprecated_aliases must be sorted"
    rt = v2["replacement_tool"]
    ros = v2["replacement_operation_or_scope"]
    assert all(v == "format" for v in rt.values()), "All aliases should replace to 'format'"
    assert set(rt.keys()) == set(aliases)
    assert set(ros.keys()) == set(aliases)
    # spot-check one alias
    assert rt["apply_number_format"] == "format"
    assert ros["apply_number_format"] == "number_format"


# ===========================================================================
# 8. test_total_callable_endpoints
# ===========================================================================
def test_total_callable_endpoints():
    v2 = _v2()
    assert v2["total_callable_endpoints"] == 65, (
        f"Expected 65, got {v2['total_callable_endpoints']}"
    )
    # parity equation: primary + deprecated == total
    assert len(v2["primary_tools"]) + len(v2["deprecated_aliases"]) == v2["total_callable_endpoints"]


# ===========================================================================
# 9. test_operation_scope_enums_count_and_sort
# ===========================================================================
def test_operation_scope_enums_count_and_sort():
    v2 = _v2()
    ose = v2["operation_scope_enums"]
    assert len(ose) == 18, f"Expected 18 dispatcher enums, got {len(ose)}: {sorted(ose.keys())}"
    for name, ops in ose.items():
        assert ops == sorted(ops), f"{name}: operations must be sorted"


# ===========================================================================
# 10. test_operation_scope_enums_exact_values
# ===========================================================================
def test_operation_scope_enums_exact_values():
    v2 = _v2()
    ose = v2["operation_scope_enums"]
    assert ose["format"] == sorted(FORMAT_OPERATIONS)
    assert ose["sheet"] == sorted(SHEET_OPERATIONS)
    assert ose["cell"] == sorted(CELL_OPERATIONS)
    assert ose["range_io"] == sorted(RANGE_IO_OPERATIONS)
    assert ose["named_range"] == sorted(NAMED_RANGE_OPERATIONS)
    assert ose["workbook_metadata"] == sorted(WORKBOOK_METADATA_OPERATIONS)
    assert ose["merge"] == sorted(MERGE_OPERATIONS)
    assert ose["rows"] == sorted(ROWS_OPERATIONS)
    assert ose["cols"] == sorted(COLS_OPERATIONS)
    assert ose["protect"] == sorted(PROTECT_SCOPES)
    assert ose["cell_comment"] == sorted(CELL_COMMENT_OPERATIONS)
    assert ose["sheet_properties"] == sorted(SHEET_PROPERTIES_OPERATIONS)
    assert ose["data_validation"] == sorted(DATA_VALIDATION_OPERATIONS)
    assert ose["hyperlink"] == sorted(HYPERLINK_OPERATIONS)
    assert ose["chart"] == sorted(CHART_OPERATIONS)
    assert ose["export_as_pdf"] == sorted(EXPORT_AS_PDF_SCOPES)
    assert ose["pivot_table"] == sorted(PIVOT_TABLE_OPERATIONS)
    assert ose["get_workbook_context"] == sorted(GET_WORKBOOK_CONTEXT_LEVELS)


# ===========================================================================
# 11. test_write_gate_metadata_coverage
# ===========================================================================
def test_write_gate_metadata_coverage():
    v2 = _v2()
    wgm = v2["write_gate_metadata"]
    # Write-gated tools that must be present
    write_gated = [name for name, e in TOOL_REGISTRY.items() if "write" in e["gates"]]
    for name in write_gated:
        assert name in wgm, f"{name} should be in write_gate_metadata"
        assert wgm[name]["env_var"] == "EXCEL_ENABLE_WRITE"
        assert wgm[name]["requires_confirm"] is True
    # Read-only tools must NOT be in write_gate_metadata
    read_only = [name for name, e in TOOL_REGISTRY.items() if "write" not in e["gates"]]
    for name in read_only:
        assert name not in wgm, f"{name} should not be in write_gate_metadata"


# ===========================================================================
# 12. test_gate_classes_structure
# ===========================================================================
def test_gate_classes_structure():
    v2 = _v2()
    gc = v2["gate_classes"]
    assert len(gc) == 4
    kinds = {c["kind"] for c in gc}
    assert kinds == {"write", "com", "macros", "allowlist"}
    # allowlist must have scope=universal
    allowlist_entry = next(c for c in gc if c["kind"] == "allowlist")
    assert allowlist_entry["scope"] == "universal"
    assert allowlist_entry["env_var"] == "EXCEL_ALLOWLIST_ROOTS"
    assert allowlist_entry["requires_confirm"] is False


# ===========================================================================
# 13. test_gate_metadata_run_macro
# ===========================================================================
def test_gate_metadata_run_macro():
    v2 = _v2()
    gm = v2["gate_metadata"]
    assert "run_macro" in gm
    gate_kinds = {g["kind"] for g in gm["run_macro"]["gates"]}
    assert gate_kinds == {"write", "com", "macros"}, f"run_macro gates: {gate_kinds}"


# ===========================================================================
# 14. test_gate_metadata_list_macros
# ===========================================================================
def test_gate_metadata_list_macros():
    v2 = _v2()
    gm = v2["gate_metadata"]
    assert "list_macros" in gm
    gate_kinds = {g["kind"] for g in gm["list_macros"]["gates"]}
    assert gate_kinds == {"com"}, f"list_macros gates: {gate_kinds}"


# ===========================================================================
# 15. test_gate_metadata_recalculate_workbook
# ===========================================================================
def test_gate_metadata_recalculate_workbook():
    v2 = _v2()
    gm = v2["gate_metadata"]
    assert "recalculate_workbook" in gm
    gate_kinds = {g["kind"] for g in gm["recalculate_workbook"]["gates"]}
    assert gate_kinds == {"com", "write"}, f"recalculate_workbook gates: {gate_kinds}"


# ===========================================================================
# 16. test_gate_metadata_readonly_tools_absent
# ===========================================================================
def test_gate_metadata_readonly_tools_absent():
    v2 = _v2()
    gm = v2["gate_metadata"]
    read_only_tools = ["capabilities", "get_used_range", "diff_workbooks"]
    for name in read_only_tools:
        assert name not in gm, f"{name} should not appear in gate_metadata (no gates)"


# ===========================================================================
# 17. test_gate_metadata_write_consistency
# ===========================================================================
def test_gate_metadata_write_consistency():
    v2 = _v2()
    gm = v2["gate_metadata"]
    wgm = v2["write_gate_metadata"]
    # All tools in write_gate_metadata must appear in gate_metadata
    for name in wgm:
        assert name in gm, f"{name} in write_gate_metadata but absent from gate_metadata"
    # The set of tools with a "write" gate in gate_metadata must equal write_gate_metadata keys
    gm_write_tools = {
        name for name, meta in gm.items()
        if any(g["kind"] == "write" for g in meta["gates"])
    }
    assert gm_write_tools == set(wgm.keys()), (
        f"Mismatch: gate_metadata write-tools={sorted(gm_write_tools)}, "
        f"write_gate_metadata keys={sorted(wgm.keys())}"
    )


# ===========================================================================
# 18. test_tool_modes_count_and_values
# ===========================================================================
def test_tool_modes_count_and_values():
    v2 = _v2()
    tm = v2["tool_modes"]
    assert len(tm) == 65, f"Expected 65 tool_modes, got {len(tm)}"
    openpyxl_count = sum(1 for m in tm.values() if m == "openpyxl")
    com_cond_count = sum(1 for m in tm.values() if m == "com_conditional")
    com_sess_count = sum(1 for m in tm.values() if m == "com_session")
    assert openpyxl_count == 57, f"Expected 57 openpyxl, got {openpyxl_count}"
    assert com_cond_count == 6, f"Expected 6 com_conditional, got {com_cond_count}"
    assert com_sess_count == 2, f"Expected 2 com_session, got {com_sess_count}"


# ===========================================================================
# 19. test_tool_modes_spot_checks
# ===========================================================================
def test_tool_modes_spot_checks():
    v2 = _v2()
    tm = v2["tool_modes"]
    assert tm["capabilities"] == "openpyxl"
    assert tm["format"] == "openpyxl"
    assert tm["sheet"] == "openpyxl"
    assert tm["export_as_pdf"] == "com_conditional"
    assert tm["pivot_table"] == "com_conditional"
    assert tm["recalculate_workbook"] == "com_conditional"
    assert tm["hydrate_template"] == "com_conditional"
    assert tm["list_macros"] == "com_conditional"
    assert tm["run_macro"] == "com_conditional"
    assert tm["get_active_workbook_context"] == "com_session"
    assert tm["get_all_open_workbooks"] == "com_session"
    assert tm["diff_workbooks"] == "openpyxl"


# ===========================================================================
# 20. test_deprecation_policy
# ===========================================================================
def test_deprecation_policy():
    v2 = _v2()
    dp = v2["deprecation_policy"]
    assert dp["window_releases"] == 2
    assert dp["removal_date_iso"] is None
    assert dp["telemetry_field"] == "deprecated_alias_called"
    # Also check the module-level constant
    assert DEPRECATION_POLICY["window_releases"] == 2
    assert DEPRECATION_POLICY["removal_date_iso"] is None


# ===========================================================================
# 21. test_backward_compat_capabilities_top_level
# ===========================================================================
def test_backward_compat_capabilities_top_level(monkeypatch):
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", "C:/Temp")
    from excelmcp._io import capabilities
    caps = capabilities([])
    # capabilities_v2 must be nested, not merged into top level
    assert "capabilities_v2" in caps
    v2 = caps["capabilities_v2"]
    assert "primary_tools" in v2
    # Top-level capabilities keys must NOT include capabilities_v2 contents directly
    assert "primary_tools" not in caps
    assert "operation_scope_enums" not in caps
    assert "gate_classes" not in caps


# ===========================================================================
# 22. test_parity_flat_tools_vs_registry
# ===========================================================================
def test_parity_flat_tools_vs_registry(monkeypatch):
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", "C:/Temp")
    from excelmcp._io import capabilities
    caps = capabilities([])
    tools_in_caps = set(caps["tools"])
    tools_in_registry = set(TOOL_REGISTRY.keys())
    assert tools_in_caps == tools_in_registry, (
        f"Parity mismatch:\n"
        f"  in caps but not registry: {sorted(tools_in_caps - tools_in_registry)}\n"
        f"  in registry but not caps: {sorted(tools_in_registry - tools_in_caps)}"
    )


# ===========================================================================
# 23. test_session_mode_removed_from_cell
# ===========================================================================
def test_session_mode_removed_from_cell():
    from excelmcp.server_io import cell
    params = list(inspect.signature(cell).parameters.keys())
    assert "session_mode" not in params, f"cell still has session_mode: {params}"


# ===========================================================================
# 24. test_session_mode_removed_from_range_io
# ===========================================================================
def test_session_mode_removed_from_range_io():
    from excelmcp.server_io import range_io
    params = list(inspect.signature(range_io).parameters.keys())
    assert "session_mode" not in params, f"range_io still has session_mode: {params}"


# ===========================================================================
# 25. test_session_mode_removed_from_append_rows
# ===========================================================================
def test_session_mode_removed_from_append_rows():
    from excelmcp.server_io import append_rows
    params = list(inspect.signature(append_rows).parameters.keys())
    assert "session_mode" not in params, f"append_rows still has session_mode: {params}"


# ===========================================================================
# 26. test_session_mode_removed_from_create_workbook
# ===========================================================================
def test_session_mode_removed_from_create_workbook():
    from excelmcp.server_io import create_workbook
    params = list(inspect.signature(create_workbook).parameters.keys())
    assert "session_mode" not in params, f"create_workbook still has session_mode: {params}"


