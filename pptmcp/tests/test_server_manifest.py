"""Unit tests for pptmcp._server_manifest — no COM required."""
from __future__ import annotations

import os
import sys

sys.path.insert(0, str(__file__).split("tests")[0] + "src")

os.environ.setdefault("PPT_ALLOWLIST_ROOTS", r"C:\Temp")

from pptmcp._server_manifest import (  # noqa: E402
    DEPRECATION_POLICY,
    EXPORT_SCOPES,
    MANAGE_HYPERLINKS_OPERATIONS,
    SET_FORMAT_TARGETS,
    SHAPE_OPERATIONS,
    SLIDE_OPERATIONS,
    TOOL_REGISTRY,
    build_capabilities_v2,
)
from pptmcp._server_compact import COMPACT_TOOL_NAMES  # noqa: E402


# ---------------------------------------------------------------------------
# 1. Registry structure
# ---------------------------------------------------------------------------


def test_registry_size():
    assert len(TOOL_REGISTRY) == 46


def test_registry_all_entries_have_required_keys():
    required_keys = {"kind", "replacement_tool", "replacement_operation_or_scope", "write_gated", "operations_or_scopes"}
    for name, meta in TOOL_REGISTRY.items():
        assert set(meta.keys()) == required_keys, f"Entry {name!r} missing keys"


def test_registry_kind_values():
    valid_kinds = {"dispatcher", "standalone", "deprecated_alias"}
    for name, meta in TOOL_REGISTRY.items():
        assert meta["kind"] in valid_kinds, f"Entry {name!r} has invalid kind={meta['kind']!r}"


# ---------------------------------------------------------------------------
# 2. Deprecated alias
# ---------------------------------------------------------------------------


def test_add_content_is_deprecated_alias():
    assert TOOL_REGISTRY["add_content"]["kind"] == "deprecated_alias"


def test_add_content_replacement_tool():
    assert TOOL_REGISTRY["add_content"]["replacement_tool"] == "shape"


def test_add_content_write_gated():
    assert TOOL_REGISTRY["add_content"]["write_gated"] is True


def test_excluded_com_only_tools_absent():
    assert "run_slide_show" not in TOOL_REGISTRY
    assert "recalculate_charts" not in TOOL_REGISTRY


# ---------------------------------------------------------------------------
# 3. Dispatcher enums
# ---------------------------------------------------------------------------


def test_slide_operations():
    assert SLIDE_OPERATIONS == sorted(SLIDE_OPERATIONS)
    assert set(SLIDE_OPERATIONS) == {"add", "copy", "delete", "reorder"}


def test_shape_operations():
    assert SHAPE_OPERATIONS == sorted(SHAPE_OPERATIONS)
    assert "set_geometry" in SHAPE_OPERATIONS
    assert len(SHAPE_OPERATIONS) == 7


def test_export_scopes():
    assert EXPORT_SCOPES == sorted(EXPORT_SCOPES)
    assert "slide_png" not in EXPORT_SCOPES
    assert len(EXPORT_SCOPES) == 4


def test_set_format_targets():
    assert SET_FORMAT_TARGETS == sorted(SET_FORMAT_TARGETS)
    assert set(SET_FORMAT_TARGETS) == {"paragraph", "text"}


def test_manage_hyperlinks_operations():
    assert MANAGE_HYPERLINKS_OPERATIONS == sorted(MANAGE_HYPERLINKS_OPERATIONS)
    assert set(MANAGE_HYPERLINKS_OPERATIONS) == {"add", "list", "remove"}


# ---------------------------------------------------------------------------
# 4. Deprecation policy
# ---------------------------------------------------------------------------


def test_deprecation_policy_has_window_releases():
    assert "window_releases" in DEPRECATION_POLICY


def test_deprecation_policy_window_releases_is_int():
    assert isinstance(DEPRECATION_POLICY["window_releases"], int)


def test_deprecation_policy_removal_date_none():
    assert DEPRECATION_POLICY["removal_date_iso"] is None


def test_deprecation_policy_has_telemetry_field():
    assert "telemetry_field" in DEPRECATION_POLICY


# ---------------------------------------------------------------------------
# 5. build_capabilities_v2 — default surface
# ---------------------------------------------------------------------------


def test_v2_primary_tools_count():
    v2 = build_capabilities_v2(None)
    assert len(v2["primary_tools"]) == 45


def test_v2_deprecated_aliases():
    v2 = build_capabilities_v2(None)
    assert v2["deprecated_aliases"] == ["add_content"]


def test_v2_replacement_tool():
    v2 = build_capabilities_v2(None)
    assert v2["replacement_tool"] == {"add_content": "shape"}


def test_v2_replacement_operation_or_scope():
    v2 = build_capabilities_v2(None)
    assert v2["replacement_operation_or_scope"] == {"add_content": ""}


def test_v2_total_callable_endpoints():
    v2 = build_capabilities_v2(None)
    assert v2["total_callable_endpoints"] == 46
    assert v2["total_callable_endpoints"] == len(v2["primary_tools"]) + len(v2["deprecated_aliases"])


def test_v2_primary_tools_sorted():
    v2 = build_capabilities_v2(None)
    assert v2["primary_tools"] == sorted(v2["primary_tools"])


def test_v2_add_content_not_in_primary():
    v2 = build_capabilities_v2(None)
    assert "add_content" not in v2["primary_tools"]


def test_v2_add_content_in_deprecated():
    v2 = build_capabilities_v2(None)
    assert "add_content" in v2["deprecated_aliases"]


def test_v2_com_only_tools_absent():
    v2 = build_capabilities_v2(None)
    assert "run_slide_show" not in v2["primary_tools"]
    assert "recalculate_charts" not in v2["primary_tools"]


# ---------------------------------------------------------------------------
# 6. operation_scope_enums — default surface
# ---------------------------------------------------------------------------


def test_v2_operation_scope_enums_keys():
    v2 = build_capabilities_v2(None)
    ose = v2["operation_scope_enums"]
    assert set(ose.keys()) == {"export", "manage_hyperlinks", "set_format", "shape", "slide"}


def test_v2_operation_scope_enums_no_add_content():
    v2 = build_capabilities_v2(None)
    assert "add_content" not in v2["operation_scope_enums"]


def test_v2_operation_scope_enums_slide():
    v2 = build_capabilities_v2(None)
    assert v2["operation_scope_enums"]["slide"] == sorted(["add", "copy", "delete", "reorder"])


def test_v2_operation_scope_enums_shape():
    v2 = build_capabilities_v2(None)
    assert v2["operation_scope_enums"]["shape"] == sorted([
        "add_autoshape", "add_table", "add_text_box",
        "delete", "set_geometry", "set_properties", "set_table_style",
    ])


def test_v2_operation_scope_enums_export():
    v2 = build_capabilities_v2(None)
    assert v2["operation_scope_enums"]["export"] == sorted(["deck_pdf", "slide", "slide_images", "slide_text"])
    assert "slide_png" not in v2["operation_scope_enums"]["export"]


def test_v2_operation_scope_enums_set_format():
    v2 = build_capabilities_v2(None)
    assert v2["operation_scope_enums"]["set_format"] == sorted(["paragraph", "text"])


def test_v2_operation_scope_enums_manage_hyperlinks():
    v2 = build_capabilities_v2(None)
    assert v2["operation_scope_enums"]["manage_hyperlinks"] == sorted(["add", "list", "remove"])


def test_v2_operation_scope_enums_all_sorted():
    v2 = build_capabilities_v2(None)
    for vals in v2["operation_scope_enums"].values():
        assert vals == sorted(vals)


# ---------------------------------------------------------------------------
# 7. write_gate_metadata — default surface
# ---------------------------------------------------------------------------


def test_v2_write_gate_metadata_count():
    v2 = build_capabilities_v2(None)
    assert len(v2["write_gate_metadata"]) == 27


def test_v2_write_gate_dispatchers():
    v2 = build_capabilities_v2(None)
    wgm = v2["write_gate_metadata"]
    for dispatcher in ["export", "manage_hyperlinks", "set_format", "shape", "slide"]:
        assert dispatcher in wgm
        assert wgm[dispatcher]["env_var"] == "PPT_ENABLE_WRITE"
        assert wgm[dispatcher]["requires_confirm"] is True


def test_v2_write_gate_includes_add_content():
    v2 = build_capabilities_v2(None)
    assert "add_content" in v2["write_gate_metadata"]


def test_v2_write_gate_excludes_readonly():
    v2 = build_capabilities_v2(None)
    wgm = v2["write_gate_metadata"]
    for readonly in ["capabilities", "read_presentation", "list_slides",
                     "read_slide", "detect_overlapping_shapes", "validate_contract"]:
        assert readonly not in wgm


# ---------------------------------------------------------------------------
# 8. deprecation_policy
# ---------------------------------------------------------------------------


def test_v2_deprecation_policy():
    v2 = build_capabilities_v2(None)
    dp = v2["deprecation_policy"]
    assert dp["window_releases"] == 2
    assert dp["telemetry_field"] == "deprecated_alias_called"
    assert dp["removal_date_iso"] is None


# ---------------------------------------------------------------------------
# 9. Compact surface filtering
# ---------------------------------------------------------------------------


def test_v2_compact_surface():
    compact_list = list(COMPACT_TOOL_NAMES)  # 22 tools
    v2c = build_capabilities_v2(None, active_tools=compact_list)
    assert v2c["total_callable_endpoints"] == 22
    assert len(v2c["primary_tools"]) + len(v2c["deprecated_aliases"]) == 22
    # add_content is in compact surface (transition window) — it's a deprecated_alias
    assert "add_content" in v2c["deprecated_aliases"]
    assert "add_content" not in v2c["primary_tools"]
    # manage_hyperlinks is NOT in compact surface
    assert "manage_hyperlinks" not in v2c["primary_tools"]
    assert "manage_hyperlinks" not in v2c.get("operation_scope_enums", {})
    # 4 dispatcher keys (not 5 — manage_hyperlinks excluded)
    assert set(v2c["operation_scope_enums"].keys()) == {"export", "set_format", "shape", "slide"}


# ---------------------------------------------------------------------------
# 10. Registry keys spot checks
# ---------------------------------------------------------------------------


def test_registry_contains_expected_tools():
    for name in [
        "slide", "shape", "export", "set_format", "manage_hyperlinks",
        "add_content", "capabilities", "save", "read_presentation",
        "export_slide", "export_deck",
    ]:
        assert name in TOOL_REGISTRY
