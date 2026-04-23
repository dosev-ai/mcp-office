"""Tests for excelmcp._lineage — 11 tests."""
from __future__ import annotations

import ast
import inspect

import pytest

from excelmcp._lineage import (
    FIELD_ALIASES,
    map_lineage_fields,
    normalize_metadata,
)


def _canonical() -> dict:
    return {
        "path": "/tmp/test.xlsx",
        "workbook_name": "test.xlsx",
        "sheet_names": ["Sheet1"],
        "active_sheet": "Sheet1",
        "sheet_count": 1,
        "has_formulas": False,
        "last_modified": "2026-03-11T00:00:00+00:00",
        "size_bytes": 1024,
        "format": ".xlsx",
        "contract_version": "1.0",
    }


def test_map_lineage_identity():
    """A canonical dict passes through map_lineage_fields unchanged."""
    raw = _canonical()
    result = map_lineage_fields(raw)
    for field in raw:
        assert result[field] == raw[field]


def test_map_lineage_alias_name():
    """Legacy key 'name' is normalized to 'workbook_name'."""
    raw = _canonical()
    del raw["workbook_name"]
    raw["name"] = "test.xlsx"
    result = map_lineage_fields(raw)
    assert result["workbook_name"] == "test.xlsx"
    assert "name" not in result


def test_map_lineage_alias_sheets():
    """Legacy key 'sheets' is normalized to 'sheet_names'."""
    raw = _canonical()
    del raw["sheet_names"]
    raw["sheets"] = ["Sheet1"]
    result = map_lineage_fields(raw)
    assert result["sheet_names"] == ["Sheet1"]


def test_map_lineage_alias_active():
    """Legacy key 'active' is normalized to 'active_sheet'."""
    raw = _canonical()
    del raw["active_sheet"]
    raw["active"] = "Sheet1"
    result = map_lineage_fields(raw)
    assert result["active_sheet"] == "Sheet1"


def test_map_lineage_alias_formula():
    """Legacy key 'formula' is normalized to 'has_formulas'."""
    raw = _canonical()
    del raw["has_formulas"]
    raw["formula"] = True
    result = map_lineage_fields(raw)
    assert result["has_formulas"] is True


def test_map_lineage_unknown_key_ignored():
    """Unrecognised extra keys are silently dropped."""
    raw = _canonical()
    raw["unknown_garbage_key"] = "should be dropped"
    raw["_internal"] = 99
    result = map_lineage_fields(raw)
    assert "unknown_garbage_key" not in result
    assert "_internal" not in result


def test_map_lineage_missing_required_raises():
    """If a required field is missing after normalization, raises ValueError."""
    raw = _canonical()
    del raw["path"]
    with pytest.raises(ValueError, match="path"):
        map_lineage_fields(raw)


def test_normalize_metadata_smoke():
    """normalize_metadata is a public alias for map_lineage_fields."""
    raw = _canonical()
    result = normalize_metadata(raw)
    assert result["contract_version"] == "1.0"


def test_field_aliases_constant_type():
    """FIELD_ALIASES is a dict mapping str -> str."""
    assert isinstance(FIELD_ALIASES, dict)
    assert all(isinstance(k, str) and isinstance(v, str) for k, v in FIELD_ALIASES.items())
    assert FIELD_ALIASES["name"] == "workbook_name"
    assert FIELD_ALIASES["sheets"] == "sheet_names"


def test_no_server_import_in_lineage_module():
    """_lineage must not import from excelmcp.server*, fastmcp.server, or _com."""
    import excelmcp._lineage as mod
    source = inspect.getsource(mod)
    tree = ast.parse(source)
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module:
            assert not node.module.startswith("excelmcp.server"), (
                f"_lineage must not import from server modules; found: {node.module}"
            )
            assert node.module != "fastmcp.server", (
                "_lineage must not import fastmcp.server"
            )
            assert "_com" not in node.module, (
                "_lineage must not import COM modules"
            )


@pytest.mark.parametrize("alias,canonical,value", [
    ("formulas",  "has_formulas",     True),
    ("modified",  "last_modified",    "2026-03-11T00:00:00+00:00"),
    ("bytes",     "size_bytes",       2048),
    ("ext",       "format",           ".xlsm"),
    ("version",   "contract_version", "1.0"),
])
def test_map_lineage_remaining_aliases(alias: str, canonical: str, value):
    """Each remaining FIELD_ALIASES entry normalizes correctly."""
    raw = _canonical()
    del raw[canonical]
    raw[alias] = value
    result = map_lineage_fields(raw)
    assert result[canonical] == value, (
        f"Alias {alias!r} \u2192 {canonical!r}: expected {value!r}, got {result[canonical]!r}"
    )
    assert alias not in result, f"Alias key {alias!r} should be removed after normalization"
