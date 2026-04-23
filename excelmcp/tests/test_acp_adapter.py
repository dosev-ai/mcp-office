"""Unit tests for excelmcp ACP adapter (build_workbook_context).

Covers security invariants, ACP level contracts, and field correctness.
All external I/O is mocked — no real workbook files needed.
"""
from __future__ import annotations

from unittest.mock import patch

import pytest
from excelmcp._acp_adapter import build_workbook_context
from excelmcp._core import ValidationError

# ---------------------------------------------------------------------------
# Shared fake data
# ---------------------------------------------------------------------------

_FAKE_META: dict = {
    "workbook_name": "test.xlsx",
    "sheet_count": 2,
    "sheet_names": ["Sheet1", "Sheet2"],
    "last_modified": "2026-01-01T00:00:00Z",
    "format": "xlsx",
    "contract_version": "1.0",
    "has_formulas": False,
    "size_bytes": 1024,
}

_PATH = "/allowed/test.xlsx"

# patch targets (the module where names are *used*)
_CP = "excelmcp._acp_adapter._check_path"
_BM = "excelmcp._acp_adapter.build_workbook_metadata"
_RT = "excelmcp._acp_adapter.resolve_tables_resource"
_RN = "excelmcp._acp_adapter.resolve_named_ranges_resource"
_GP = "excelmcp._acp_adapter.get_effective_policy"
_VA = "excelmcp._acp_adapter.validate_acp_annotations"


def _call(level: str = "focused", annotations=None) -> dict:
    """Call build_workbook_context with all I/O stubbed."""
    with (
        patch(_CP),
        patch(_BM, return_value=_FAKE_META),
        patch(_RT, return_value={"tables": []}),
        patch(_RN, return_value={"named_ranges": []}),
        patch(_GP, return_value="permissive"),
        patch(_VA, return_value=[]),
    ):
        return build_workbook_context(_PATH, level, annotations)


# ---------------------------------------------------------------------------
# L-003 tests
# ---------------------------------------------------------------------------

@pytest.mark.unit
def test_level_index_returns_required_fields():
    result = _call("index")
    for key in ("acp_version", "artifact_type", "tool_name", "artifact_id", "summary", "timestamp"):  # noqa: E501
        assert key in result, f"Missing required ACPIndex field: {key!r}"


@pytest.mark.unit
def test_level_focused_includes_content():
    result = _call("focused")
    assert "content" in result
    content = result["content"]
    assert "sheet_names" in content or "sheet_count" in content


@pytest.mark.unit
def test_level_deep_includes_detail():
    result = _call("deep")
    assert "detail" in result


@pytest.mark.unit
def test_path_check_called_before_file_io():
    """_check_path() must be the first call — ValidationError must propagate."""
    with (
        patch(_CP, side_effect=ValidationError("not allowed")),
        patch(_BM, side_effect=AssertionError("build_workbook_metadata must not be called")),  # noqa: E501
    ):
        with pytest.raises(ValidationError, match="not allowed"):
            build_workbook_context(_PATH, "focused")


@pytest.mark.unit
def test_invalid_annotations_rejected():
    """Annotations with injection-style values must raise ValueError."""
    bad_annotations = [{"key": "=injection"}]
    with (
        patch(_CP),
        patch(_BM, return_value=_FAKE_META),
        patch(_RT, return_value={"tables": []}),
        patch(_RN, return_value={"named_ranges": []}),
        patch(_GP, return_value="permissive"),
        patch(_VA, return_value=["injection detected"]),  # validator flags it
    ):
        with pytest.raises(ValueError, match="Invalid ACP annotations"):
            build_workbook_context(_PATH, "deep", bad_annotations)


@pytest.mark.unit
def test_artifact_type_is_spreadsheet():
    result = _call("index")
    assert result["artifact_type"] == "spreadsheet"


@pytest.mark.unit
def test_acp_version_is_1_0():
    result = _call("index")
    assert result["acp_version"] == "1.0"


@pytest.mark.unit
def test_no_cell_values_in_deep_detail():
    """S-001: ACPDetail must never expose cell values."""
    result = _call("deep")
    assert "cell_values" not in result.get("detail", {})


@pytest.mark.unit
def test_workbook_name_formula_prefix_sanitized():
    """BLOCKER-4: workbook_name starting with '=' must be sanitised (OWASP A03)."""
    dangerous_meta = {**_FAKE_META, "workbook_name": "=DANGEROUS(A1)"}
    with (
        patch(_CP),
        patch(_BM, return_value=dangerous_meta),
        patch(_RT, return_value={"tables": []}),
        patch(_RN, return_value={"named_ranges": []}),
        patch(_GP, return_value="permissive"),
        patch(_VA, return_value=[]),
    ):
        result = build_workbook_context(_PATH, "index")

    summary = result.get("summary", "")
    artifact_id = result.get("artifact_id", "")

    # The workbook name in summary must NOT appear unescaped — "workbook: =" is injection  # noqa: E501
    assert "workbook: =" not in summary, (
        f"Unsanitised formula-prefix in summary (name starts with '='): {summary!r}"
    )
    # artifact_id must NOT be "wb:=..." — that would be raw injection
    assert not artifact_id.startswith("wb:="), (
        f"Unsanitised formula-prefix in artifact_id: {artifact_id!r}"
    )
    # The Excel-style apostrophe escape must appear (sanitised name starts with "'=")
    assert "'=" in summary or "'=" in artifact_id, (
        "Expected apostrophe prefix-escape (\"'=\") in summary or artifact_id "
        f"after sanitisation. summary={summary!r}, artifact_id={artifact_id!r}"
    )


@pytest.mark.unit
def test_sheet_names_formula_prefix_sanitized():
    """H-001: sheet_names with formula-injection prefixes must be prefix-escaped."""
    meta_with_bad_sheets = {
        **_FAKE_META,
        "sheet_names": ["=Sheet1", "+Sheet2", "NormalSheet"],
    }
    with (
        patch(_CP),
        patch(_BM, return_value=meta_with_bad_sheets),
        patch(_RT, return_value={"tables": []}),
        patch(_RN, return_value={"named_ranges": []}),
        patch(_GP, return_value="permissive"),
        patch(_VA, return_value=[]),
    ):
        result = build_workbook_context(_PATH, "focused")

    sheet_names = result["content"]["sheet_names"]
    assert sheet_names[0] == "'=Sheet1", (
        f"Expected '=Sheet1 but got {sheet_names[0]!r}"
    )
    assert sheet_names[1] == "'+Sheet2", (
        f"Expected '+Sheet2 but got {sheet_names[1]!r}"
    )
    assert sheet_names[2] == "NormalSheet", (
        f"Normal sheet name should be unchanged but got {sheet_names[2]!r}"
    )


@pytest.mark.unit
def test_lineage_workbook_name_uses_clean_value():
    """H-002: lineage_summary[workbook_name] must use sanitised value, not raw meta."""
    meta_with_dangerous_name = {**_FAKE_META, "workbook_name": "=DangerousName"}
    with (
        patch(_CP),
        patch(_BM, return_value=meta_with_dangerous_name),
        patch(_RT, return_value={"tables": []}),
        patch(_RN, return_value={"named_ranges": []}),
        patch(_GP, return_value="permissive"),
        patch(_VA, return_value=[]),
    ):
        result = build_workbook_context(_PATH, "deep")

    lineage_name = result["detail"]["lineage_summary"]["workbook_name"]
    assert lineage_name != "=DangerousName", (
        f"Raw dangerous workbook_name leaked into lineage_summary: {lineage_name!r}"
    )
    assert lineage_name == "'=DangerousName", (
        "Expected prefix-escaped value ""'=DangerousName"" but got {lineage_name!r}"
    )


@pytest.mark.unit
def test_table_names_formula_prefix_sanitized():
    """F-003: table_names with formula-injection prefixes must be prefix-escaped."""
    with (
        patch(_CP),
        patch(_BM, return_value=_FAKE_META),
        patch(_RT, return_value={"tables": [{"name": "=DangerousTable"}]}),
        patch(_RN, return_value={"named_ranges": []}),
        patch(_GP, return_value="permissive"),
        patch(_VA, return_value=[]),
    ):
        result = build_workbook_context(_PATH, "focused")

    table_names = result["content"]["table_names"]
    assert "=DangerousTable" not in table_names, (
        f"Unsanitised formula-prefix table name found: {table_names!r}"
    )
    assert any(n.startswith("'") for n in table_names), (
        f"Expected an apostrophe-escaped table name but got: {table_names!r}"
    )


@pytest.mark.unit
def test_invalid_level_raises_value_error():
    """GAP-5: invalid level string must raise ValueError before any I/O."""
    with pytest.raises(ValueError, match="Invalid level"):
        build_workbook_context(_PATH, "invalid_level")  # type: ignore[arg-type]
