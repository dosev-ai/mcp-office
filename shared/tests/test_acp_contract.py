"""Unit tests for ACP contract v1 (mcpshared._acp_contract).

≥ 15 tests covering TypedDicts and validate_acp_annotations.
"""

from mcpshared import (
    ACPAnnotation,
    ACPContent,
    ACPDeep,
    ACPDetail,
    ACPFocused,
    ACPIndex,
    validate_acp_annotations,
)
from mcpshared._acp_contract import _sanitize_text_field

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_index(**overrides) -> ACPIndex:
    base: ACPIndex = {
        "acp_version": "1.0",
        "artifact_type": "spreadsheet",
        "artifact_id": "wb:budget.xlsx",
        "tool_name": "excelmcp",
        "summary": "Q1 budget workbook",
        "timestamp": "2025-01-15T14:00:00Z",
    }
    base.update(overrides)  # type: ignore[typeddict-item]
    return base


# ---------------------------------------------------------------------------
# Test 1 — ACPIndex accepts all required fields
# ---------------------------------------------------------------------------


def test_t01_acpindex_required_fields():
    idx = _make_index()
    assert idx["acp_version"] == "1.0"
    assert idx["tool_name"] == "excelmcp"
    assert idx["artifact_id"] == "wb:budget.xlsx"
    assert idx["summary"] == "Q1 budget workbook"
    assert idx["timestamp"] == "2025-01-15T14:00:00Z"


# ---------------------------------------------------------------------------
# Test 2 — ACPFocused inherits ACPIndex and accepts content
# ---------------------------------------------------------------------------


def test_t02_acpfocused_inherits_index_and_accepts_content():
    content: ACPContent = {"sheet_count": 3, "sheet_names": ["Sheet1", "Sheet2", "Sheet3"]}  # noqa: E501
    focused: ACPFocused = {**_make_index(), "content": content}
    assert focused["acp_version"] == "1.0"
    assert focused["content"]["sheet_count"] == 3


# ---------------------------------------------------------------------------
# Test 3 — ACPDeep accepts detail and annotations
# ---------------------------------------------------------------------------


def test_t03_acpdeep_accepts_detail_and_annotations():
    ann: ACPAnnotation = {"key": "reviewed", "value": "yes", "source": "user"}
    detail: ACPDetail = {"metadata_policy": "strict", "named_ranges": ["Revenue"]}
    deep: ACPDeep = {
        **_make_index(),
        "detail": detail,
        "annotations": [ann],
    }
    assert deep["detail"]["metadata_policy"] == "strict"
    assert deep["annotations"][0]["key"] == "reviewed"


# ---------------------------------------------------------------------------
# Test 4a — artifact_type accepts "presentation"
# ---------------------------------------------------------------------------


def test_t04a_artifact_type_presentation():
    idx = _make_index(artifact_type="presentation", artifact_id="deck:report.pptx", tool_name="pptmcp")  # noqa: E501
    assert idx["artifact_type"] == "presentation"


# ---------------------------------------------------------------------------
# Test 4b — artifact_type accepts "spreadsheet"
# ---------------------------------------------------------------------------


def test_t04b_artifact_type_spreadsheet():
    idx = _make_index(artifact_type="spreadsheet")
    assert idx["artifact_type"] == "spreadsheet"


# ---------------------------------------------------------------------------
# Test 5 — ACPContent PPT variant
# ---------------------------------------------------------------------------


def test_t05_acpcontent_ppt_variant():
    content: ACPContent = {
        "slide_count": 5,
        "slide_titles": ["Intro", "Market", "Finance", "Risks", "Appendix"],
    }
    assert content["slide_count"] == 5
    assert len(content["slide_titles"]) == 5


# ---------------------------------------------------------------------------
# Test 6 — ACPContent Excel variant
# ---------------------------------------------------------------------------


def test_t06_acpcontent_excel_variant():
    content: ACPContent = {
        "sheet_count": 2,
        "sheet_names": ["Data", "KPIs"],
        "table_names": ["tbl_Data", "tbl_KPIs"],
    }
    assert content["sheet_count"] == 2
    assert "tbl_Data" in content["table_names"]


# ---------------------------------------------------------------------------
# Test 7 — ACPDetail PPT variant (png_paths + review_findings)
# ---------------------------------------------------------------------------


def test_t07_acpdetail_ppt_variant():
    detail: ACPDetail = {
        "png_paths": ["/tmp/slide_01.png", "/tmp/slide_02.png"],
        "review_findings": [{"slide": 1, "issue": "missing title"}],
    }
    assert len(detail["png_paths"]) == 2
    assert detail["review_findings"][0]["issue"] == "missing title"


# ---------------------------------------------------------------------------
# Test 8 — ACPDetail Excel variant — NO cell values (S-001)
# ---------------------------------------------------------------------------


def test_t08_acpdetail_excel_variant_no_cell_values():
    detail: ACPDetail = {
        "metadata_policy": "strict",
        "named_ranges": ["Revenue", "COGS"],
        "lineage_summary": {"source": "ETL", "version": "2025-01"},
    }
    assert detail["metadata_policy"] == "strict"
    assert "Revenue" in detail["named_ranges"]
    assert detail["lineage_summary"]["source"] == "ETL"
    # Confirm no "cell_values" field exists in the TypedDict definition
    from mcpshared._acp_contract import ACPDetail as _ACPDetail
    assert "cell_values" not in _ACPDetail.__annotations__


# ---------------------------------------------------------------------------
# Test 9 — validate: empty list → no errors
# ---------------------------------------------------------------------------


def test_t09_validate_empty_list_no_errors():
    assert validate_acp_annotations([]) == []


# ---------------------------------------------------------------------------
# Test 10 — validate: key > 64 chars → error containing "key"
# ---------------------------------------------------------------------------


def test_t10_validate_key_too_long():
    ann: ACPAnnotation = {"key": "x" * 65, "value": "ok", "source": "test"}
    errors = validate_acp_annotations([ann])
    assert any("key" in e for e in errors)


# ---------------------------------------------------------------------------
# Test 11 — validate: key starts with "=" → error
# ---------------------------------------------------------------------------


def test_t11_validate_key_starts_with_equals():
    ann: ACPAnnotation = {"key": "=DROP TABLES", "value": "evil", "source": "attacker"}
    errors = validate_acp_annotations([ann])
    assert errors, "Expected at least one error for forbidden leading char '='"
    assert any("forbidden" in e for e in errors)


# ---------------------------------------------------------------------------
# Test 12 — validate: key starts with "-" → error
# ---------------------------------------------------------------------------


def test_t12_validate_key_starts_with_dash():
    ann: ACPAnnotation = {"key": "-bad", "value": "ok", "source": "test"}
    errors = validate_acp_annotations([ann])
    assert any("forbidden" in e for e in errors)


def test_t12b_validate_key_starts_with_plus():
    errs = validate_acp_annotations([{"key": "+injection", "value": "ok", "source": "test"}])  # noqa: E501
    assert any("key" in e for e in errs)


def test_t12c_validate_key_starts_with_at():
    errs = validate_acp_annotations([{"key": "@injection", "value": "ok", "source": "test"}])  # noqa: E501
    assert any("key" in e for e in errs)


# ---------------------------------------------------------------------------
# Test 13 — validate: value > 1024 chars → error
# ---------------------------------------------------------------------------


def test_t13_validate_value_too_long():
    ann: ACPAnnotation = {"key": "note", "value": "v" * 1025, "source": "test"}
    errors = validate_acp_annotations([ann])
    assert any("value" in e for e in errors)


# ---------------------------------------------------------------------------
# Test 14 — validate: NUL byte in key → error
# ---------------------------------------------------------------------------


def test_t14_validate_nul_byte_in_key():
    ann: ACPAnnotation = {"key": "bad\x00key", "value": "ok", "source": "test"}
    errors = validate_acp_annotations([ann])
    assert any("NUL" in e or "nul" in e.lower() for e in errors)


# ---------------------------------------------------------------------------
# Test 15 — validate: source > 64 chars → error
# ---------------------------------------------------------------------------


def test_t15_validate_source_too_long():
    ann: ACPAnnotation = {"key": "note", "value": "ok", "source": "s" * 65}
    errors = validate_acp_annotations([ann])
    assert any("source" in e for e in errors)


# ---------------------------------------------------------------------------
# Test 16 — validate: fully valid annotation list → 0 errors
# ---------------------------------------------------------------------------


def test_t16_validate_fully_valid_annotations():
    annotations: list[ACPAnnotation] = [
        {"key": "reviewed", "value": "yes", "source": "user", "rule_id": "IBCS-1.3"},
        {"key": "status", "value": "approved", "source": "rule-engine"},
    ]
    errors = validate_acp_annotations(annotations)
    assert errors == [], f"Expected no errors, got: {errors}"


# ---------------------------------------------------------------------------
# BLOCKER-1 tests — formula-prefix check on value
# ---------------------------------------------------------------------------


def test_b1_validate_value_starts_with_equals_forbidden():
    """BLOCKER-1: value='=SUM(A1)' must be rejected."""
    ann: ACPAnnotation = {"key": "note", "value": "=SUM(A1)", "source": "test"}
    errors = validate_acp_annotations([ann])
    assert errors, "Expected at least one error for formula-prefix in value"
    assert any("value" in e for e in errors)


def test_b2_validate_value_starts_with_plus_forbidden():
    """BLOCKER-1: value='+1' must be rejected."""
    ann: ACPAnnotation = {"key": "note", "value": "+1", "source": "test"}
    errors = validate_acp_annotations([ann])
    assert errors
    assert any("value" in e for e in errors)


def test_b3_validate_value_starts_with_at_forbidden():
    """BLOCKER-1: value='@cmd' must be rejected."""
    ann: ACPAnnotation = {"key": "note", "value": "@cmd", "source": "test"}
    errors = validate_acp_annotations([ann])
    assert errors
    assert any("value" in e for e in errors)


def test_b8_validate_value_valid_does_not_trigger_prefix_check():
    """BLOCKER-1: normal value does not trigger prefix check."""
    ann: ACPAnnotation = {"key": "note", "value": "normal text", "source": "test"}
    errors = validate_acp_annotations([ann])
    assert errors == [], f"Expected no errors for normal value, got: {errors}"


# ---------------------------------------------------------------------------
# BLOCKER-2 tests — rule_id validation
# ---------------------------------------------------------------------------


def test_b4_validate_rule_id_too_long():
    """BLOCKER-2: rule_id > 64 chars must be rejected."""
    ann: ACPAnnotation = {"key": "note", "value": "ok", "source": "test", "rule_id": "R" * 65}  # noqa: E501
    errors = validate_acp_annotations([ann])
    assert errors, "Expected error for rule_id exceeding 64 chars"
    assert any("rule_id" in e for e in errors)


def test_b5_validate_rule_id_starts_with_equals():
    """BLOCKER-2: rule_id='=formula' must be rejected."""
    ann: ACPAnnotation = {"key": "note", "value": "ok", "source": "test", "rule_id": "=formula"}  # noqa: E501
    errors = validate_acp_annotations([ann])
    assert errors
    assert any("rule_id" in e for e in errors)


def test_b6_validate_rule_id_nul_byte():
    """BLOCKER-2: rule_id with \\x00 must be rejected."""
    ann: ACPAnnotation = {"key": "note", "value": "ok", "source": "test", "rule_id": "bad\x00id"}  # noqa: E501
    errors = validate_acp_annotations([ann])
    assert errors
    assert any("rule_id" in e for e in errors)


def test_b7_validate_fully_valid_with_rule_id():
    """BLOCKER-2: all fields valid including rule_id → 0 errors."""
    ann: ACPAnnotation = {
        "key": "checked",
        "value": "passed",
        "source": "rule-engine",
        "rule_id": "IBCS-2.1",
    }
    errors = validate_acp_annotations([ann])
    assert errors == [], f"Expected no errors, got: {errors}"


# ---------------------------------------------------------------------------
# F-001 / F-004 tests — sanitizer max_len bound and None guard
# ---------------------------------------------------------------------------


def test_sanitize_max_len_prefix_escape_stays_within_bound():
    """F-001: '='-prefixed max_len-char input must return exactly max_len chars."""
    dangerous = "=" + "x" * 254  # 255 chars total, starts with '='
    result = _sanitize_text_field(dangerous, max_len=255)
    assert len(result) == 255, (
        f"Expected exactly 255 chars after sanitisation, got {len(result)}: {result!r}"
    )
    assert result.startswith("'"), "Expected apostrophe prefix-escape"


def test_sanitize_none_input_returns_empty():
    """F-004: _sanitize_text_field(None) returns '' — not TypeError."""
    result = _sanitize_text_field(None)  # type: ignore[arg-type]
    assert result == "", f"Expected '' for None input, got {result!r}"


def test_sanitize_plain_text_truncated_at_max_len():
    """GAP-1: plain text longer than max_len must be silently truncated."""
    long_text = "a" * 300
    result = _sanitize_text_field(long_text, max_len=255)
    assert len(result) == 255, (
        f"Expected exactly 255 chars after truncation, got {len(result)}"
    )
    assert result == "a" * 255
    assert not result.startswith("'"), "Normal text must NOT gain an apostrophe prefix"


def test_sanitize_empty_string_returns_empty():
    """GAP-2: empty string input must return empty string without error."""
    result = _sanitize_text_field("")
    assert result == "", f"Expected '' for empty-string input, got {result!r}"


def test_validate_nul_byte_in_value():
    """GAP-3 (security): NUL byte in annotation value must be rejected."""
    ann: ACPAnnotation = {"key": "note", "value": "bad\x00value", "source": "test"}
    errors = validate_acp_annotations([ann])
    assert errors, "Expected at least one error for NUL byte in value"
    assert any("value" in e for e in errors), (
        f"Error message must mention 'value', got: {errors}"
    )


def test_validate_nul_byte_in_source():
    """GAP-4 (security): NUL byte in annotation source must be rejected."""
    ann: ACPAnnotation = {"key": "note", "value": "ok", "source": "bad\x00src"}
    errors = validate_acp_annotations([ann])
    assert errors, "Expected at least one error for NUL byte in source"
    assert any("source" in e for e in errors), (
        f"Error message must mention 'source', got: {errors}"
    )

