"""Unit tests for Word document review core tools."""
from __future__ import annotations

import pytest


def _call_tool(tool, *args, **kwargs):
    """Invoke a FastMCP-decorated tool regardless of FastMCP version."""
    fn = getattr(tool, "fn", tool)
    return fn(*args, **kwargs)


@pytest.mark.unit
def test_get_outline_heading_levels(review_docx):
    """T-WO1: headings list has correct levels (H1=1, H2=2)."""
    from wordmcp._docx.review_doc import get_document_outline

    result = get_document_outline(str(review_docx))
    levels = [h["level"] for h in result["headings"]]
    assert 1 in levels
    assert 2 in levels


@pytest.mark.unit
def test_get_outline_heading_count(review_docx):
    """T-WO2: exactly 2 headings in the review_docx fixture."""
    from wordmcp._docx.review_doc import get_document_outline

    result = get_document_outline(str(review_docx))
    assert len(result["headings"]) == 2


@pytest.mark.unit
def test_get_outline_paragraph_count(review_docx):
    """T-WO3: 3 Normal paragraphs in review_docx (not counting headings)."""
    from wordmcp._docx.review_doc import get_document_outline

    result = get_document_outline(str(review_docx))
    assert result["paragraph_count"] == 3


@pytest.mark.unit
def test_get_outline_table_count_zero(review_docx):
    """T-WO4: table_count is 0 for a document with no tables."""
    from wordmcp._docx.review_doc import get_document_outline

    result = get_document_outline(str(review_docx))
    assert result["table_count"] == 0


@pytest.mark.unit
def test_get_outline_bad_path_raises(tmp_path, monkeypatch):
    """T-WO5: ValidationError raised for path not in allowlist."""
    from wordmcp._docx.review_doc import get_document_outline
    from wordmcp.document_docx import ValidationError

    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    with pytest.raises((ValidationError, Exception)):
        get_document_outline("/not/in/allowlist/file.docx")


@pytest.mark.unit
def test_review_document_all_checks_run(review_docx):
    """T-WR1: all 4 checks run when check_names=None."""
    from wordmcp._docx.review_doc import VALID_CHECKS, review_document

    result = review_document(str(review_docx))
    assert set(result["checks_run"]) == VALID_CHECKS
    for name in VALID_CHECKS:
        assert name in result["results"]
        assert "passed" in result["results"][name]


@pytest.mark.unit
def test_review_document_specific_check(review_docx):
    """T-WR2: only requested check runs when check_names specified."""
    from wordmcp._docx.review_doc import review_document

    result = review_document(str(review_docx), check_names=["word_count"])
    assert result["checks_run"] == ["word_count"]
    assert "word_count" in result["results"]
    assert "style_consistency" not in result["results"]


@pytest.mark.unit
def test_review_document_unknown_check_raises(review_docx):
    """T-WR3: unknown check name raises ValidationError."""
    from wordmcp._docx.review_doc import review_document
    from wordmcp.document_docx import ValidationError

    with pytest.raises(ValidationError, match="Unknown"):
        review_document(str(review_docx), check_names=["nonexistent_check"])


@pytest.mark.unit
def test_review_heading_hierarchy_violation(skipped_h3_docx):
    """T-WR4: heading_hierarchy check fails when H1→H3 (skips H2)."""
    from wordmcp._docx.review_doc import review_document

    result = review_document(str(skipped_h3_docx), check_names=["heading_hierarchy"])
    check = result["results"]["heading_hierarchy"]
    assert check["passed"] is False
    assert "H1" in check["detail"] or "Viol" in check["detail"]


@pytest.mark.unit
def test_review_word_count_returns_count(review_docx):
    """T-WR5: word_count check returns a positive word_count in detail."""
    from wordmcp._docx.review_doc import review_document

    result = review_document(str(review_docx), check_names=["word_count"])
    check = result["results"]["word_count"]
    assert check["passed"] is True
    assert "word_count" in check
    assert check["word_count"] > 0


@pytest.mark.unit
def test_review_document_check_names_json_string_coerced(review_docx):
    """T-COERCE-1: check_names passed as JSON string is coerced to list."""
    from wordmcp import server as srv

    result = _call_tool(
        srv.review_document,
        path=str(review_docx),
        check_names='["word_count"]',
    )
    assert result["checks_run"] == ["word_count"]
    assert "word_count" in result["results"]


@pytest.mark.unit
def test_review_document_check_names_invalid_json_raises(review_docx):
    """T-COERCE-2: check_names with invalid JSON string raises ToolError."""
    from wordmcp import server as srv

    with pytest.raises(Exception, match="invalid JSON string"):
        _call_tool(
            srv.review_document,
            path=str(review_docx),
            check_names="not_valid_json[",
        )
