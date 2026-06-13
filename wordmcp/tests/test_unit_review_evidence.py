"""Unit tests for Word review evidence and export tools."""
from __future__ import annotations

import json
from pathlib import Path

import pytest


def _call_tool(tool, *args, **kwargs):
    """Invoke a FastMCP-decorated tool regardless of FastMCP version."""
    fn = getattr(tool, "fn", tool)
    return fn(*args, **kwargs)


@pytest.mark.unit
def test_check_evidence_path_valid(tmp_path, monkeypatch):
    """T-CEP1: valid .jsonl path within allowlist returns resolved Path."""
    from wordmcp._docx.core import _check_evidence_path

    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    evidence = tmp_path / "results.jsonl"
    result = _check_evidence_path(str(evidence))
    assert result == evidence.resolve()


@pytest.mark.unit
def test_check_evidence_path_wrong_suffix(tmp_path, monkeypatch):
    """T-CEP2: path without .jsonl suffix raises ValidationError."""
    from wordmcp._docx.core import _check_evidence_path
    from wordmcp.document_docx import ValidationError

    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    with pytest.raises(ValidationError, match=".jsonl"):
        _check_evidence_path(str(tmp_path / "results.txt"))


@pytest.mark.unit
def test_check_evidence_path_null_byte(tmp_path, monkeypatch):
    """T-CEP3: path with null byte raises ValidationError."""
    from wordmcp._docx.core import _check_evidence_path
    from wordmcp.document_docx import ValidationError

    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    with pytest.raises(ValidationError, match="Null byte"):
        _check_evidence_path("evil\x00.jsonl")


@pytest.mark.unit
def test_check_evidence_path_outside_allowlist(tmp_path, monkeypatch):
    """T-CEP4: path outside allowlist roots raises ValidationError."""
    from wordmcp._docx.core import _check_evidence_path
    from wordmcp.document_docx import ValidationError

    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    outside = Path("C:/outside_allowlist/evidence.jsonl").resolve()
    with pytest.raises(ValidationError, match="allowed roots"):
        _check_evidence_path(str(outside))


@pytest.mark.unit
def test_check_evidence_path_traversal_blocked(tmp_path, monkeypatch):
    """T-CEP5: path with ../ that resolves outside tmp_path raises ValidationError."""
    from wordmcp._docx.core import _check_evidence_path
    from wordmcp.document_docx import ValidationError

    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    traversal_path = tmp_path / ".." / "traversal_evidence.jsonl"
    with pytest.raises(ValidationError):
        _check_evidence_path(str(traversal_path))


@pytest.mark.unit
def test_check_evidence_path_with_no_allowlist_configured(tmp_path):
    """CBR-001: WORD_ALLOWLIST_ROOTS unset → ValidationError with 'No allowlist configured'."""
    import os

    from wordmcp._docx.core import _check_evidence_path
    from wordmcp.document_docx import ValidationError

    old = os.environ.pop("WORD_ALLOWLIST_ROOTS", None)
    try:
        with pytest.raises(ValidationError, match="No allowlist configured"):
            _check_evidence_path(str(tmp_path / "out.jsonl"))
    finally:
        if old is not None:
            os.environ["WORD_ALLOWLIST_ROOTS"] = old


@pytest.mark.unit
def test_wf1_no_write_env_raises_before_confirm(tmp_path, monkeypatch):
    """T-WF1: WORD_ENABLE_WRITE not set → ToolError (write gate fires first)."""
    from fastmcp.exceptions import ToolError
    from wordmcp import server as srv

    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.delenv("WORD_ENABLE_WRITE", raising=False)
    with pytest.raises(ToolError, match="WORD_ENABLE_WRITE"):
        _call_tool(
            srv.write_review_findings,
            evidence_path=str(tmp_path / "ev.jsonl"),
            findings=[],
            document_path="dummy.docx",
            confirm=True,
        )


@pytest.mark.unit
def test_wf2_write_env_set_but_confirm_false_raises(tmp_path, monkeypatch):
    """T-WF2: WORD_ENABLE_WRITE=true but confirm=False → ToolError (confirm gate)."""
    from fastmcp.exceptions import ToolError
    from wordmcp import server as srv

    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    with pytest.raises(ToolError, match="confirm"):
        _call_tool(
            srv.write_review_findings,
            evidence_path=str(tmp_path / "ev.jsonl"),
            findings=[],
            document_path="dummy.docx",
            confirm=False,
        )


@pytest.mark.unit
def test_wf3_wrong_suffix_raises(tmp_path, monkeypatch):
    """T-WF3: env+confirm set but evidence_path has wrong suffix → ToolError."""
    from fastmcp.exceptions import ToolError
    from wordmcp import server as srv

    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    with pytest.raises(ToolError, match=".jsonl"):
        _call_tool(
            srv.write_review_findings,
            evidence_path=str(tmp_path / "ev.txt"),
            findings=[],
            document_path="dummy.docx",
            confirm=True,
        )


@pytest.mark.unit
def test_wf4_outside_allowlist_raises(tmp_path, monkeypatch):
    """T-WF4: env+confirm set but evidence_path outside allowlist → ToolError."""
    from fastmcp.exceptions import ToolError
    from wordmcp import server as srv

    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    outside_path = tmp_path.parent / "outside_evidence.jsonl"
    with pytest.raises(ToolError, match="allowed roots"):
        _call_tool(
            srv.write_review_findings,
            evidence_path=str(outside_path),
            findings=[],
            document_path="dummy.docx",
            confirm=True,
        )


@pytest.mark.unit
def test_wf5_parent_dir_missing_raises(tmp_path, monkeypatch):
    """T-WF5: all gates pass but evidence parent dir missing → ToolError."""
    from fastmcp.exceptions import ToolError
    from wordmcp import server as srv

    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    missing_parent = tmp_path / "nonexistent_dir" / "evidence.jsonl"
    with pytest.raises(ToolError, match="parent directory"):
        _call_tool(
            srv.write_review_findings,
            evidence_path=str(missing_parent),
            findings=[],
            document_path="dummy.docx",
            confirm=True,
        )


@pytest.mark.unit
def test_wf6_success_writes_jsonl(write_review_env):
    """T-WF6: all gates pass → success, JSONL record appended to file."""
    from wordmcp import server as srv

    tmp_path = write_review_env
    evidence_path = tmp_path / "findings.jsonl"
    findings = [
        {"check_name": "word_count", "passed": True, "detail": "Total word count: 42"}
    ]

    result = _call_tool(
        srv.write_review_findings,
        evidence_path=str(evidence_path),
        findings=findings,
        document_path="report.docx",
        review_id="test-review-001",
        confirm=True,
    )

    assert result["review_id"] == "test-review-001"
    assert result["findings_count"] == 1
    assert evidence_path.exists()

    records = [json.loads(line) for line in evidence_path.read_text().splitlines()]
    assert len(records) == 1
    assert records[0]["review_id"] == "test-review-001"
    assert records[0]["findings"][0]["check_name"] == "word_count"


@pytest.mark.unit
def test_wf7_auto_generates_uuid_when_no_review_id(write_review_env):
    """T-WF7: review_id=None → UUID auto-generated, present in return dict."""
    from wordmcp import server as srv

    tmp_path = write_review_env
    evidence_path = tmp_path / "auto_uuid.jsonl"

    result = _call_tool(
        srv.write_review_findings,
        evidence_path=str(evidence_path),
        findings=[],
        document_path="report.docx",
        confirm=True,
    )

    rid = result["review_id"]
    assert rid is not None and len(rid) > 0
    import re
    assert re.match(
        r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", rid
    ), f"review_id does not look like a UUID: {rid!r}"


@pytest.mark.unit
def test_export_review_evidence_returns_expected_keys(review_docx, monkeypatch):
    """export_review_evidence returns dict with required keys."""
    from wordmcp import server as srv

    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(review_docx.parent))

    review_results = {"checks_run": ["word_count"], "results": {"word_count": {"passed": True}}}
    result = _call_tool(
        srv.export_review_evidence,
        path=str(review_docx),
        review_results=review_results,
    )

    assert "document_path" in result
    assert "tool_version" in result
    assert "timestamp" in result
    assert "review_results" in result
    assert "export_result" in result
    assert result["review_results"] == review_results
    assert result["export_result"] is None


@pytest.mark.unit
def test_export_review_evidence_rejects_invalid_path(tmp_path, monkeypatch):
    """export_review_evidence raises error for path outside allowlist."""
    from wordmcp import server as srv

    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    outside = Path("C:/outside_allowlist/doc.docx").resolve()

    with pytest.raises(Exception, match="allowed roots|not.*allow"):
        _call_tool(
            srv.export_review_evidence,
            path=str(outside),
            review_results={"checks_run": [], "results": {}},
        )


@pytest.mark.unit
def test_write_review_findings_findings_json_string_coerced(
    tmp_path, write_review_env
):
    """T-COERCE-3: findings passed as JSON string is coerced to list and written."""
    from wordmcp import server as srv

    ev_path = write_review_env / "evidence.jsonl"
    findings_str = json.dumps([{"check_name": "word_count", "passed": True, "detail": "ok"}])
    result = _call_tool(
        srv.write_review_findings,
        evidence_path=str(ev_path),
        findings=findings_str,
        document_path=str(tmp_path / "doc.docx"),
        confirm=True,
    )
    assert result["findings_count"] == 1
    assert ev_path.exists()


@pytest.mark.unit
def test_write_review_findings_findings_invalid_json_raises(
    tmp_path, write_review_env
):
    """T-COERCE-4: findings with invalid JSON string raises ToolError."""
    from wordmcp import server as srv

    ev_path = write_review_env / "evidence.jsonl"
    with pytest.raises(Exception, match="invalid JSON string"):
        _call_tool(
            srv.write_review_findings,
            evidence_path=str(ev_path),
            findings="[not valid json",
            document_path=str(tmp_path / "doc.docx"),
            confirm=True,
        )
