"""MCP handler registration for Word document review and evidence tools.

Layer: server (MCP imports allowed, no COM).
"""
from __future__ import annotations

import datetime
import json
import logging
import uuid
from collections.abc import Callable
from datetime import timezone
from pathlib import Path
from typing import Any

from fastmcp.exceptions import ToolError

from wordmcp._server_docx_tools import _coerce_list_param
from wordmcp.document_docx import (
    NotAllowedError,
    ValidationError,
    _check_confirm,
    _check_evidence_path,
    _check_path,
    _check_write,
)

logger = logging.getLogger(__name__)


REVIEW_PROMPT_TEXT = """# Word Document Review Checklist

Use these tools to systematically review a Word document for quality and completeness.

## Step 1: Get document outline
Call `get_document_outline` with the document path. Verify headings are present and
the document structure is clear.

## Step 2: Run structural checks
Call `review_document` with `check_names=null` to run all checks. Review:
- **style_consistency**: all Normal paragraphs use same font
- **heading_hierarchy**: no skipped heading levels (H1 → H3 without H2)
- **word_count**: total word count for the document
- **table_structure**: all tables have consistent column counts

## Step 3: Record findings
If issues are found, call `write_review_findings` with:
- `evidence_path`: path to a `.jsonl` file within your allowlist
- `findings`: list of {check_name, passed, detail, recommendation} dicts
- `document_path`: the document under review
- `confirm: true`

## Step 4: Export evidence bundle
Call `export_review_evidence` to package all results for audit trail or sign-off workflow.

## Iteration
For each failed check, make the required document edits, then re-run `review_document`
to verify the fix. Repeat until all checks pass.
"""


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: Any,
) -> dict[str, Any]:
    """Register all review tools with the MCP server via bind_tool.

    Returns a dict of {tool_name: bound_tool} so server.py can expose them
    as module-level attributes (required for smoke tests).
    """

    def _get_document_outline(path: str) -> dict:
        """Extract headings, paragraph summary, and table count from a .docx file.

        Args:
            path: Absolute path to the .docx file. Must be within WORD_ALLOWLIST_ROOTS.

        Returns:
            dict with keys: path, headings (list of {level, text}),
            paragraph_count, table_count.

        DEPRECATED: Use read_document(path, scope='outline') instead.
        This alias is kept for backward compatibility and returns identical output.
        """
        import wordmcp._docx.review_doc as rd
        return rd.get_document_outline(path)

    def _review_document(path: str, check_names: list[str] | str | None = None) -> dict:
        """Run structural quality checks on a .docx document.

        Args:
            path: Absolute path to the .docx file. Must be within WORD_ALLOWLIST_ROOTS.
            check_names: Optional list of check names to run. If None, all checks run.
                Valid values: style_consistency, heading_hierarchy, word_count,
                table_structure.

        Returns:
            dict with keys: path, checks_run, results ({check_name: {passed, detail}}).
        """
        import wordmcp._docx.review_doc as rd
        if check_names is not None:
            check_names = _coerce_list_param(check_names, "check_names")
        return rd.review_document(path, check_names=check_names)

    def _write_review_findings(
        evidence_path: str,
        findings: list[dict] | str,
        document_path: str,
        review_id: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """Append review findings as a JSONL record to an evidence file.

        Requires WORD_ENABLE_WRITE=true and confirm=True.

        Args:
            evidence_path: Absolute path to a .jsonl file within WORD_ALLOWLIST_ROOTS.
                           Parent directory must already exist.
            findings: List of finding dicts. Each should have at minimum:
                      {check_name, passed, detail}.
            document_path: Path of the document under review (for evidence record).
            review_id: Optional UUID string. Auto-generated if not provided.
            confirm: Must be True to write.

        Returns:
            dict with keys: review_id, evidence_path, findings_count.

        Gates (in order): WORD_ENABLE_WRITE env var, confirm=True, path allowlist
        + .jsonl suffix, parent directory exists.
        """
        # GATE 1: write env var check
        try:
            _check_write()
        except NotAllowedError as exc:
            raise ToolError(f"{type(exc).__name__}: {exc}") from exc

        # GATE 2: confirm
        try:
            _check_confirm(confirm)
        except ValidationError as exc:
            raise ToolError(f"{type(exc).__name__}: {exc}") from exc

        # Coerce findings from JSON string if FastMCP serialised it
        findings = _coerce_list_param(findings, "findings")

        # GATE 3: evidence_path allowlist + .jsonl suffix
        try:
            ep = _check_evidence_path(evidence_path)
        except ValidationError as exc:
            raise ToolError(f"{type(exc).__name__}: {exc}") from exc

        # GATE 4: parent directory must exist — no auto-create
        if not ep.parent.is_dir():
            raise ToolError(
                "evidence_path parent directory does not exist"
            )

        # Build record
        rid = review_id or str(uuid.uuid4())
        record = dict(
            review_id=rid,
            document_path=str(Path(document_path)),
            timestamp=datetime.datetime.now(timezone.utc).isoformat(),
            findings=findings,
        )

        try:
            with ep.open("a", encoding="utf-8") as f:
                f.write(json.dumps(record) + "\n")
        except OSError as exc:
            logger.error("Failed to write evidence file: %s", exc)
            raise ToolError(f"Failed to write evidence file ({type(exc).__name__})") from exc

        logger.info("Review findings written to %s (review_id=%s)", ep, rid)
        return dict(
            review_id=rid,
            evidence_path=str(ep),
            findings_count=len(findings),
        )

    def _export_review_evidence(
        path: str,
        review_results: dict,
        export_result: dict | None = None,
    ) -> dict:
        """Package review results into a structured evidence bundle.

        DEPRECATED: Use export(scope='review') when a dedicated review-export scope
        is added.  This tool is kept as-is for backward compatibility and continues
        to function identically — it does NOT write a file, it returns an evidence
        bundle dict that the caller can pass to write_review_findings or store
        externally.

        Args:
            path: Absolute path to the document reviewed.
            review_results: Output from review_document tool.
            export_result: Optional export metadata (e.g., from export_as_pdf).

        Returns:
            dict with keys: document_path, tool_version, timestamp,
            review_results, export_result.
        """
        checked = _check_path(path)

        try:
            from importlib.metadata import version
            tool_version = version("wordmcp")
        except Exception:
            tool_version = "unknown"

        return dict(
            document_path=str(checked),
            tool_version=tool_version,
            timestamp=datetime.datetime.now(timezone.utc).isoformat(),
            review_results=review_results,
            export_result=export_result,
        )

    return {
        "get_document_outline": bind_tool(_get_document_outline, "get_document_outline"),
        "review_document": bind_tool(_review_document, "review_document"),
        "write_review_findings": bind_tool(_write_review_findings, "write_review_findings"),
        "export_review_evidence": bind_tool(_export_review_evidence, "export_review_evidence"),
    }
