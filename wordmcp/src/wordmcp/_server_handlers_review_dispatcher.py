"""
Review dispatcher handler for wordmcp.

Provides a single `review` tool that routes to existing individual review
tools based on the `operation` parameter.  All individual tools remain
registered (with deprecation notes) — this dispatcher is an additive
convenience wrapper.
"""
from __future__ import annotations

import os
from collections.abc import Callable
from typing import Any

from wordmcp import _server_com_tools, _server_docx_tools
from wordmcp._server_runtime import ServerRuntime
from wordmcp.document_docx import (
    NotAllowedError,
    ValidationError,
    _check_evidence_path,
    _check_path,
)

_VALID_OPERATIONS = (
    "review",
    "write_findings",
    "export_evidence",
    "manage_comments",
    "manage_tracked_changes",
)


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _review(
        operation: str,
        # --- review / export_evidence / manage_comments / manage_tracked_changes ---
        path: str | None = None,
        # --- review op ---
        check_names: list[str] | str | None = None,
        # --- write_findings op ---
        evidence_path: str | None = None,
        findings: list[dict] | str | None = None,
        document_path: str | None = None,
        review_id: str | None = None,
        # --- export_evidence op ---
        review_results: dict | None = None,
        export_result: dict | None = None,
        # --- manage_comments op ---
        comment_operation: str | None = None,
        text: str | None = None,
        author: str | None = None,
        paragraph_index: int | None = None,
        # --- manage_tracked_changes op ---
        tracked_changes_operation: str | None = None,
        revision_index: int | None = None,
        # --- mutation gate (forwarded to write sub-operations) ---
        confirm: bool = False,
    ) -> Any:
        """Unified review dispatcher — routes to the appropriate review tool.

        operation (required): one of review | write_findings | export_evidence |
            manage_comments | manage_tracked_changes.

        Per-operation parameters:

          review              — path (str, required — absolute path to .docx),
                                check_names (list[str] | None, optional — subset of
                                checks to run; None runs all: style_consistency,
                                heading_hierarchy, word_count, table_structure).
                                Read-only — no gate required.

          write_findings      — evidence_path (str, required — absolute path to .jsonl
                                file within WORD_ALLOWLIST_ROOTS), findings (list[dict],
                                required — each dict should have at minimum check_name,
                                passed, detail), document_path (str, required — path of
                                document under review), review_id (str, optional — UUID,
                                auto-generated if omitted).
                                Requires WORD_ENABLE_WRITE=true and confirm=True.

          export_evidence     — path (str, required — absolute path to reviewed document),
                                review_results (dict, required — output from review op),
                                export_result (dict, optional — e.g., from export tool).
                                Read-only — returns a structured evidence bundle dict.

          manage_comments     — path (str, required), comment_operation (str, required —
                                one of list | add | resolve | delete), text (str, optional
                                — required for add), author (str, optional — defaults to
                                'MCP' for add), paragraph_index (int, optional — comment
                                anchor for add; comment_id for resolve/delete).
                                Read actions (list) require no gate; write actions (add,
                                resolve, delete) require WORD_ENABLE_WRITE=true and
                                confirm=True — the underlying tool enforces these gates.

          manage_tracked_changes — path (str, required), tracked_changes_operation (str,
                                required — one of list | accept_all | reject_all),
                                revision_index (int, optional), confirm (bool).
                                Read actions (list) require no gate; write actions require
                                WORD_ENABLE_WRITE=true and confirm=True — the underlying
                                tool enforces these gates. Requires COM (Word desktop).

        Write gate summary:
          write_findings: WORD_ENABLE_WRITE checked here before delegation.
          export_evidence: read-only, no gate.
          review: read-only, no gate.
          manage_comments / manage_tracked_changes: confirm passed through; underlying
            tools apply their own gates for write sub-operations.

        Unknown operation returns a structured error dict:
          {"error": "unknown_operation", "operation": op, "valid_operations": [...]}

        Deprecated aliases (still registered on the server module):
          review_document        — deprecated, use review(operation="review", ...)
          write_review_findings  — deprecated, use review(operation="write_findings", ...)
          export_review_evidence — deprecated, use review(operation="export_evidence", ...)
          manage_comments        — deprecated, use review(operation="manage_comments", ...)
          manage_tracked_changes — deprecated, use review(operation="manage_tracked_changes", ...)
        """
        from fastmcp.exceptions import ToolError

        op = operation.strip().lower()

        # ------------------------------------------------------------------ #
        # review — read-only structural quality checks                        #
        # ------------------------------------------------------------------ #
        if op == "review":
            if path is None:
                raise ToolError("path is required for operation='review'")
            import wordmcp._docx.review_doc as rd
            from wordmcp._server_docx_tools import _coerce_list_param
            cn = check_names
            if cn is not None:
                cn = _coerce_list_param(cn, "check_names")
            return rd.review_document(path, check_names=cn)

        # ------------------------------------------------------------------ #
        # write_findings — write-gated JSONL evidence writer                  #
        # ------------------------------------------------------------------ #
        if op == "write_findings":
            # Gate 1: env var — checked first, before any param validation
            if os.getenv("WORD_ENABLE_WRITE", "false").lower() != "true":
                raise ToolError(
                    "NotAllowedError: WORD_ENABLE_WRITE is not set to 'true'. "
                    "Set WORD_ENABLE_WRITE=true to enable review write operations."
                )
            # Gate 2: confirm
            if not confirm:
                raise ToolError(
                    "ValidationError: confirm=True is required for review write operations. "
                    "Pass confirm=True to proceed."
                )
            if evidence_path is None:
                raise ToolError("evidence_path is required for operation='write_findings'")
            if findings is None:
                raise ToolError("findings is required for operation='write_findings'")
            if document_path is None:
                raise ToolError("document_path is required for operation='write_findings'")
            # Delegate to the backend implementation (same logic as _write_review_findings)
            import datetime
            import json as _json
            import uuid
            from datetime import timezone
            from pathlib import Path

            from wordmcp._server_docx_tools import _coerce_list_param
            findings_list = _coerce_list_param(findings, "findings")
            try:
                ep = _check_evidence_path(evidence_path)
            except ValidationError as exc:
                raise ToolError(f"{type(exc).__name__}: {exc}") from exc
            if not ep.parent.is_dir():
                raise ToolError("evidence_path parent directory does not exist")
            rid = review_id or str(uuid.uuid4())
            record = dict(
                review_id=rid,
                document_path=str(Path(document_path)),
                timestamp=datetime.datetime.now(timezone.utc).isoformat(),
                findings=findings_list,
            )
            try:
                with ep.open("a", encoding="utf-8") as f:
                    f.write(_json.dumps(record) + "\n")
            except OSError as exc:
                raise ToolError(
                    f"Failed to write evidence file ({type(exc).__name__})"
                ) from exc
            return dict(
                review_id=rid,
                evidence_path=str(ep),
                findings_count=len(findings_list),
            )

        # ------------------------------------------------------------------ #
        # export_evidence — read-only evidence bundle builder                 #
        # ------------------------------------------------------------------ #
        if op == "export_evidence":
            if path is None:
                raise ToolError("path is required for operation='export_evidence'")
            if review_results is None:
                raise ToolError("review_results is required for operation='export_evidence'")
            import datetime
            from datetime import timezone
            try:
                checked = _check_path(path)
            except (ValidationError, NotAllowedError) as exc:
                raise ToolError(f"{type(exc).__name__}: {exc}") from exc
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

        # ------------------------------------------------------------------ #
        # manage_comments — mixed gate; underlying tool handles write gates   #
        # ------------------------------------------------------------------ #
        if op == "manage_comments":
            if path is None:
                raise ToolError("path is required for operation='manage_comments'")
            if comment_operation is None:
                raise ToolError(
                    "comment_operation is required for operation='manage_comments' "
                    "(one of: list | add | resolve | delete)"
                )
            return _server_docx_tools.manage_comments(
                runtime,
                path=path,
                operation=comment_operation,
                text=text,
                author=author,
                paragraph_index=paragraph_index,
                confirm=confirm,
            )

        # ------------------------------------------------------------------ #
        # manage_tracked_changes — mixed gate; underlying tool handles gates  #
        # ------------------------------------------------------------------ #
        if op == "manage_tracked_changes":
            if path is None:
                raise ToolError("path is required for operation='manage_tracked_changes'")
            if tracked_changes_operation is None:
                raise ToolError(
                    "tracked_changes_operation is required for operation='manage_tracked_changes' "
                    "(one of: list | accept_all | reject_all)"
                )
            _VALID_TRACKED = frozenset({"list", "accept_all", "reject_all"})
            if tracked_changes_operation not in _VALID_TRACKED:
                raise ToolError(
                    f"Unknown tracked_changes_operation: {tracked_changes_operation!r}. "
                    f"Valid: {sorted(_VALID_TRACKED)}"
                )
            return _server_com_tools.manage_tracked_changes(
                runtime,
                path=path,
                operation=tracked_changes_operation,
                revision_index=revision_index,
                confirm=confirm,
            )

        # Unknown operation — return structured error (no exception, caller can inspect)
        return {
            "error": "unknown_operation",
            "operation": operation,
            "valid_operations": list(_VALID_OPERATIONS),
        }

    return {
        "review": bind_tool(_review, "review"),
    }
