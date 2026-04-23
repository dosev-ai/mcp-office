"""Artifact Context Packet (ACP) adapter for excelmcp.

Pure logic module — no FastMCP / MCP imports.
Called by server_acp.py which provides the @mcp.tool() wiring.

Security rules enforced:
  S-PATH  _check_path() is the FIRST call before any file I/O.
  S-001   ACPDetail MUST NOT contain cell values.  Named ranges are
          name strings only; lineage_summary is structural fields only.
  S-ANN   validate_acp_annotations() called before annotations are stored.
"""
from __future__ import annotations

import logging
from datetime import datetime, timezone
from typing import Literal

from mcpshared._acp_contract import _sanitize_text_field, validate_acp_annotations

from excelmcp._core import (  # noqa: F401  (re-used by callers)
    ExcelMCPError,
    _check_path,
)
from excelmcp._metadata_contract import build_workbook_metadata
from excelmcp._resource import resolve_named_ranges_resource, resolve_tables_resource
from excelmcp.runtime_config import get_effective_policy

logger = logging.getLogger(__name__)

def build_workbook_context(
    path: str,
    level: Literal["index", "focused", "deep"] = "focused",
    annotations: list[dict] | None = None,
) -> dict:
    """Build an Artifact Context Packet dict for an Excel workbook.

    Levels:
      index   — ACPIndex only (identity + summary)
      focused — + ACPContent (sheet_names, sheet_count, table_names)
      deep    — + ACPDetail (metadata_policy, named_ranges, lineage_summary)
                + validated annotations if provided

    Security:
      _check_path() is called first (allowlist + extension + size gate).
      ACPDetail never contains cell values (S-001).
      Annotations are validated via validate_acp_annotations() before inclusion.

    Returns a plain dict (not a TypedDict instance) suitable for FastMCP
    JSON serialisation.
    """
    if level not in ("index", "focused", "deep"):
        raise ValueError(
            f"Invalid level {level!r}. Must be 'index', 'focused', or 'deep'."
        )

    # S-PATH: allowlist + extension check first, before any file I/O
    _check_path(path)

    # Build structural metadata — opens workbook read-only via openpyxl internally
    meta = dict(build_workbook_metadata(path))

    # ── Level 1: ACPIndex (always present) ───────────────────────────────────
    # Sanitise potentially user-controlled string fields (OWASP A03)
    _workbook_name_clean = _sanitize_text_field(meta.get("workbook_name", ""))
    summary = (
        f"Excel workbook: {_workbook_name_clean}, "
        f"{meta['sheet_count']} sheet(s), "
        f"last modified {meta['last_modified'][:10]}"
    )[:120]

    acp: dict = {
        "acp_version": "1.0",
        "artifact_type": "spreadsheet",
        "artifact_id": f"wb:{_workbook_name_clean}",
        "tool_name": "excelmcp",
        "summary": summary,
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }

    # ── Level 2: ACPContent (focused + deep) ─────────────────────────────────
    if level in ("focused", "deep"):
        tables_result = resolve_tables_resource(path)
        table_names = [t["name"] for t in tables_result.get("tables", [])]
        acp["content"] = {
            "sheet_count": meta["sheet_count"],
            "sheet_names": [_sanitize_text_field(s) for s in meta.get("sheet_names", [])],
            "table_names": [_sanitize_text_field(t) for t in table_names],
        }

    # ── Level 3: ACPDetail + annotations (deep only) ─────────────────────────
    if level == "deep":
        nr_result = resolve_named_ranges_resource(path)
        named_range_names = [
            nr["name"] for nr in nr_result.get("named_ranges", [])
        ]

        policy = get_effective_policy()

        # S-001: structural fields only — zero cell content
        # H-002: use _workbook_name_clean (already sanitised) instead of raw meta["workbook_name"]
        lineage_summary = {
            k: meta[k]
            for k in (
                "format",
                "contract_version",
                "has_formulas",
                "last_modified",
                "size_bytes",
            )
            if k in meta
        }
        lineage_summary["workbook_name"] = _workbook_name_clean

        acp["detail"] = {
            "metadata_policy": policy,
            "named_ranges": named_range_names,
            "lineage_summary": lineage_summary,
        }

        if annotations:
            # S-ANN: validate before storing
            errors = validate_acp_annotations(annotations)  # type: ignore[arg-type]
            if errors:
                raise ValueError(f"Invalid ACP annotations: {errors}")
            acp["annotations"] = annotations

    return acp
