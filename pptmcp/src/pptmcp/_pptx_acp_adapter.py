"""Artifact Context Packet (ACP) adapter for pptmcp.

Pure logic module — no FastMCP / MCP imports.
Called by _server_handlers_acp.py which provides the @mcp.tool() wiring.

Security rules enforced:
  S-PATH  _check_path() is the FIRST call before any file I/O.
  S-ANN   validate_acp_annotations() called before annotations are stored.
  S-WG    Read-only — no write gate (PPT_ENABLE_WRITE) required.
"""
from __future__ import annotations

import logging
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Literal

from mcpshared._acp_contract import _sanitize_text_field, validate_acp_annotations

from pptmcp.presentation_pptx import (
    PPTMCPError,  # noqa: F401  (re-used by callers)
    ValidationError,  # noqa: F401
    _check_path,
    _load_prs,
    _slide_title,
)

logger = logging.getLogger(__name__)

# Artifact-ID prefix for PowerPoint decks (GAP-7).
# Must match ARTIFACT_ID_PREFIX_DECK in workflowruntime.contracts.constants.
# Not imported from there to avoid cross-package dependency in production source.
_ARTIFACT_ID_PREFIX_DECK = "deck:"


def build_presentation_context(
    path: str,
    level: Literal["index", "focused", "deep"] = "focused",
    annotations: list[dict] | None = None,
) -> dict:
    """Build an Artifact Context Packet dict for a PowerPoint presentation.

    Levels:
      index   — ACPIndex only (identity + summary)
      focused — + ACPContent (slide_count, slide_titles)
      deep    — + ACPDetail (png_paths if export dir exists, review_findings=[])
                + validated annotations if provided

    Security:
      _check_path() is called first (allowlist + extension + size gate).
      Annotations are validated via validate_acp_annotations() before inclusion.

    Returns a plain dict (not a TypedDict instance) suitable for FastMCP
    JSON serialisation.

    Note: detail.review_findings is intentionally empty — callers populate it
    by running produce_evidence_bundle() then calling this tool with annotations
    to attach findings as ACPAnnotations.
    """
    if level not in ("index", "focused", "deep"):
        raise ValueError(
            f"Invalid level {level!r}. Must be 'index', 'focused', or 'deep'."
        )

    # S-PATH: allowlist + extension check first, before any file I/O
    resolved = _check_path(path)
    filename = resolved.name
    filename_clean = _sanitize_text_field(filename)
    stem = resolved.stem

    # Open presentation read-only (python-pptx, no write gate needed)
    prs = _load_prs(resolved)
    slide_count = len(prs.slides)

    stat = resolved.stat()
    last_modified = datetime.fromtimestamp(
        stat.st_mtime, tz=timezone.utc
    ).strftime("%Y-%m-%d")

    # ── Level 1: ACPIndex (always present) ───────────────────────────────────
    summary = (
        f"PowerPoint presentation: {filename_clean}, "
        f"{slide_count} slide(s), "
        f"last modified {last_modified}"
    )[:120]

    acp: dict = {
        "acp_version": "1.0",
        "artifact_type": "presentation",
        "artifact_id": f"{_ARTIFACT_ID_PREFIX_DECK}{filename_clean}",
        "tool_name": "pptmcp",
        "summary": summary,
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }

    # ── Level 2: ACPContent (focused + deep) ─────────────────────────────────
    if level in ("focused", "deep"):
        titles: list[str] = []
        for slide in prs.slides:
            t = _slide_title(slide)
            titles.append(_sanitize_text_field(t) if t is not None else "")
        acp["content"] = {
            "slide_count": slide_count,
            "slide_titles": titles,
        }

    # ── Level 3: ACPDetail + annotations (deep only) ─────────────────────────
    if level == "deep":
        # Scan allowlist roots for a {stem}_exports directory with rendered PNGs
        png_paths: list[str] = []
        roots_env = os.getenv("PPT_ALLOWLIST_ROOTS", "").strip()
        if roots_env:
            for root in (r.strip() for r in roots_env.split(",") if r.strip()):
                candidate = Path(root) / f"{stem}_exports"
                if candidate.is_dir():
                    export_dir_resolved = candidate.resolve()
                    png_paths = sorted(
                        str(p) for p in candidate.glob("*.png")
                        if p.resolve().is_relative_to(export_dir_resolved)
                    )
                    break

        acp["detail"] = {
            "png_paths": png_paths,
            # review_findings is intentionally empty here; callers populate
            # it from produce_evidence_bundle() results as needed.
            "review_findings": [],
        }

        if annotations:
            # S-ANN: validate before storing
            errors = validate_acp_annotations(annotations)  # type: ignore[arg-type]
            if errors:
                raise ValueError(f"Invalid ACP annotations: {errors}")
            acp["annotations"] = annotations

    return acp
