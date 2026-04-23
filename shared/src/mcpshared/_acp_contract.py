"""
Artifact Context Packet (ACP) contract v1 — TypedDict definitions and validator.

ACP levels:
  - ACPIndex    (Level 1) — always-present index metadata
  - ACPFocused  (Level 2) — index + optional content summary
  - ACPDeep     (Level 3) — focused + optional detail + annotations

Security note S-001 (deep-detail):
  ACPDetail intentionally excludes raw cell values to prevent
  leaking sensitive workbook content through the context layer.
  Adapters MUST NOT embed cell values in acp_detail fields.
"""
from __future__ import annotations

import re
from typing import Literal, TypedDict

# ---------- Level 0 primitive ----------


class ACPAnnotation(TypedDict, total=False):
    key: str    # max 64 chars, no leading =+-@, no NUL
    value: str  # max 1024 chars, no NUL
    source: str  # max 64 chars, e.g. "user", "rule-engine"
    rule_id: str  # e.g. "IBCS-1.3", optional


# ---------- Level 2 content ----------


class ACPContent(TypedDict, total=False):
    # PPT fields
    slide_count: int
    slide_titles: list[str]   # one per slide, may be empty string
    # Excel fields
    sheet_count: int
    sheet_names: list[str]
    table_names: list[str]    # Excel Table (ListObject) names


# ---------- Level 3 detail ----------


class ACPDetail(TypedDict, total=False):
    # PPT detail
    png_paths: list[str]           # export artefact paths if rendered
    review_findings: list[dict]    # from produce_evidence_bundle findings list
    # Excel detail — NO CELL VALUES (S-001)
    metadata_policy: str           # "strict" | "lenient"
    named_ranges: list[str]        # named range names only, not values
    lineage_summary: dict          # output of normalize_metadata() summary


# ---------- Level 1: index (always present) ----------


class ACPIndex(TypedDict):
    acp_version: str                                    # always "1.0"
    artifact_type: Literal["presentation", "spreadsheet"]
    artifact_id: str   # e.g. "deck:report.pptx" or "wb:budget.xlsx"
    tool_name: str     # MCP server name: "pptmcp" | "excelmcp"
    summary: str       # one-line human description, max 120 chars
    timestamp: str     # ISO 8601 UTC, e.g. "2025-01-15T14:00:00Z"


# ---------- Level 2: focused ----------


class ACPFocused(ACPIndex, total=False):
    content: ACPContent


# ---------- Level 3: deep ----------


class ACPDeep(ACPFocused, total=False):
    detail: ACPDetail
    annotations: list[ACPAnnotation]


# ---------- Security helpers (OWASP A03 — formula-injection prevention) ----------

_FORMULA_PREFIX = re.compile(r"^[=+\-@\t\r\n]")


def _sanitize_text_field(value: str | None, max_len: int = 255) -> str:
    """Strip formula-injection prefixes and cap length. OWASP A03."""
    if not isinstance(value, str):
        return ""
    if _FORMULA_PREFIX.match(value):
        value = "'" + value[:max_len - 1]  # prefix-escape; len ≤ max_len
    else:
        value = value[:max_len]
    return value


# ---------- Validator ----------

_FORBIDDEN_LEADING = re.compile(r"^[=+\-@\t\r\n]")
_NUL = re.compile(r"\x00")


def validate_acp_annotations(annotations: list[ACPAnnotation]) -> list[str]:
    """Validate annotation list; return list of error strings (empty = valid)."""
    errors: list[str] = []
    for i, ann in enumerate(annotations):
        prefix = f"annotations[{i}]"
        key = ann.get("key", "")
        value = ann.get("value", "")
        source = ann.get("source", "")
        if len(key) > 64:
            errors.append(f"{prefix}.key exceeds 64 chars")
        if _FORBIDDEN_LEADING.match(key):
            errors.append(f"{prefix}.key starts with forbidden char")
        if _NUL.search(key):
            errors.append(f"{prefix}.key contains NUL byte")
        if len(value) > 1024:
            errors.append(f"{prefix}.value exceeds 1024 chars")
        if _FORBIDDEN_LEADING.match(value):
            errors.append(f"{prefix}.value starts with forbidden char")
        if _NUL.search(value):
            errors.append(f"{prefix}.value contains NUL byte")
        if len(source) > 64:
            errors.append(f"{prefix}.source exceeds 64 chars")
        if _NUL.search(source):
            errors.append(f"{prefix}.source contains NUL byte")
        rule_id = ann.get("rule_id", "")
        if len(rule_id) > 64:
            errors.append(f"{prefix}.rule_id exceeds 64 chars")
        if _FORBIDDEN_LEADING.match(rule_id):
            errors.append(f"{prefix}.rule_id starts with forbidden char")
        if _NUL.search(rule_id):
            errors.append(f"{prefix}.rule_id contains NUL byte")
    return errors
