"""Assembly tools: bulk_find_replace, manage_hyperlinks, verify_no_placeholders.

COM layer: None (file-based only, python-docx).
MCP imports: None — this module is the backend; server.py wires the tools.

Scope note for verify_no_placeholders:
  Scans doc.paragraphs (body paragraphs only) and all table cells.
  Headers, footers, text boxes, and footnotes are explicitly excluded — they
  are not reachable via doc.paragraphs or doc.tables in python-docx.

Implementation is split across:
  - assembly_hyperlinks.py  (manage_hyperlinks, _do_hyperlink_replace, _validate_url)
  - assembly_verify.py      (verify_no_placeholders)
  - assembly.py             (bulk_find_replace — stays here for monkeypatch compatibility)
"""
from __future__ import annotations

from typing import Any

from wordmcp._docx import _facade
from wordmcp._docx.assembly_hyperlinks import (
    _ALLOWED_SCHEMES,
    _HYPERLINK_REL_TYPE,
    _TOKEN_RE,
    _do_hyperlink_replace,
    _validate_url,
    manage_hyperlinks,
)
from wordmcp._docx.assembly_verify import verify_no_placeholders
from wordmcp._docx.core import _do_find_replace

_MAX_REPLACEMENTS = 500


# ---------------------------------------------------------------------------
# Tool: bulk_find_replace
# ---------------------------------------------------------------------------


def bulk_find_replace(
    path: str,
    replacements: dict[str, str],
    confirm: bool = False,
    hyperlink_fields: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Replace multiple {{TOKEN}} placeholders in a Word document in a single pass.

    hyperlink_fields maps "{{TOKEN}}" -> "https://..." for tokens that should
    become clickable w:hyperlink XML elements instead of plain text. Tokens
    listed in hyperlink_fields must NOT also appear in replacements.

    Gate order:
      1. _check_write()
      2. _check_confirm(confirm)
      3. _check_path(path)
      4. Validate replacements dict
      4b. Validate hyperlink_fields dict (if provided)
      5. Load document, loop replacements, apply hyperlink_fields, save, evict, audit
    """
    facade = _facade()

    # Gate 1 — write enable (MUST be first)
    facade._check_write()
    # Gate 2 — explicit caller confirmation
    facade._check_confirm(confirm)
    # Gate 3 — path allowlist
    resolved = facade._check_path(path)

    # Gate 4 — validate replacements
    if not replacements:
        raise facade.ValidationError("replacements must not be empty")
    if len(replacements) > _MAX_REPLACEMENTS:
        raise facade.ValidationError(
            f"Too many entries: {len(replacements)} > {_MAX_REPLACEMENTS}"
        )
    for token, value in replacements.items():
        if not _TOKEN_RE.fullmatch(token):
            raise facade.ValidationError(
                f"Invalid token {token!r}: must match {{{{...}}}} pattern "
                f"with no nested braces"
            )
        if value is None or not isinstance(value, str):
            raise facade.ValidationError(
                f"Replacement value for {token!r} must be a str, got {type(value).__name__}"
            )

    # Gate 4b — validate hyperlink_fields (if provided)
    if hyperlink_fields is not None:
        if not isinstance(hyperlink_fields, dict):
            raise facade.ValidationError(
                f"hyperlink_fields must be a dict, got {type(hyperlink_fields).__name__}"
            )
        for token, url in hyperlink_fields.items():
            if not _TOKEN_RE.fullmatch(token):
                raise facade.ValidationError(
                    f"hyperlink_fields token {token!r}: must match {{{{...}}}} pattern "
                    f"with no nested braces"
                )
            _validate_url(url, facade)
        overlap = set(replacements) & set(hyperlink_fields)
        if overlap:
            raise facade.ValidationError(
                f"Tokens appear in both replacements and hyperlink_fields: "
                f"{sorted(overlap)}"
            )

    # Gate 5 — document operations
    doc = facade._load_doc(resolved)

    results: dict[str, int] = {}
    for token, value in replacements.items():
        results[token] = _do_find_replace(doc, token, value)

    if hyperlink_fields:
        part = doc.part
        for token, url in hyperlink_fields.items():
            results[token] = _do_hyperlink_replace(doc, part, token, url)

    tokens_not_found: list[str] = [t for t, c in results.items() if c == 0]
    total_replaced = sum(results.values())

    facade._atomic_save(doc, resolved)
    facade._evict_doc(resolved)
    facade._audit_log(
        "bulk_find_replace",
        resolved,
        extra={"replaced": total_replaced},
    )

    return {
        "replaced": total_replaced,
        "results": results,
        "tokens_not_found": tokens_not_found,
        "path": str(resolved),
    }


__all__ = [
    "bulk_find_replace",
    "manage_hyperlinks",
    "verify_no_placeholders",
    # Internal helpers exposed for backward-compatibility
    "_do_find_replace",
    "_do_hyperlink_replace",
    "_validate_url",
    "_HYPERLINK_REL_TYPE",
    "_TOKEN_RE",
    "_ALLOWED_SCHEMES",
]
