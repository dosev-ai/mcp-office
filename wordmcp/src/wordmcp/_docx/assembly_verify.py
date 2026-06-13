"""verify_no_placeholders.

Extracted from assembly.py (structural decomposition, no logic changes).
"""
from __future__ import annotations

import re
from typing import Any

from wordmcp._docx import _facade
from wordmcp._docx.assembly_hyperlinks import _TOKEN_RE

_MAX_IGNORE = 1000


def verify_no_placeholders(
    path: str,
    ignore: list[str] | None = None,
) -> dict[str, Any]:
    """Verify that no unreplaced {{TOKEN}} placeholders remain in a Word document.

    Read-only: no write gate, no confirm gate, no save, no evict, no audit log.

    Scan scope (explicit):
      - doc.paragraphs: body paragraphs only (excludes headers, footers, footnotes)
      - doc.tables: all table cells, via table.rows[*].cells[*].paragraphs
      Headers and footers are explicitly out of scope and not scanned.
    """
    facade = _facade()

    # Gate 1 — path allowlist only (read-only operation)
    resolved = facade._check_path(path)

    # Validate ignore list
    if ignore is not None:
        if len(ignore) > _MAX_IGNORE:
            raise facade.ValidationError(
                f"ignore list too large: {len(ignore)} > {_MAX_IGNORE}"
            )
        for entry in ignore:
            if not _TOKEN_RE.fullmatch(entry):
                raise facade.ValidationError(
                    f"ignore entry {entry!r} must match {{{{...}}}} pattern "
                    f"with no nested braces"
                )

    doc = facade._load_doc(resolved)

    pattern = re.compile(r"\{\{[^{}]+\}\}")
    ignored: set[str] = set(ignore or [])
    residuals: list[str] = []

    # Scan body paragraphs only (not headers/footers — see module docstring)
    scanned_paragraphs = 0
    for para in doc.paragraphs:
        scanned_paragraphs += 1
        for match in pattern.finditer(para.text):
            token = match.group()
            if token not in ignored:
                residuals.append(token)

    # Scan all table cells
    scanned_cells = 0
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                scanned_cells += 1
                for para in cell.paragraphs:
                    for match in pattern.finditer(para.text):
                        token = match.group()
                        if token not in ignored:
                            residuals.append(token)

    status = "pass" if not residuals else "fail"

    return {
        "status": status,
        "residuals": residuals,
        "scanned_paragraphs": scanned_paragraphs,
        "scanned_cells": scanned_cells,
        "path": str(resolved),
    }
