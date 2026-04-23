"""Export path validation helpers for excelmcp COM infrastructure.

No MCP imports. No win32com imports at module scope.

Public API:
    _check_export_path()      -- validates PDF output path against EXCEL_EXPORT_ROOTS
    _check_csv_export_path()  -- validates CSV output path against EXCEL_CSV_EXPORT_ROOTS
    _sanitize_csv_value()     -- escapes CSV-injection formula prefixes (OWASP)
"""
from __future__ import annotations

import os
from pathlib import Path

from excelmcp._core import ValidationError

_CSV_FORMULA_PREFIXES: frozenset[str] = frozenset({"=", "+", "-", "@"})


def _check_export_path(path: str) -> Path:
    """Validate a PDF export output path against EXCEL_EXPORT_ROOTS.

    Extension must be .pdf. Path must resolve inside at least one root in
    the EXCEL_EXPORT_ROOTS comma-separated env var.

    Returns the resolved Path on success.
    Raises ValidationError on any violation.
    """
    if "\x00" in str(path):
        raise ValidationError("Path contains illegal null byte")

    if not Path(path).is_absolute():
        raise ValidationError(
            "PDF output path must be an absolute path, got relative: "
            + repr(str(path))
        )

    roots_env = os.getenv("EXCEL_EXPORT_ROOTS", "").strip()
    if not roots_env:
        raise ValidationError(
            "EXCEL_EXPORT_ROOTS is not configured — set it to allow PDF export"
        )

    resolved = Path(path).resolve()

    ext = resolved.suffix.lower()
    if ext != ".pdf":
        raise ValidationError(
            f"PDF export output must have .pdf extension, got: {ext!r}"
        )

    roots = [r.strip() for r in roots_env.split(",") if r.strip()]
    for root in roots:
        try:
            if resolved.is_relative_to(Path(root).resolve()):
                return resolved
        except (ValueError, TypeError):
            continue

    raise ValidationError(
        "PDF output path is not within any configured EXCEL_EXPORT_ROOTS directory"
    )


def _check_csv_export_path(path: str) -> Path:
    """Validate a CSV output path against EXCEL_CSV_EXPORT_ROOTS.

    Extension must be .csv. Path must resolve inside at least one root in
    the EXCEL_CSV_EXPORT_ROOTS semicolon-separated env var.

    Returns the resolved Path on success.
    Raises ValidationError on any violation.
    """
    if "\x00" in str(path):
        raise ValidationError("Path contains illegal null byte")

    if not Path(path).is_absolute():
        raise ValidationError("output_path must be an absolute path")

    resolved = Path(path).resolve()

    if resolved.suffix.lower() != ".csv":
        raise ValidationError(
            f"output_path must have .csv suffix, got: {resolved.suffix!r}"
        )

    roots_env = os.environ.get("EXCEL_CSV_EXPORT_ROOTS", "").strip()
    if not roots_env:
        raise ValidationError(
            "No CSV export allowlist configured: EXCEL_CSV_EXPORT_ROOTS is unset"
        )

    roots = [Path(r.strip()).resolve() for r in roots_env.split(";") if r.strip()]
    for root in roots:
        try:
            if resolved.is_relative_to(root):
                return resolved
        except (ValueError, TypeError):
            continue

    raise ValidationError(
        f"output_path not within any configured EXCEL_CSV_EXPORT_ROOTS: {resolved}"
    )


def _sanitize_csv_value(value: object) -> str:
    """Escape CSV-injection formula prefixes per OWASP CSV injection guidance.

    Prefixes =, +, -, @ are dangerous in spreadsheet tools that auto-evaluate
    cell values. This function prepends a single quote (') to neutralise them.

    Args:
        value: Any cell value. None -> "".

    Returns:
        A safe string representation of the value.
    """
    if value is None:
        return ""
    s = str(value)
    if s and s[0] in _CSV_FORMULA_PREFIXES:
        return "'" + s
    return s
