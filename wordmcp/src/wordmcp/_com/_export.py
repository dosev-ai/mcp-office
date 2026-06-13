"""
Export COM operations: PDF and HTML document export.

No FastMCP/MCP imports.
"""
from __future__ import annotations

import contextlib
import os
import pathlib

from wordmcp._com._base import (
    OfficeCOMError,
    ValidationError,
    _check_com_enabled,
    _check_path,
    _check_write,
    _check_confirm,
    _is_com_error,
    _com_err_msg,
    _word_app_context,
)


def _check_output_path_html(raw_path: str) -> pathlib.Path:
    """Validate an export output path for HTML.

    Ordering (mirrors _check_output_path for PDF):
      1. Null-byte check
      2. UNC/network path check  (before any Path() call)
      3. Allowlist roots present
      4. Path().resolve()
      5. Extension check (.html only)
      6. Allowlist containment check
    """
    # Step 1: null-byte check
    if "\x00" in raw_path:
        raise ValidationError("Null byte in output path")
    # Step 2: UNC path check — BEFORE any Path() call
    if raw_path.startswith(("\\\\", "//")):
        raise ValidationError("UNC/network paths are not allowed")
    # Step 3: allowlist roots present
    roots_env = os.environ.get("WORD_ALLOWLIST_ROOTS", "").strip()
    if not roots_env:
        raise ValidationError("No allowlist configured: WORD_ALLOWLIST_ROOTS is unset")
    # Step 4: resolve
    resolved = pathlib.Path(raw_path).resolve()
    # Step 5: extension check
    if resolved.suffix.lower() != ".html":
        raise ValidationError("Output path must have .html extension")
    # Step 6: allowlist containment
    roots = [r.strip() for r in roots_env.split(",") if r.strip()]
    for root in roots:
        try:
            if resolved.is_relative_to(pathlib.Path(root).resolve()):
                return resolved
        except (ValueError, TypeError):
            continue
    raise ValidationError("Output path is not within the configured allowlist")


def _check_output_path(raw_path: str) -> pathlib.Path:
    """Validate an export output path for PDF.

    Ordering (per BLOCKER-2 spec):
      1. Null-byte check
      2. UNC/network path check  (B-2 fix — before any Path() call)
      3. Allowlist roots present
      4. Path().resolve()
      5. Extension check (.pdf only)
      6. Allowlist containment check
    """
    # Step 1: null-byte check
    if "\x00" in raw_path:
        raise ValidationError("Null byte in output path")
    # Step 2: UNC path check — BEFORE any Path() call (B-2 fix)
    if raw_path.startswith(("\\\\", "//")):
        raise ValidationError("UNC/network paths are not allowed")
    # Step 3: allowlist roots present
    roots_env = os.environ.get("WORD_ALLOWLIST_ROOTS", "").strip()
    if not roots_env:
        raise ValidationError("No allowlist configured: WORD_ALLOWLIST_ROOTS is unset")
    # Step 4: resolve
    resolved = pathlib.Path(raw_path).resolve()
    # Step 5: extension check
    if resolved.suffix.lower() != ".pdf":
        raise ValidationError("Output path must have .pdf extension")
    # Step 6: allowlist containment
    roots = [r.strip() for r in roots_env.split(",") if r.strip()]
    for root in roots:
        try:
            if resolved.is_relative_to(pathlib.Path(root).resolve()):
                return resolved
        except (ValueError, TypeError):
            continue
    raise ValidationError("Output path is not within the configured allowlist")


def export_as_pdf(path: str, output_path: str, confirm: bool = False) -> dict:
    """Export a .docx file to PDF via Word COM.

    Gate order: _check_com_enabled → _check_write → _check_confirm
                → _check_path → _check_output_path → COM open → export.

    Args:
        path:        Absolute path to the source .docx file.
        output_path: Absolute path for the output .pdf file.
        confirm:     Must be True to authorise the write operation.

    Returns:
        {"status": "ok", "output_path": str}
    """
    _check_com_enabled()
    _check_write()
    _check_confirm(confirm)
    checked_path = _check_path(path)
    resolved_out = _check_output_path(output_path)
    doc = None
    with _word_app_context() as app:
        try:
            doc = app.Documents.Open(str(checked_path), ReadOnly=True)
            # ExportAsFixedFormat args: OutputFileName, ExportFormat=17 (PDF),
            # OpenAfterExport=False, OptimizeFor=0, Range=0 (wdExportAllDocument),
            # From=1, To=1, Item=0, IncludeDocProperties=False
            doc.ExportAsFixedFormat(str(resolved_out), 17, False, 0, 0, 1, 1, 0, False)
            return {"status": "ok", "output_path": str(resolved_out)}
        except OfficeCOMError:
            raise
        except Exception as exc:
            if _is_com_error(exc):
                msg = _com_err_msg(exc).lower()
                if any(k in msg for k in ("access", "sharing", "used by another")):
                    raise OfficeCOMError(
                        f"PDF export failed: target file is locked by another application. "
                        f"Close it and retry. (detail: {_com_err_msg(exc)})"
                    ) from exc
                raise OfficeCOMError(f"Export failed: {_com_err_msg(exc)}") from exc
            raise
        finally:
            if doc is not None:
                with contextlib.suppress(Exception):
                    doc.Close(False)


def export_document(
    path: str,
    output_path: str,
    format: str = "pdf",  # noqa: A002
    confirm: bool = False,
) -> dict:
    """Route document export to the appropriate COM function.

    Formats:
      'pdf'  → export_as_pdf(path, output_path, confirm)
      'html' — export as filtered HTML via COM SaveAs2 FileFormat=10
      unknown → raises ValidationError
    """
    if format not in ("pdf", "html"):
        raise ValidationError(
            f"Unknown format: {format!r}. Valid: pdf, html"
        )
    _check_com_enabled()
    _check_write()
    _check_confirm(confirm)
    if format == "pdf":
        return export_as_pdf(path, output_path, confirm)
    # format == "html"
    checked_path = _check_path(path)
    resolved_out = _check_output_path_html(output_path)
    doc = None
    with _word_app_context() as app:
        try:
            doc = app.Documents.Open(str(checked_path))
            # wdFormatFilteredHTML = 10
            doc.SaveAs2(FileName=str(resolved_out), FileFormat=10)
            return {"status": "ok", "output_path": str(resolved_out)}
        except OfficeCOMError:
            raise
        except Exception as exc:
            if _is_com_error(exc):
                raise OfficeCOMError(f"HTML export failed: {_com_err_msg(exc)}") from exc
            raise
        finally:
            if doc is not None:
                with contextlib.suppress(Exception):
                    doc.Close(False)
