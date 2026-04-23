"""Export tools: export_as_pdf and export_range_as_csv.

Tools registered on shared mcp instance at import time (side-effect import).
Do not instantiate FastMCP here.

Env vars consumed:
    EXCEL_ENABLE_COM      -- gate for all COM tools
    EXCEL_ALLOWLIST_ROOTS -- input workbook path root
    EXCEL_EXPORT_ROOTS    -- PDF output path root
    EXCEL_CSV_EXPORT_ROOTS -- CSV output path root
    EXCEL_ENABLE_WRITE    -- gate for file-write operations
"""
from __future__ import annotations

import hashlib
import logging
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal, cast

from fastmcp.exceptions import ToolError

from excelmcp._com import (
    _check_csv_export_path,
    _check_export_path,
    _close_workbook,
    _com_excel_app,
    _ensure_com_gate,
    _open_workbook,
    _sanitize_csv_value,
)
from excelmcp._core import (
    _RANGE_ADDRESS_RE,
    NotAllowedError,
    OfficeCOMError,
    ValidationError,
    _check_confirm,
    _check_path,
    _check_range_size,
    _check_write,
)
from excelmcp._server_instance import mcp

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    """Return the SHA-256 digest for a file."""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(chunk_size):
            digest.update(chunk)
    return digest.hexdigest()


def _utc_now_iso() -> str:
    """Return the current UTC timestamp as an ISO-8601 string."""
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace(
        "+00:00", "Z"
    )


def _require_absolute_path(path: str, field_name: str) -> None:
    """Reject relative paths at the public tool boundary."""
    if not Path(path).is_absolute():
        raise ValidationError(f"{field_name} must be an absolute path")


def _validate_export_scope(scope: str) -> Literal["sheet", "workbook"]:
    """Validate the runtime export scope explicitly."""
    if scope not in {"sheet", "workbook"}:
        raise ValidationError(
            f"scope must be 'sheet' or 'workbook', got: {scope!r}"
        )
    return cast(Literal["sheet", "workbook"], scope)


def _validate_export_sheet_argument(
    *,
    scope: Literal["sheet", "workbook"],
    sheet: str | None,
) -> None:
    """Validate the scope/sheet contract explicitly at the tool boundary."""
    if scope == "sheet":
        if sheet is None:
            raise ValidationError("sheet is required when scope='sheet'")
        return
    if sheet is not None:
        raise ValidationError("sheet is not allowed when scope='workbook'")


def _assert_pdf_export_target_clearable(output_path: Path) -> None:
    """Validate that an existing PDF target is safe to replace."""
    if not output_path.exists():
        return
    if not output_path.is_file():
        raise ToolError(
            f"PDF export target already exists and is not a file: {output_path}"
        )


def _clear_preexisting_pdf_export_target(output_path: Path) -> None:
    """Clear any pre-existing PDF target so post-export evidence is always fresh."""
    if not output_path.exists():
        return
    try:
        output_path.unlink()
    except OSError as exc:
        raise ToolError(
            f"PDF export failed: could not clear pre-existing target at {output_path}: {exc}"
        ) from exc


def _finalize_pdf_export(
    *,
    output_path: Path,
    started_at: float,
    payload: dict[str, Any],
) -> dict[str, Any]:
    """Verify the exported PDF and append evidence fields."""
    if not output_path.exists():
        raise ToolError(f"PDF export failed: output file was not created at {output_path}")
    if not output_path.is_file():
        raise ToolError(f"PDF export failed: output path is not a file: {output_path}")

    size_bytes = output_path.stat().st_size
    if size_bytes <= 0:
        raise ToolError(f"PDF export failed: output file is empty: {output_path}")

    payload.update(
        {
            "ok": True,
            "sha256": _sha256_file(output_path),
            "size_bytes": size_bytes,
            "exported_at": _utc_now_iso(),
            "elapsed_ms": round((time.monotonic() - started_at) * 1000),
        }
    )
    return payload


# ---------------------------------------------------------------------------
# Tool T5/T6 unified: export_as_pdf
# ---------------------------------------------------------------------------

@mcp.tool()
def export_as_pdf(
    scope: Literal["sheet", "workbook"],
    path: str,
    output_path: str,
    sheet: str | None = None,
    quality: int = 0,
    confirm: bool = False,
) -> dict[str, Any]:
    """Export a worksheet or entire workbook to PDF via Excel COM.

    scope='sheet' exports a single named worksheet (sheet param required).
    scope='workbook' exports all visible sheets in print order.
    The source workbook is opened read-only and is not modified.

    Args:
        scope: 'sheet' to export one worksheet; 'workbook' to export all visible sheets.
        path: Absolute path to the source .xlsx/.xlsm workbook.
              Must be within EXCEL_ALLOWLIST_ROOTS.
        output_path: Absolute path for the output PDF file.
                     Must be within EXCEL_EXPORT_ROOTS and end in .pdf.
        sheet: Worksheet name (required when scope='sheet'; must be omitted when
               scope='workbook').
        quality: 0 = standard quality (default), 1 = minimum quality.
        confirm: Must be True.

    Returns:
        scope='sheet':    {"ok": True, "sheet": str, "output_path": str, "quality": int,
                   "sha256": str, "size_bytes": int, "exported_at": str,
                   "elapsed_ms": int}
        scope='workbook': {"ok": True, "output_path": str, "quality": int, "source": str,
                   "sha256": str, "size_bytes": int, "exported_at": str,
                   "elapsed_ms": int}

    Requires: EXCEL_ENABLE_COM=true, confirm=True.
    Does NOT require EXCEL_ENABLE_WRITE (source workbook is not modified).
    """
    try:
        _ensure_com_gate()
        _check_confirm(confirm)
        scope = _validate_export_scope(scope)
        _require_absolute_path(path, "path")
        resolved = _check_path(path)
        resolved_out = _check_export_path(output_path)
        _assert_pdf_export_target_clearable(resolved_out)
        if quality not in (0, 1):
            raise ValidationError(
                f"quality must be 0 (standard) or 1 (minimum), got: {quality!r}"
            )
        _validate_export_sheet_argument(scope=scope, sheet=sheet)
        if scope == "sheet":
            started_at = time.monotonic()
            with _com_excel_app() as excel:
                wb = _open_workbook(excel, resolved, read_only=True)
                try:
                    sheet_names = [
                        wb.Sheets(i + 1).Name
                        for i in range(wb.Sheets.Count)
                    ]
                    if sheet not in sheet_names:
                        raise ValidationError(
                            f"Sheet {sheet!r} not found in workbook "
                            f"(available: {sheet_names})"
                        )
                    _clear_preexisting_pdf_export_target(resolved_out)
                    ws = wb.Sheets(sheet)
                    ws.ExportAsFixedFormat(
                        Type=0,
                        Filename=str(resolved_out),
                        Quality=quality,
                        IncludeDocProperties=True,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False,
                    )
                finally:
                    _close_workbook(wb, save=False)
                    del wb
            return _finalize_pdf_export(
                output_path=resolved_out,
                started_at=started_at,
                payload={
                    "sheet": sheet,
                    "output_path": str(resolved_out),
                    "quality": quality,
                },
            )
        else:  # scope == "workbook"
            started_at = time.monotonic()
            with _com_excel_app() as excel:
                wb = _open_workbook(excel, resolved, read_only=True)
                try:
                    _clear_preexisting_pdf_export_target(resolved_out)
                    wb.ExportAsFixedFormat(
                        Type=0,
                        Filename=str(resolved_out),
                        Quality=quality,
                        OpenAfterPublish=False,
                    )
                finally:
                    _close_workbook(wb, save=False)
                    del wb
            return _finalize_pdf_export(
                output_path=resolved_out,
                started_at=started_at,
                payload={
                    "output_path": str(resolved_out),
                    "quality": quality,
                    "source": str(resolved),
                },
            )
    except (NotAllowedError, ValidationError):
        raise
    except ToolError:
        raise
    except OfficeCOMError as exc:
        logger.error("[excelmcp] export_as_pdf COM error: %s", exc, exc_info=True)
        raise ToolError(str(exc)) from exc
    except Exception as exc:
        logger.error("[excelmcp] export_as_pdf unexpected error: %s", exc, exc_info=True)
        raise ToolError("An unexpected error occurred -- check server logs for details.") from exc


# ---------------------------------------------------------------------------
# Tool: export_range_as_csv
# ---------------------------------------------------------------------------

@mcp.tool()
def export_range_as_csv(
    path: str,
    sheet: str,
    address: str,
    include_headers: bool = True,
    delimiter: str = ",",
    output_path: str | None = None,
    live: bool = False,
    confirm: bool = False,
) -> dict:
    """Export a cell range as CSV.

    live=False (default): reads cached values via openpyxl -- no COM, CI-safe.
    live=True: reads live-evaluated values via Excel COM (Windows only,
               requires EXCEL_ENABLE_COM=true).
    output_path + any mode: write CSV to file (needs EXCEL_CSV_EXPORT_ROOTS +
               EXCEL_ENABLE_WRITE=true + confirm=True).

    Args:
        path: Workbook path within EXCEL_ALLOWLIST_ROOTS.
        sheet: Worksheet name.
        address: Cell range in A1 notation, e.g. "A1:C10".
        include_headers: If True (default), all rows are written to CSV.
                         If False, the first row (header) is skipped.
        delimiter: CSV delimiter character. One of: , \\t ; | (default: ,).
        output_path: If provided, write CSV to this .csv file.
        live: If True, use Excel COM to read live calculated values.
        confirm: Required (True) when output_path is provided.

    Returns:
        {"path": str, "sheet": str, "address": str, "rows": int, "cols": int,
         "csv": str | None, "output_path": str | None}

    Requires:
        live=True:          EXCEL_ENABLE_COM=true, Windows only.
        output_path set:    EXCEL_CSV_EXPORT_ROOTS, EXCEL_ENABLE_WRITE=true, confirm=True.
    """
    import csv as _csv  # noqa: PLC0415
    import io as _io_mod  # noqa: PLC0415

    try:
        _VALID_DELIMITERS = {",", "\t", ";", "|"}
        if delimiter not in _VALID_DELIMITERS:
            raise ValidationError(
                f"delimiter must be one of: {sorted(_VALID_DELIMITERS)!r}"
            )

        if "!" in address:
            sheet_prefix, _, addr_part = address.partition("!")
            canonical_prefix = sheet_prefix.strip("'")
            if canonical_prefix != sheet:
                raise ValidationError(
                    f"address sheet prefix {sheet_prefix!r} does not match "
                    f"sheet parameter {sheet!r}"
                )
            address = addr_part

        if not _RANGE_ADDRESS_RE.match(address):
            raise ValidationError(
                f"address {address!r} is not a valid A1-notation range "
                f"(examples: 'A1', 'A1:C10', '$A$1:$C$10')"
            )

        checked_path = _check_path(path)

        validated_op: Path | None = None
        if output_path is not None:
            _check_write()
            _check_confirm(confirm)
            validated_op = _check_csv_export_path(output_path)
            if not validated_op.parent.is_dir():
                raise ToolError(
                    f"output_path parent directory does not exist: {validated_op.parent}"
                )

        if live:
            _ensure_com_gate()
            with _com_excel_app() as excel:
                wb_com = _open_workbook(excel, checked_path, read_only=True)
                try:
                    sheet_names = [
                        wb_com.Sheets(i + 1).Name for i in range(wb_com.Sheets.Count)
                    ]
                    if sheet not in sheet_names:
                        raise ValidationError(
                            f"Sheet {sheet!r} not found in workbook "
                            f"(available: {sheet_names})"
                        )
                    ws_com = wb_com.Sheets(sheet)
                    rng = ws_com.Range(address.upper())
                    _check_range_size(rng.Count)
                    raw = rng.Value
                finally:
                    _close_workbook(wb_com, save=False)
                    del wb_com

            if not isinstance(raw, (list, tuple)):
                values: list[list] = [[raw]]
            elif raw and isinstance(raw[0], (list, tuple)):
                values = [[cell for cell in row] for row in raw]
            else:
                values = [[cell for cell in raw]]
        else:
            from excelmcp._io import read_range as _read_range  # noqa: PLC0415

            raw_rows = _read_range(str(checked_path), sheet, address.upper())
            values = [[cell["value"] for cell in row] for row in raw_rows]

        if not include_headers and values:
            values = values[1:]

        buf = _io_mod.StringIO()
        writer = _csv.writer(buf, delimiter=delimiter)
        for row in values:
            writer.writerow([_sanitize_csv_value(cell) for cell in row])
        csv_str = buf.getvalue()

        if output_path is None:
            if len(csv_str.encode("utf-8")) > 5 * 1024 * 1024:
                raise ToolError(
                    "CSV output exceeds 5 MB inline limit. "
                    "Provide output_path to write directly to file."
                )

        written_path: str | None = None
        if output_path is not None:
            assert validated_op is not None
            validated_op.write_text(csv_str, encoding="utf-8")
            written_path = str(validated_op)

        return {
            "path": str(checked_path),
            "sheet": sheet,
            "address": address.upper(),
            "rows": len(values),
            "cols": len(values[0]) if values else 0,
            "csv": csv_str if written_path is None else None,
            "output_path": written_path,
        }

    except (NotAllowedError, ValidationError):
        raise
    except ToolError:
        raise
    except OfficeCOMError as exc:
        logger.error("[excelmcp] export_range_as_csv COM error: %s", exc, exc_info=True)
        raise ToolError(str(exc)) from exc
    except Exception as exc:
        logger.error(
            "[excelmcp] export_range_as_csv unexpected error: %s", exc, exc_info=True
        )
        raise ToolError("CSV export failed -- check server logs for details.") from exc


# ---------------------------------------------------------------------------
# Backward-compat shims -- NOT MCP-registered
# ---------------------------------------------------------------------------

def export_sheet_as_pdf(
    path: str,
    sheet: str,
    output_path: str,
    quality: int = 0,
    confirm: bool = False,
) -> dict:
    """Backward compat. Use export_as_pdf(scope='sheet') in new code."""
    return export_as_pdf(
        scope="sheet",
        path=path,
        output_path=output_path,
        sheet=sheet,
        quality=quality,
        confirm=confirm,
    )


def export_workbook_as_pdf(
    path: str,
    output_path: str,
    quality: int = 0,
    confirm: bool = False,
) -> dict:
    """Backward compat. Use export_as_pdf(scope='workbook') in new code."""
    return export_as_pdf(
        scope="workbook",
        path=path,
        output_path=output_path,
        quality=quality,
        confirm=confirm,
    )
