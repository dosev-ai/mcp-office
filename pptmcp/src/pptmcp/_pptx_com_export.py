"""
COM export functions for pptmcp.

Covers:
  ppt_export_slide, export_slide_as_png, ppt_export_deck, export_deck_as_pdf,
  _make_stamped_export_dir, export_slides_to_stamped_dir,
  export_changed_slides_only.

Private module — not part of the public API.
"""
from __future__ import annotations

import contextlib
import hashlib
import logging
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from pptmcp._pptx_com_guards import _check_com_enabled, _ppt_app_context
from pptmcp._pptx_com_validators import (
    _VALID_DECK_FORMATS,
    _VALID_SLIDE_FORMATS,
    _check_in_allowlist,
    _check_output_path_deck,
    _check_output_path_slide,
    _validate_dpi,
    _validate_export_format,
)
from pptmcp.presentation_pptx import (
    OfficeCOMError,
    ValidationError,
    _audit_log,
    _check_confirm,
    _check_path,
    _check_write,
)

__all__: list[str] = []

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# New unified export functions: ppt_export_slide / ppt_export_deck
# ---------------------------------------------------------------------------


def ppt_export_slide(
    file_path: str,
    slide_index: int,
    fmt: str,
    output_path: str,
    dpi: "int | None" = None,
    width_px: int = 0,
    height_px: int = 0,
    confirm: bool = False,
) -> dict:
    """Export a single slide as PNG, PDF, or SVG via PowerPoint COM.

    Slide indexing is 0-based (slide_index=0 is the first slide).

    Gate order:
      1. _check_path(file_path)            — path security (platform-independent)
      2. _validate_export_format             — format security (platform-independent)
      3. _check_com_enabled()
      4. _check_write()
      5. _check_confirm(confirm)
      6. _check_output_path_slide — ext + allowlist
      7. Existence check (B-5)
      8. _validate_dpi for PNG (M-3); warn-and-ignore for PDF/SVG
      9. slide_index bounds check (0-based)

    Args:
        file_path:    Absolute path to the source .pptx file.
        slide_index:  0-based slide index to export.
        fmt:          Export format: 'png', 'pdf', or 'svg' (case-insensitive).
        output_path:  Absolute path for the output file.
        dpi:          Image resolution 72-600 inclusive. Only used for PNG.
                      None = PowerPoint default (96 DPI).
        confirm:      Must be True to authorise the write operation.

    Returns:
        {"exported": True, "slide_number": int, "format": str, "output_path": str}

    Raises:
        NotAllowedError:  COM disabled, write gate off, or confirm=False.
        ValidationError:  Path outside allowlist, bad format, DPI out of range,
                          slide_index out of range, output exists, bad extension.
        OfficeCOMError:   COM call failed (incl. SVG on Office 2016).

    Blockers: B-3, B-5, B-7 (pdf), B-8, M-1, M-2, M-3.
    """
    # ── Gate chain ───────────────────────────────────────────────────────
    # Input security checks run BEFORE COM gate (platform-independent)
    resolved = _check_path(file_path)

    # M-2: normalise format to lower-case BEFORE validation
    fmt = fmt.lower() if isinstance(fmt, str) else fmt
    fmt = _validate_export_format(fmt, _VALID_SLIDE_FORMATS)

    _check_com_enabled()
    _check_write()
    _check_confirm(confirm)

    resolved_output = _check_output_path_slide(output_path, fmt)

    # B-5: refuse to overwrite an existing output file
    if resolved_output.exists():
        raise ValidationError(
            f"Output file already exists: {resolved_output.name}. "
            "Choose a different output_path or remove the existing file first."
        )

    # M-3: validate DPI for PNG; warn-and-ignore for PDF/SVG
    if fmt == "png":
        dpi = _validate_dpi(dpi)
    elif dpi is not None:
        logger.warning(
            "dpi parameter has no effect for %s format; ignoring", fmt
        )
        dpi = None

    # Pixel-dimension validation (US-09)
    if not isinstance(width_px, int) or isinstance(width_px, bool):
        raise ValidationError("width_px must be an integer")
    if not isinstance(height_px, int) or isinstance(height_px, bool):
        raise ValidationError("height_px must be an integer")
    if not (0 <= width_px <= 7680):
        raise ValidationError("width_px must be in range 0\u20137680")
    if not (0 <= height_px <= 7680):
        raise ValidationError("height_px must be in range 0\u20137680")
    if (width_px == 0) != (height_px == 0):
        raise ValidationError(
            "width_px and height_px must both be set or both be 0"
        )
    if width_px > 0 and height_px > 0:
        if fmt != "png":
            raise ValidationError(
                "width_px/height_px are only valid when fmt='png'"
            )
        if dpi is not None:
            logger.warning(
                "Both dpi and width_px/height_px supplied; "
                "dpi=%d will be ignored (width_px/height_px take precedence)",
                dpi,
            )
            dpi = None

    resolved_output.parent.mkdir(parents=True, exist_ok=True)

    import pythoncom

    _audit_log("ppt_export_slide", resolved, slide_idx=slide_index,
               extra={"fmt": fmt, "output_path": str(resolved_output)})

    with _ppt_app_context(visible=False) as app:
        try:
            prs_com = app.Presentations.Open(
                str(resolved), ReadOnly=True, WithWindow=False
            )
        except pythoncom.com_error as exc:
            logger.error("COM open failed for %s: %s", resolved.name, exc)
            raise OfficeCOMError(
                f"Failed to open presentation: {resolved.name}. "
                "Check PowerPoint installation."
            ) from exc

        slide_count = prs_com.Slides.Count

        # 0-based bounds check
        if slide_index < 0 or slide_index >= slide_count:
            with contextlib.suppress(Exception):
                prs_com.Close()
            raise ValidationError(
                f"slide_index {slide_index} is out of range "
                f"(presentation has {slide_count} slide(s); 0-based indexing)"
            )

        try:
            if fmt == "png":
                slide = prs_com.Slides(slide_index + 1)
                if width_px > 0 and height_px > 0:
                    slide.Export(str(resolved_output), "PNG", width_px, height_px)
                elif dpi is not None:
                    scale_width = int(dpi * 10.0)
                    scale_height = int(dpi * 7.5)
                    slide.Export(str(resolved_output), "PNG", scale_width, scale_height)
                else:
                    slide.Export(str(resolved_output), "PNG")

            elif fmt == "pdf":
                # Single-slide PDF using ppPrintSlideRange(4) — all-positional
                # to match the working deck PDF path. pywin32 cannot reliably
                # mix positional + keyword args around a skipped IDispatch slot
                # like PrintRange, which previously raised a SlideStart keyword
                # error. Passing None at pos 7 maps to
                # VT_EMPTY / VBA "Nothing".
                # B-7: IncludeDocProperties=0 — no metadata leakage.
                prs_com.ExportAsFixedFormat(
                    str(resolved_output),  # pos  0: Path
                    2,                     # pos  1: ppFixedFormatTypePDF
                    2,                     # pos  2: Intent: ppFixedFormatIntentPrint
                    0,                     # pos  3: FrameSlides: msoFalse
                    2,                     # pos  4: HandoutOrder
                    1,                     # pos  5: OutputType: ppPrintOutputSlides
                    0,                     # pos  6: PrintHiddenSlides: msoFalse
                    None,                  # pos  7: PrintRange: Nothing
                    4,                     # pos  8: RangeType: ppPrintSlideRange
                    slide_index + 1,       # pos  9: SlideStart
                    slide_index + 1,       # pos 10: SlideEnd
                    0,                     # pos 11: OpenAfterExport
                    0,                     # pos 12: IncludeDocProperties (B-7)
                )

            elif fmt == "svg":
                # SVG requires Office 2019+ / Microsoft 365 (Build 13328+)
                # Office 2016 raises com_error on "SVG" FilterName
                try:
                    prs_com.Slides(slide_index + 1).Export(str(resolved_output), "SVG")
                except pythoncom.com_error as exc:
                    logger.error(
                        "SVG export failed for %s slide %d (possibly older Office): %s",
                        resolved.name, slide_index, exc,
                    )
                    raise OfficeCOMError(
                        "SVG export requires PowerPoint 2019 or Microsoft 365 "
                        "(Build 13328+). Upgrade Office or use format='png' instead."
                    ) from exc

        except (ValidationError, OfficeCOMError):
            raise
        except pythoncom.com_error as exc:
            logger.error("ppt_export_slide COM error for %s: %s",
                         resolved.name, exc)
            raise OfficeCOMError("ppt_export_slide failed") from exc  # B-8
        finally:
            with contextlib.suppress(Exception):
                prs_com.Close()

    return {
        "exported": True,
        "slide_number": slide_index + 1,
        "format": fmt,
        "output_path": str(resolved_output),
    }


def ppt_export_deck(
    file_path: str,
    fmt: str,
    output_path: str,
    dpi: "int | None" = None,
    confirm: bool = False,
) -> dict:
    """Export an entire presentation deck as PDF, PPTX copy, or PNG images.

    Gate order:
      1. _check_path(file_path)            — path security (platform-independent)
      2. _validate_export_format             — format security (platform-independent)
      3. _check_com_enabled()
      4. _check_write()
      5. _check_confirm(confirm)
      6. _check_output_path_deck — incl. M-4 for pptx
      7. Existence/overwrite check (B-5 for pdf/pptx; warn-only for images)
      8. _validate_dpi for 'images' (M-3); warn-and-ignore for pdf/pptx

    Args:
        file_path:   Absolute path to the source .pptx file.
        fmt:         Export format: 'pdf', 'pptx', or 'images' (case-insensitive).
        output_path: Output FILE path for pdf/pptx, OUTPUT DIRECTORY for images.
        dpi:         Image resolution 72-600. Only used for 'images' format.
                     None = PowerPoint default.
        confirm:     Must be True to authorise the write operation.

    Returns:
        {"exported": True, "format": str, "slide_count": int, "output_path": str}
        For 'images': additionally includes "output_dir" and "files" keys.

    Raises:
        NotAllowedError:  COM disabled, write gate off, or confirm=False.
        ValidationError:  Any gate failure (path, format, DPI, extension, M-4).
        OfficeCOMError:   COM call failed.

    Blockers: B-3, B-5, B-7, B-8, M-1, M-2, M-3, M-4.
    """
    # ── Gate chain ───────────────────────────────────────────────────────
    # Input security checks run BEFORE COM gate (platform-independent)
    resolved = _check_path(file_path)

    # M-2: normalise format BEFORE _validate_export_format
    fmt = fmt.lower() if isinstance(fmt, str) else fmt
    fmt = _validate_export_format(fmt, _VALID_DECK_FORMATS)

    _check_com_enabled()
    _check_write()
    _check_confirm(confirm)

    # M-4 same-path guard is inside _check_output_path_deck for 'pptx'
    resolved_output = _check_output_path_deck(output_path, fmt, resolved)

    # B-5 / overwrite policy per format
    if fmt in ("pdf", "pptx"):
        if resolved_output.exists():
            raise ValidationError(
                f"Output file already exists: {resolved_output.name}. "
                "Choose a different output_path or remove the existing file first."
            )
        resolved_output.parent.mkdir(parents=True, exist_ok=True)
    else:  # fmt == "images": directory may pre-exist; warn if non-empty
        if resolved_output.exists() and any(resolved_output.iterdir()):
            logger.warning(
                "Output directory %s already exists and is non-empty — "
                "COM may overwrite existing Slide*.PNG files",
                resolved_output,
            )
        resolved_output.mkdir(parents=True, exist_ok=True)

    # M-3: validate DPI (images only; warn-and-ignore for pdf/pptx)
    if fmt == "images":
        dpi = _validate_dpi(dpi)
    elif dpi is not None:
        logger.warning(
            "dpi parameter has no effect for %s format; ignoring", fmt
        )
        dpi = None

    import pythoncom

    _audit_log("ppt_export_deck", resolved,
               extra={"fmt": fmt, "output_path": str(resolved_output)})

    with _ppt_app_context(visible=False) as app:
        try:
            prs_com = app.Presentations.Open(
                str(resolved), ReadOnly=True, WithWindow=False
            )
        except pythoncom.com_error as exc:
            logger.error("COM open failed for %s: %s", resolved.name, exc)
            raise OfficeCOMError(
                f"Failed to open presentation: {resolved.name}. "
                "Check PowerPoint installation."
            ) from exc

        slide_count = prs_com.Slides.Count

        try:
            if fmt == "pdf":
                # B-7: IncludeDocProperties=0 (pos 12) — no metadata leakage
                prs_com.ExportAsFixedFormat(
                    str(resolved_output),  # pos  0: Path
                    2,                     # pos  1: ppFixedFormatTypePDF
                    2,                     # pos  2: Intent: ppFixedFormatIntentPrint
                    0,                     # pos  3: FrameSlides: msoFalse
                    2,                     # pos  4: HandoutOrder
                    1,                     # pos  5: OutputType: ppPrintOutputSlides
                    0,                     # pos  6: PrintHiddenSlides: msoFalse
                    None,                  # pos  7: PrintRange: Nothing
                    1,                     # pos  8: RangeType: ppPrintAll
                    1,                     # pos  9: SlideStart
                    slide_count,           # pos 10: SlideEnd
                    0,                     # pos 11: OpenAfterExport
                    0,                     # pos 12: IncludeDocProperties (B-7)
                )
                result: dict = {
                    "exported": True,
                    "format": fmt,
                    "slide_count": slide_count,
                    "page_count": slide_count,
                    "output_path": str(resolved_output),
                }

            elif fmt == "pptx":
                # SaveCopyAs: writes new file at output_path WITHOUT modifying source.
                # ppSaveAsPresentationOpenXML = 24
                prs_com.SaveCopyAs(str(resolved_output), 24)
                result = {
                    "exported": True,
                    "format": fmt,
                    "slide_count": slide_count,
                    "output_path": str(resolved_output),
                }

            else:  # fmt == "images"
                # Presentation.Export writes "Slide1.PNG", "Slide2.PNG", etc.
                # into the output DIRECTORY.
                if dpi is not None:
                    scale_width = int(dpi * 10.0)
                    scale_height = int(dpi * 7.5)
                    prs_com.Export(str(resolved_output), "PNG",
                                   scale_width, scale_height)
                else:
                    prs_com.Export(str(resolved_output), "PNG")

                result = {
                    "exported": True,
                    "format": fmt,
                    "slide_count": slide_count,
                    "output_dir": str(resolved_output),
                    "files": [f"Slide{i}.PNG" for i in range(1, slide_count + 1)],
                }

        except (ValidationError, OfficeCOMError):
            raise
        except pythoncom.com_error as exc:
            logger.error("ppt_export_deck COM error for %s: %s",
                         resolved.name, exc)
            raise OfficeCOMError("ppt_export_deck failed") from exc  # B-8
        finally:
            with contextlib.suppress(Exception):
                prs_com.Close()

    return result


def _make_stamped_export_dir(base_dir: str) -> Path:
    """Create and return a timestamped export directory.

    Creates: <base_dir>/exports/run-<UTC-timestamp>/png/

    The base_dir must already be validated by the caller.
    The timestamp format is Windows-filesystem-safe (no colons).
    """
    stamp = datetime.now(tz=timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
    run_dir = Path(base_dir).resolve() / "exports" / f"run-{stamp}" / "png"
    try:
        run_dir.mkdir(parents=True, exist_ok=False)
    except FileExistsError:
        raise ValidationError(f"Export run directory already exists: {run_dir}")
    return run_dir


def export_slide_as_png(
    path: str,
    slide_index: int,
    output_path: str,
    width_px: int = 0,
    height_px: int = 0,
    confirm: bool = False,
) -> dict:
    """Deprecated: use ppt_export_slide(fmt='png', slide_index=slide_index).

    Delegates to ppt_export_slide with 0-based slide_index.
    This wrapper will be removed in sprint P3.

    Args:
        path:         Absolute path to the source .pptx file.
        slide_index:  0-based slide index (legacy API).
        output_path:  Absolute path for the output .png file.
        width_px:     PNG output width in pixels (0 = PowerPoint default). 0–7680.
        height_px:    PNG output height in pixels (0 = PowerPoint default). 0–7680.
        confirm:      Must be True to authorise the write operation.

    Returns:
        {"exported": True, "slide_index": int, "output_path": str,
         "width_px": int, "height_px": int}
    """
    # US-09: early validation before COM is touched
    if not isinstance(width_px, int) or isinstance(width_px, bool):
        raise ValidationError("width_px must be an integer")
    if not isinstance(height_px, int) or isinstance(height_px, bool):
        raise ValidationError("height_px must be an integer")
    if not (0 <= width_px <= 7680):
        raise ValidationError("width_px must be 0-7680")
    if not (0 <= height_px <= 7680):
        raise ValidationError("height_px must be 0-7680")
    logger.warning(
        "export_slide_as_png is deprecated and will be removed in sprint P3. "
        "Use ppt_export_slide(fmt='png', slide_index=%d) instead.",
        slide_index,
    )
    result = ppt_export_slide(
        file_path=path,
        slide_index=slide_index,
        fmt="png",
        output_path=output_path,
        width_px=width_px,
        height_px=height_px,
        confirm=confirm,
    )
    return {
        "exported": result["exported"],
        "slide_index": slide_index,
        "output_path": result["output_path"],
        "width_px": width_px,
        "height_px": height_px,
    }


def export_deck_as_pdf(
    path: str,
    output_path: str,
    confirm: bool = False,
) -> dict:
    """Deprecated: use ppt_export_deck(fmt='pdf').

    Delegates directly to ppt_export_deck. This wrapper will be removed in sprint P3.

    Args:
        path:         Absolute path to the source .pptx file.
        output_path:  Absolute path for the output .pdf file.
        confirm:      Must be True to authorise the write operation.

    Returns:
        {"exported": True, "slide_count": int, "output_path": str}
        (slide_count key: same as before, sourced from ppt_export_deck result)
    """
    logger.warning(
        "export_deck_as_pdf is deprecated and will be removed in sprint P3. "
        "Use ppt_export_deck(fmt='pdf') instead."
    )
    result = ppt_export_deck(
        file_path=path,
        fmt="pdf",
        output_path=output_path,
        confirm=confirm,
    )
    return {
        "exported": result["exported"],
        "slide_count": result["slide_count"],
        "page_count": result["page_count"],  # US-11: alias for downstream consumers
        "output_path": result["output_path"],
    }


def export_slides_to_stamped_dir(
    path: str,
    base_dir: str,
    slide_indices: "list[int] | None" = None,
    width_px: int = 0,
    height_px: int = 0,
    confirm: bool = False,
) -> dict:
    """Export slides to exports/run-<timestamp>/png/slide-NNN.png under base_dir.

    slide_indices=None exports all slides (0-based).
    Returns: {"run_dir": str, "exported_slides": [str], "slide_count": int}

    Gate chain: _check_com_enabled → _check_write → _check_confirm
                → _check_path → allowlist(base_dir) → pixel validation.

    Security:
    - Requires PPT_ENABLE_COM=true
    - Requires PPT_ENABLE_WRITE=true
    - Requires confirm=True
    - path validated against PPT_ALLOWLIST_ROOTS
    - base_dir validated against PPT_ALLOWLIST_ROOTS
    - width_px, height_px: 0 ≤ value ≤ 7680

    US-10
    """
    _check_com_enabled()
    _check_write()
    _check_confirm(confirm)
    resolved = _check_path(path)
    if "\x00" in str(base_dir):
        raise ValidationError("Invalid base_dir: embedded null character")
    _check_in_allowlist(Path(base_dir).resolve())

    if not isinstance(width_px, int) or isinstance(width_px, bool):
        raise ValidationError("width_px must be an integer")
    if not isinstance(height_px, int) or isinstance(height_px, bool):
        raise ValidationError("height_px must be an integer")
    if not (0 <= width_px <= 7680):
        raise ValidationError("width_px must be in range 0\u20137680")
    if not (0 <= height_px <= 7680):
        raise ValidationError("height_px must be in range 0\u20137680")
    if (width_px == 0) != (height_px == 0):
        raise ValidationError(
            "width_px and height_px must both be set or both be 0"
        )

    import pythoncom  # noqa: PLC0415

    with _ppt_app_context() as app:
        try:
            prs_com = app.Presentations.Open(
                str(resolved), ReadOnly=True, WithWindow=False
            )
        except pythoncom.com_error as exc:
            logger.error("COM open failed for %s: %s", resolved.name, exc)
            raise OfficeCOMError(
                f"Failed to open presentation: {resolved.name}. "
                "Check PowerPoint installation."
            ) from exc
        try:
            slide_count = prs_com.Slides.Count
            indices: list[int] = (
                list(range(slide_count))
                if slide_indices is None
                else list(slide_indices)
            )

            # Pre-flight: validate ALL indices before any Export call
            for idx in indices:
                if not isinstance(idx, int) or isinstance(idx, bool):
                    raise ValidationError(
                        f"slide_indices contains non-integer value {idx!r}"
                    )
                if idx < 0 or idx >= slide_count:
                    raise ValidationError(
                        f"slide_index {idx} out of range "
                        f"(0-based; presentation has {slide_count} slide(s))"
                    )

            run_dir = _make_stamped_export_dir(str(Path(base_dir).resolve()))

            exported: list[str] = []
            try:
                for idx in indices:
                    out_path = run_dir / f"slide-{idx:03d}.png"
                    slide = prs_com.Slides(idx + 1)
                    if width_px > 0 and height_px > 0:
                        slide.Export(str(out_path), "PNG", width_px, height_px)
                    else:
                        slide.Export(str(out_path), "PNG")
                    exported.append(str(out_path))
            except (ValidationError, OfficeCOMError):
                raise
            except Exception as exc:
                logger.error(
                    "export_slides_to_stamped_dir COM error at slide %d: %s",
                    idx,
                    exc,
                )
                raise OfficeCOMError(
                    f"export_slides_to_stamped_dir failed at slide {idx + 1}. "
                    "Partial output left in run directory."
                ) from exc
        finally:
            with contextlib.suppress(Exception):
                prs_com.Close()

    return {
        "run_dir": str(run_dir),
        "exported_slides": exported,
        "slide_count": len(exported),
    }


# ---------------------------------------------------------------------------
# US-17: export_changed_slides_only
# ---------------------------------------------------------------------------

_MD5_HEX_RE = re.compile(r"^[0-9a-f]{32}$")


def _compute_slide_hash(slide: Any) -> str:
    """Compute deterministic MD5 hash of slide XML content."""
    return hashlib.md5(slide.element.xml.encode("utf-8")).hexdigest()


def export_changed_slides_only(
    path: str,
    base_dir: str,
    previous_hashes: dict,
    width_px: int = 0,
    height_px: int = 0,
    confirm: bool = False,
) -> dict:
    """Export only slides whose content has changed since previous_hashes was computed.

    previous_hashes: dict of {str(slide_index): md5_hex_32_chars}.
    Returns: {run_dir, exported_slides, unchanged_slides, new_hashes}

    Gate chain: _check_com_enabled → _check_write → _check_confirm
                → _check_path → allowlist(base_dir) → pixel validation
                → previous_hashes validation → python-pptx hash comparison
                → COM export (changed slides only).

    Security:
    - Requires PPT_ENABLE_COM=true
    - Requires PPT_ENABLE_WRITE=true
    - Requires confirm=True
    - path validated against PPT_ALLOWLIST_ROOTS
    - base_dir: null-byte check + PPT_ALLOWLIST_ROOTS
    - previous_hashes values: validated as ^[0-9a-f]{32}$ before use
    - width_px, height_px: 0 ≤ value ≤ 7680, isinstance guard excluding bool

    US-17
    """
    # ── Gate chain ─────────────────────────────────────────────────────────
    _check_com_enabled()
    _check_write()
    _check_confirm(confirm)
    resolved = _check_path(path)
    if not resolved.is_file():
        raise ValidationError(f"Presentation file not found: {resolved.name}")

    if "\x00" in str(base_dir):
        raise ValidationError("Invalid base_dir: embedded null character")
    _check_in_allowlist(Path(base_dir).resolve())
    if not Path(base_dir).resolve().is_dir():
        raise ValidationError(
            f"base_dir does not exist or is not a directory: {base_dir}"
        )

    # ── Pixel-dimension validation (same bounds as export_slide_as_png) ───
    if not isinstance(width_px, int) or isinstance(width_px, bool):
        raise ValidationError("width_px must be an integer")
    if not isinstance(height_px, int) or isinstance(height_px, bool):
        raise ValidationError("height_px must be an integer")
    if not (0 <= width_px <= 7680):
        raise ValidationError("width_px must be in range 0\u20137680")
    if not (0 <= height_px <= 7680):
        raise ValidationError("height_px must be in range 0\u20137680")
    if (width_px == 0) != (height_px == 0):
        raise ValidationError(
            "width_px and height_px must both be set or both be 0"
        )

    # ── Validate previous_hashes values are well-formed MD5 hex strings ──
    if not isinstance(previous_hashes, dict):
        raise ValidationError("previous_hashes must be a dict")
    for k, v in previous_hashes.items():
        if not isinstance(v, str) or not _MD5_HEX_RE.match(v):
            raise ValidationError(
                f"previous_hashes[{k!r}] is not a valid 32-char MD5 hex string: {v!r}"
            )

    # ── Load via python-pptx to compute slide content hashes ─────────────
    import pptx as _pptx  # noqa: PLC0415

    prs_obj = _pptx.Presentation(str(resolved))

    new_hashes: dict[str, str] = {}
    changed: list[int] = []
    unchanged: list[int] = []

    for i, slide in enumerate(prs_obj.slides):
        h = _compute_slide_hash(slide)
        new_hashes[str(i)] = h
        prev = previous_hashes.get(str(i))
        if prev is None or prev != h:
            changed.append(i)
        else:
            unchanged.append(i)

    # ── Always create the run directory ──────────────────────────────────
    run_dir = _make_stamped_export_dir(str(Path(base_dir).resolve()))

    # ── Export only changed slides via COM ────────────────────────────────
    if changed:
        import pythoncom  # noqa: PLC0415

        _audit_log(
            "export_changed_slides_only",
            resolved,
            extra={"changed_count": len(changed), "run_dir": str(run_dir)},
        )

        with _ppt_app_context() as app:
            try:
                prs_com = app.Presentations.Open(
                    str(resolved), ReadOnly=True, WithWindow=False
                )
            except pythoncom.com_error as exc:
                logger.error("COM open failed for %s: %s", resolved.name, exc)
                raise OfficeCOMError(
                    f"Failed to open presentation: {resolved.name}. "
                    "Check PowerPoint installation."
                ) from exc
            try:
                for idx in changed:
                    out_path = run_dir / f"slide-{idx:03d}.png"
                    slide_com = prs_com.Slides(idx + 1)
                    if width_px > 0 and height_px > 0:
                        slide_com.Export(str(out_path), "PNG", width_px, height_px)
                    else:
                        slide_com.Export(str(out_path), "PNG")
            except (ValidationError, OfficeCOMError):
                raise
            except Exception as exc:
                logger.error(
                    "export_changed_slides_only COM error at slide %d: %s",
                    idx,
                    exc,
                )
                raise OfficeCOMError(
                    f"export_changed_slides_only failed at slide {idx + 1}."
                ) from exc
            finally:
                with contextlib.suppress(Exception):
                    prs_com.Close()

    result = {
        "run_dir": str(run_dir),
        "exported_slides": changed,
        "unchanged_slides": unchanged,
        "new_hashes": new_hashes,
    }
    if not changed:
        result["exported_count"] = 0
        result["note"] = "no changed slides to export"
    return result
