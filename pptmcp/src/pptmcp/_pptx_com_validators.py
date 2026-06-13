"""
Path validators, format validators, hash utilities and output-path constants
for the pptmcp COM layer.

Private module — not part of the public API.
Does NOT import from _pptx_com_guards (avoids circular deps).
"""
from __future__ import annotations

import logging
import os
from pathlib import Path

from pptmcp.presentation_pptx import ValidationError

__all__: list[str] = []

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Output path validator (PNG / PDF) — legacy constants
# ---------------------------------------------------------------------------

_ALLOWED_OUTPUT_EXTENSIONS: frozenset[str] = frozenset({".png", ".pdf"})

# ── Export format allowlists ──────────────────────────────────────────────
_VALID_SLIDE_FORMATS: frozenset[str] = frozenset({"png", "pdf", "svg"})
_VALID_DECK_FORMATS: frozenset[str] = frozenset({"pdf", "pptx", "images"})


def _validate_export_format(fmt: str, allowed: frozenset) -> str:
    """Validate and normalise an export format string.

    M-1: isinstance guard prevents AttributeError on None input.
    M-2: normalises to lowercase so routing dict lookups cannot raise KeyError.

    Args:
        fmt:     Client-supplied format string (may be None or wrong case).
        allowed: Frozen set of lower-case allowed format strings.

    Returns:
        Normalised (lower-case) format string.

    Raises:
        ValidationError: fmt is not a str, empty, or not in the allowed set.
    """
    # M-1: isinstance check BEFORE .lower() — prevents AttributeError on None
    if not isinstance(fmt, str) or fmt.lower() not in allowed:
        raise ValidationError(
            f"Unsupported format: {fmt!r}. "
            f"Allowed: {sorted(allowed)}"
        )
    return fmt.lower()


def _validate_dpi(dpi: "int | None") -> "int | None":
    """Validate the DPI parameter for image exports.

    M-3: Enforces 72 ≤ dpi ≤ 600 and rejects floats and non-int types.
    None is passed through unchanged (means: use PowerPoint default of 96 DPI).

    Args:
        dpi: DPI value to validate, or None.

    Returns:
        dpi unchanged if None; validated int otherwise.

    Raises:
        ValidationError: dpi is not None and fails the int/bounds check.
    """
    if dpi is None:
        return None
    # M-3: isinstance rejects float (e.g. 150.0) which would cause COM VARIANT errors
    if not isinstance(dpi, int) or isinstance(dpi, bool) or not (72 <= dpi <= 600):
        raise ValidationError(
            f"dpi must be an integer between 72 and 600, got {dpi!r}. "
            "Use dpi=None for the PowerPoint default (96 DPI)."
        )
    return dpi


def _check_in_allowlist(resolved: Path) -> None:
    """Assert a fully-resolved path is within PPT_ALLOWLIST_ROOTS.

    Used by both _check_output_file and _check_output_dir to avoid
    duplicating allowlist logic.

    Args:
        resolved: Absolute, fully-resolved Path to check.

    Raises:
        ValidationError: PPT_ALLOWLIST_ROOTS is not configured, or the path
                         is not within any configured root.
    """
    roots_env = os.getenv("PPT_ALLOWLIST_ROOTS", "").strip()
    if not roots_env:
        raise ValidationError("No allowlist configured (PPT_ALLOWLIST_ROOTS)")
    roots = [r.strip() for r in roots_env.split(",") if r.strip()]
    for root in roots:
        try:
            if resolved.is_relative_to(Path(root).resolve()):
                return
        except (ValueError, TypeError):
            continue
    raise ValidationError(
        f"Output path is not within the configured allowlist "
        f"(PPT_ALLOWLIST_ROOTS={roots_env!r})"
    )


def _check_output_file(output_path: str, expected_ext: str) -> Path:
    """Validate a file output path: null bytes, correct extension, within allowlist.

    B-3: null-byte check is the FIRST operation, before Path() construction.

    Args:
        output_path:  Absolute path for the output file (not yet created).
        expected_ext: Required suffix including the dot, e.g. '.png'.

    Returns:
        Resolved Path.

    Raises:
        ValidationError: null byte, wrong extension, or outside allowlist.
    """
    # B-3: explicit null-byte check BEFORE Path() call
    if "\x00" in str(output_path):
        raise ValidationError("Invalid output path: embedded null character")
    resolved = Path(output_path).resolve()
    if resolved.suffix.lower() != expected_ext:
        raise ValidationError(
            f"Output extension must be {expected_ext!r} for this format, "
            f"got {resolved.suffix.lower()!r}"
        )
    _check_in_allowlist(resolved)
    return resolved


def _check_output_dir(output_path: str) -> Path:
    """Validate a directory output path for 'images' format.

    B-3: null-byte check is the FIRST operation, before Path() construction.
    The directory does NOT need to exist yet; it will be created at export time.

    Args:
        output_path: Path that will be used as the output directory.

    Returns:
        Resolved Path of the directory (may not yet exist).

    Raises:
        ValidationError: null byte, path is an existing FILE, or outside allowlist.
    """
    # B-3: explicit null-byte check BEFORE Path() call
    if "\x00" in str(output_path):
        raise ValidationError("Invalid output path: embedded null character")
    resolved = Path(output_path).resolve()
    if resolved.is_file():
        raise ValidationError(
            f"Output path {resolved.name!r} already exists as a file. "
            "Provide a directory path for 'images' format."
        )
    _check_in_allowlist(resolved)
    return resolved


# Extension map for slide export formats
_SLIDE_EXT_MAP: dict[str, str] = {"png": ".png", "pdf": ".pdf", "svg": ".svg"}


def _check_output_path_slide(output_path: str, fmt: str) -> Path:
    """Validate slide export output path for fmt in {'png','pdf','svg'}.

    Delegates to _check_output_file with the correct expected extension.
    fmt MUST already be normalised (lower-case) before calling this function.

    Args:
        output_path: Absolute output file path.
        fmt:         Normalised format string ('png', 'pdf', or 'svg').

    Returns:
        Resolved Path.

    Raises:
        ValidationError: from _check_output_file.
    """
    return _check_output_file(output_path, _SLIDE_EXT_MAP[fmt])


# Extension map for deck file export formats
_DECK_EXT_MAP: dict[str, str] = {"pdf": ".pdf", "pptx": ".pptx"}


def _check_output_path_deck(
    output_path: str,
    fmt: str,
    source_resolved: Path,
) -> Path:
    """Validate deck export output path for fmt in {'pdf','pptx','images'}.

    M-4: For 'pptx' format, explicitly blocks same source == output path.
         This is an INTENTIONAL guard, independent of the B-5 file-exists check.

    Args:
        output_path:      Absolute output file path (pdf/pptx) or directory (images).
        fmt:              Normalised format string ('pdf', 'pptx', or 'images').
        source_resolved:  Resolved source .pptx path (for M-4 same-path check).

    Returns:
        Resolved Path (file for pdf/pptx; directory for images).

    Raises:
        ValidationError: extension mismatch, allowlist violation, or M-4 same-path.
    """
    if fmt in _DECK_EXT_MAP:
        resolved_output = _check_output_file(output_path, _DECK_EXT_MAP[fmt])
        # M-4: Explicit same-path guard for pptx — does NOT rely on B-5 file-exists.
        # Intent: prevent SaveCopyAs from silently replacing the source. This guard
        # is race-condition-proof because it compares resolved paths (not strings).
        if fmt == "pptx" and resolved_output == source_resolved:
            raise ValidationError(
                "output_path cannot be the same as the source file. "
                "Provide a separate destination path for the .pptx copy."
            )
        return resolved_output
    else:  # fmt == "images"
        return _check_output_dir(output_path)


def _check_output_path(output_path: str) -> Path:
    """Validate an export output path against PPT_ALLOWLIST_ROOTS.

    Accepts .png and .pdf extensions only.
    Raises ValidationError for null bytes, bad extension, or path outside allowlist.
    Addresses B-3 (null-byte injection).
    """
    # B-3: Explicit null-byte check BEFORE Path() call
    if "\x00" in str(output_path):
        raise ValidationError("Invalid output path: embedded null character")

    roots_env = os.getenv("PPT_ALLOWLIST_ROOTS", "").strip()
    if not roots_env:
        raise ValidationError("No allowlist configured (PPT_ALLOWLIST_ROOTS)")

    resolved = Path(output_path).resolve()

    ext = resolved.suffix.lower()
    if ext not in _ALLOWED_OUTPUT_EXTENSIONS:
        raise ValidationError(
            f"Unsupported output extension: {ext!r}. Allowed: .png, .pdf"
        )

    roots = [r.strip() for r in roots_env.split(",") if r.strip()]
    for root in roots:
        try:
            if resolved.is_relative_to(Path(root).resolve()):
                return resolved
        except (ValueError, TypeError):
            continue

    raise ValidationError("Output path is not within the configured allowlist")
