"""_review_export — review_slide_export function for pptmcp review pipeline."""
from __future__ import annotations

import json
import logging

from pptmcp._review_detectors import (
    _detect_low_font_size,
    _detect_png_issues,
    _detect_shapes_outside_slide,
    _detect_sparse_slide,
    _detect_table_content_clipping,
    _detect_text_overflow_heuristic,
    _detect_unfinished_placeholder,
)
from pptmcp.contract_pptx import _check_json_path, check_presentation_against_contract
from pptmcp.presentation_pptx import (
    ValidationError,
    _check_image_path,
    _check_path,
    _check_write,
    _check_confirm,
    _load_prs,
)

_log = logging.getLogger(__name__)


def review_slide_export(
    path: str,
    slide_index: int,
    png_path: str,
    contract_path: str | None = None,
    confirm: bool = False,
) -> dict:
    """Review a slide export against quality criteria.

    Checks TEXT_OVERFLOW, SHAPE_OUTSIDE_SLIDE, TABLE_CONTENT_CLIPPING,
    UNFINISHED_PLACEHOLDER, LOW_FONT_SIZE, and SPARSE_SLIDE for the specified slide.
    png_path must be a valid, existing PNG file in the allowlist (validated but not pixel-read).

    contract_path is optional. When provided, contract-driven findings (slide
    dimensions, per-slide shape/element specs from check_presentation_against_contract)
    relevant to the reviewed slide are merged into the structural ``findings``
    list. CONTRACT_* criteria contribute to ``passed`` like any other structural
    check.

    Writes ``review_result.json`` beside the PNG with the full result dict so
    downstream agents have a durable audit artefact. The output path is
    returned as ``review_result_path``.

    Write-gated (PPT_ENABLE_WRITE=true) + confirm=True required. Writes
    ``review_result.json`` beside the PNG as a durable audit artefact.

    passed=True indicates no structural issues detected; visual delivery
    readiness requires human review.

    Returns:
        dict with keys: slide_index, png_path, pptx_path, passed (bool),
        findings (list), visual_qa_findings (list), checks_run (list),
        findings_count (int), review_result_path (str), visual_readiness (str),
        confidence_boundary (str).

    Note:
        ``findings`` (structural) uses keys: criterion, result, detail, severity (lowercase).
        ``visual_qa_findings`` (advisory) uses keys: severity (lowercase), check, detail.
        These two lists have intentionally different schemas; ``passed`` reflects structural
        checks only.
    """
    _check_write()
    _check_confirm(confirm)
    # --- Security: validate all paths first (fail-fast before any file I/O) ---
    resolved = _check_path(path)                          # OWASP A01 + .pptx extension
    resolved_png = _check_image_path(png_path)            # OWASP A01 + image extension
    if not resolved_png.exists():                         # BLK-D02
        raise ValidationError(f"PNG file not found: {resolved_png!r}")
    resolved_contract = None
    if contract_path is not None:
        resolved_contract = _check_json_path(contract_path)  # OWASP A01 + .json extension

    # --- Load presentation ---
    prs_obj = _load_prs(resolved)  # H-015: use _load_prs instead of bare Presentation()

    # --- BLK-D01: explicit non-negative + upper-bound slide index check ---
    if not (isinstance(slide_index, int) and 0 <= slide_index < len(prs_obj.slides)):
        raise ValidationError(
            f"slide_index {slide_index!r} out of range; "
            f"presentation has {len(prs_obj.slides)} slide(s)"
        )

    slide = prs_obj.slides[slide_index]

    # --- Run checks ---
    # Structural checks determine `passed` and populate `findings`
    findings: list[dict] = []
    findings.extend(_detect_text_overflow_heuristic(slide, prs_obj))
    findings.extend(
        _detect_shapes_outside_slide(slide, prs_obj.slide_width, prs_obj.slide_height)
    )
    findings.extend(_detect_table_content_clipping(slide))

    checks_run = [
        "TEXT_OVERFLOW",
        "SHAPE_OUTSIDE_SLIDE",
        "TABLE_CONTENT_CLIPPING",
        "UNFINISHED_PLACEHOLDER",
        "LOW_FONT_SIZE",
        "SPARSE_SLIDE",
    ]

    # --- Contract-driven findings (RCI-US-13) ---
    expected_resolution: tuple[int, int] | None = None
    if resolved_contract is not None:
        try:
            contract_result = check_presentation_against_contract(
                str(resolved), str(resolved_contract)
            )
        except ValidationError:
            raise
        except Exception as exc:  # noqa: BLE001
            _log.warning("contract evaluation failed: %s", exc)
        else:
            for cf in contract_result.get("findings", []):
                cf_idx = cf.get("slide_index")
                if cf_idx not in (None, slide_index):
                    continue
                findings.append({
                    "criterion": f"CONTRACT_{cf.get('criterion', 'UNKNOWN')}",
                    "result": "FAIL",
                    "detail": cf.get("detail", ""),
                    "severity": cf.get("severity", "high"),
                })
        checks_run.append("CONTRACT")

        try:
            import json as _json  # noqa: PLC0415
            with open(resolved_contract, encoding="utf-8") as _f:
                _raw_contract = _json.load(_f)
            res = _raw_contract.get("expected_resolution")
            if (
                isinstance(res, (list, tuple))
                and len(res) == 2
                and all(isinstance(v, int) and not isinstance(v, bool) for v in res)
            ):
                expected_resolution = (int(res[0]), int(res[1]))
        except Exception as exc:  # noqa: BLE001
            _log.debug("expected_resolution not readable: %s", exc)

    # --- Pixel-read PNG checks ---
    findings.extend(
        _detect_png_issues(resolved_png, expected_resolution=expected_resolution)
    )
    checks_run.extend(["PIXEL_DIMENSION_MISMATCH", "BLANK_SLIDE"])

    # Visual QA checks — advisory only; do NOT affect `passed`
    visual_qa_findings: list[dict] = []
    visual_qa_findings.extend(_detect_unfinished_placeholder(slide))
    visual_qa_findings.extend(_detect_low_font_size(slide))
    visual_qa_findings.extend(
        _detect_sparse_slide(slide, prs_obj.slide_width, prs_obj.slide_height)
    )

    # --- Aggregate overall status (RCI-US-14) ---
    if findings:
        overall = "FAIL"
    elif visual_qa_findings:
        overall = "WARN"
    else:
        overall = "PASS"

    # --- Write review_result.json beside the PNG (RCI-US-13) ---
    review_result_path = resolved_png.parent / "review_result.json"
    result: dict = {
        "slide_index": slide_index,
        "png_path": str(resolved_png),
        "pptx_path": str(resolved),
        "overall": overall,
        "passed": len(findings) == 0,
        "findings": findings,
        "visual_qa_findings": visual_qa_findings,
        "checks_run": checks_run,
        "findings_count": len(findings),
        "review_result_path": str(review_result_path),
        "visual_readiness": "not_assessed",
        "confidence_boundary": "structural_heuristics_only",
    }
    try:
        review_result_path.write_text(
            json.dumps(result, indent=2, default=str), encoding="utf-8"
        )
    except OSError as exc:
        _log.warning(
            "could not write review_result.json to %s: %s",
            review_result_path,
            exc,
        )

    return result
