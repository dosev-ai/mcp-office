"""Canonical dynamic-array finalization boundary for workbook artifacts.

All mutating save completions should route through
``finalize_workbook_da_artifact`` so the DA post-save flow has one
deterministic orchestration point:

1) patch -> 2) classify -> 3) invariant-validate -> 4) optional COM-probe.

Phase A keeps classify/validate/probe extensible: if newer helpers are not
available yet, their stages are marked as skipped without breaking save flows.
"""

from __future__ import annotations

import logging
import os
from typing import Any

from . import _da_com, _ooxml


class DAFinalizeError(RuntimeError):
    """Raised when DA finalization fails in strict mode."""


def _get_finalize_mode() -> str:
    mode = os.getenv("EXCEL_DA_FINALIZE_MODE", "warn").strip().lower()
    return mode if mode in {"warn", "strict"} else "warn"


def _get_com_probe_policy() -> str:
    policy = os.getenv("EXCEL_DA_COM_PROBE", "off").strip().lower()
    return policy if policy in {"off", "auto", "required"} else "off"


def finalize_workbook_da_artifact(
    path: str,
    require_com_probe: bool = False,
) -> dict[str, Any]:
    """Run DA finalization stages and return structured diagnostics.

    In ``EXCEL_DA_FINALIZE_MODE=strict``, a failed stage raises
    ``DAFinalizeError`` after diagnostics are computed.
    """

    mode = _get_finalize_mode()
    probe_policy = _get_com_probe_policy()
    diagnostics: dict[str, Any] = {
        "path": path,
        "mode": mode,
        "com_probe_policy": probe_policy,
        "stages": [],
    }
    logger = logging.getLogger(__name__)
    failed = False

    # Stage 1: patch
    patched_cells = 0
    try:
        patched_cells = int(_ooxml.patch_dynamic_array_metadata(path))
        diagnostics["stages"].append(
            {"name": "patch", "status": "ok", "patched_cells": patched_cells}
        )
    except Exception as exc:  # pragma: no cover
        failed = True
        diagnostics["stages"].append(
            {"name": "patch", "status": "error", "error": str(exc)}
        )
        logger.warning("[excelmcp] DA patch failed for %s: %s", path, exc)

    # Stage 2: classify (optional until helper lands)
    classify_fn = getattr(_ooxml, "classify_patch_outcome", None)
    if callable(classify_fn):
        try:
            classification = classify_fn(path, patched_cells)
            diagnostics["stages"].append(
                {"name": "classify", "status": "ok", "result": classification}
            )
        except Exception as exc:  # pragma: no cover
            failed = True
            diagnostics["stages"].append(
                {"name": "classify", "status": "error", "error": str(exc)}
            )
            logger.warning("[excelmcp] DA classify failed for %s: %s", path, exc)
    else:
        fallback = "patched" if patched_cells > 0 else "no_da_detected"
        diagnostics["stages"].append(
            {"name": "classify", "status": "ok", "result": fallback}
        )

    # Stage 3: invariant validation (optional until helper lands)
    validate_fn = getattr(_ooxml, "validate_dynamic_array_invariants", None)
    if callable(validate_fn):
        try:
            validation_result = validate_fn(path)
            ok = bool(validation_result.get("ok", False))
            diagnostics["stages"].append(
                {
                    "name": "validate",
                    "status": "ok" if ok else "error",
                    "result": validation_result,
                }
            )
            failed = failed or (not ok)
        except Exception as exc:  # pragma: no cover
            failed = True
            diagnostics["stages"].append(
                {"name": "validate", "status": "error", "error": str(exc)}
            )
            logger.warning("[excelmcp] DA validation failed for %s: %s", path, exc)
    else:
        diagnostics["stages"].append(
            {"name": "validate", "status": "skipped", "reason": "not_implemented"}
        )

    # Stage 4: optional COM probe (optional until helper lands)
    should_probe = require_com_probe or probe_policy == "required"
    probe_fn = getattr(_da_com, "probe_excel_reopen", None)
    if should_probe:
        if callable(probe_fn):
            try:
                probe_result = probe_fn(path)
                ok = bool(probe_result.get("ok", False))
                diagnostics["stages"].append(
                    {
                        "name": "probe",
                        "status": "ok" if ok else "error",
                        "result": probe_result,
                    }
                )
                failed = failed or (not ok)
            except Exception as exc:  # pragma: no cover
                failed = True
                diagnostics["stages"].append(
                    {"name": "probe", "status": "error", "error": str(exc)}
                )
                logger.warning("[excelmcp] DA COM probe failed for %s: %s", path, exc)
        else:
            failed = True
            diagnostics["stages"].append(
                {
                    "name": "probe",
                    "status": "error",
                    "reason": "required_but_not_implemented",
                }
            )
    else:
        diagnostics["stages"].append(
            {"name": "probe", "status": "skipped", "reason": "policy_off"}
        )

    diagnostics["ok"] = not failed
    if failed and mode == "strict":
        raise DAFinalizeError(f"DA finalization failed for {path}")
    return diagnostics
