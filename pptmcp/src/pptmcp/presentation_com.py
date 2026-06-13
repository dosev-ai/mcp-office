"""
COM automation layer for PowerPoint MCP Phase 2 (Windows-only).
Exposes: export_slide_as_png, export_deck_as_pdf, run_slide_show, recalculate_charts.

No FastMCP / MCP imports — pure win32com + stdlib only.
All mutations require both PPT_ENABLE_WRITE=true AND confirm=True.
Slide show is a display-only operation (no write gate, no confirm).

Blockers addressed:
  B-1  run_slide_show — SlideShowSettings accessed as plain attribute, not 'with' block
  B-2  recalculate_charts — msoAutomationSecurityForceDisable set BEFORE Open(); RefreshAll
         replaced by Activate()-only (SSRF-safe Option A)
  B-3  Null-byte path injection — explicit check in _check_output_path()
  B-4  run_slide_show — _is_slideshow_running() guard before Dispatch
  B-5  Export overwrite — existence check raises ValidationError before COM is called
  B-6  run_slide_show — prs_obj.Saved = True immediately after Open()
  B-7  export_deck_as_pdf — IncludeDocProperties=False
  B-8  OfficeCOMError messages strip raw HRESULT / COM detail; detail goes to stderr only
  B-9  recalculate_charts — shape.Type in (3, 7) covers OLE-embedded Excel charts

Split layout (decomposed from original 1171-line monolith):
  _pptx_com_guards.py     — _check_com_enabled, _ppt_app_context, _is_slideshow_running
  _pptx_com_validators.py — constants, path validators, format validators, hash utils
  _pptx_com_export.py     — all 7 export functions (public + deprecated)
  presentation_com.py     — run_slide_show + recalculate_charts + re-exports (this file)
"""
from __future__ import annotations

import contextlib
import logging

from pptmcp.presentation_pptx import (
    NotAllowedError,
    OfficeCOMError,
    PPTMCPError,
    ValidationError,
    _audit_log,
    _check_confirm,
    _check_path,
    _check_write,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Public COM tools (residual — display/recalculate operations)
# ---------------------------------------------------------------------------

def run_slide_show(path: str, confirm: bool = False) -> dict:
    """Launch a PowerPoint slide show for the given .pptx file.

    This is a DISPLAY-ONLY operation — it does not write to the source file.
    Requires PPT_ENABLE_WRITE=true and confirm=True to prevent accidental launches.

    Gate order: _check_com_enabled -> _check_write -> _check_confirm -> _check_path
                -> _is_slideshow_running check.

    PowerPoint is left alive after this call so the slide show window stays open.
    The COM reference is released from Python's side; POWERPNT.EXE continues
    running independently.

    Args:
        path:    Absolute path to the source .pptx file.
        confirm: Must be True to authorise the operation.

    Returns:
        {"running": True, "path": str, "slide_count": int}

    Blockers: B-1, B-4, B-6, B-8.
    """
    _check_com_enabled()
    _check_write()
    _check_confirm(confirm)
    resolved = _check_path(path)

    # B-4: Block repeated launches — check for an already-running show
    if _is_slideshow_running(resolved):
        raise ValidationError(
            "A slide show is already running for this file. "
            "Close the existing show before starting a new one."
        )

    import pythoncom
    import win32com.client

    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = None
    try:
        _audit_log("run_slide_show", resolved)

        # Use raw Dispatch — NOT _ppt_app_context — so the process stays alive
        # after this function returns (slide show window must remain open).
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True           # Must be visible: user needs to see the show
        try:
            app.DisplayAlerts = False
        except Exception:
            logger.warning("Could not set DisplayAlerts=False; continuing", exc_info=True)

        prs_obj = app.Presentations.Open(
            str(resolved),
            ReadOnly=True,
            WithWindow=True,
        )

        # B-6: Mark as clean immediately to prevent any accidental write-back
        prs_obj.Saved = True

        # B-1: Access SlideShowSettings as a plain attribute — NOT a 'with' block.
        # (win32com CDispatch objects don't implement __enter__/__exit__; a 'with'
        #  statement would raise AttributeError at runtime.)
        ss = prs_obj.SlideShowSettings
        ss.ShowType = 1              # ppShowTypeSpeaker = 1
        ss.Run()

        slide_count = prs_obj.Slides.Count

        # Release Python references; POWERPNT.EXE stays alive with the show window.
        del prs_obj
        del ss

        return {
            "running": True,
            "path": str(resolved),
            "slide_count": slide_count,
        }

    except (PPTMCPError, ValidationError, NotAllowedError):
        # Re-raise our own errors untouched
        if app is not None:
            with contextlib.suppress(Exception):
                app.Quit()
        raise

    except Exception as exc:
        logger.error("run_slide_show COM error for %s: %s", resolved.name, exc)
        if app is not None:
            with contextlib.suppress(Exception):
                app.Quit()
        raise OfficeCOMError("run_slide_show failed") from exc  # B-8

    finally:
        pythoncom.CoUninitialize()


def recalculate_charts(path: str, confirm: bool = False) -> dict:
    """Open a .pptx via COM and activate all embedded charts for recalculation.

    SSRF SAFETY (B-2, Option A):
      RefreshAll() is NOT called on any chart workbook. Only chart_data.Activate()
      is called, which recalculates formulas already in the embedded workbook
      without triggering Power Query / web queries / external data connections.
      This eliminates the SSRF vector while still updating formula-driven chart values.

    Linked charts (ChartData.IsLinked=True) are skipped — they reference external
    .xlsx files and would require network/disk access beyond the allowlist.

    OLE-embedded Excel charts (shape.Type=7) are also handled (B-9).

    msoAutomationSecurityForceDisable is set on the app BEFORE opening the file to
    prevent macro execution during open (B-2, belt-and-suspenders).

    Gate order: _check_com_enabled -> _check_write -> _check_confirm -> _check_path.

    Args:
        path:    Absolute path to the .pptx file to recalculate.
        confirm: Must be True to authorise the write operation.

    Returns:
        {
            "recalculated": True,
            "charts_activated": int,   # number of embedded charts activated
            "skipped_linked_charts": int,
            "note": str,
        }

    Blockers: B-2, B-8, B-9.
    """
    _check_com_enabled()
    _check_write()
    _check_confirm(confirm)
    resolved = _check_path(path)

    import pythoncom

    _audit_log("recalculate_charts", resolved)

    # mso shape type constants
    MSO_CHART = 3              # msoChart — native PowerPoint chart
    MSO_EMBEDDED_OLE = 7       # msoEmbeddedOLEObject — OLE-embedded object (e.g. Excel chart)

    with _ppt_app_context(visible=False) as app:
        # B-2: Disable macros and external data refresh BEFORE opening any file.
        # msoAutomationSecurityForceDisable = 3
        original_security = app.AutomationSecurity
        app.AutomationSecurity = 3

        try:
            prs_com = app.Presentations.Open(
                str(resolved),
                ReadOnly=False,
                WithWindow=False,
            )
        except pythoncom.com_error as exc:
            app.AutomationSecurity = original_security
            logger.error("COM open failed for %s: %s", resolved.name, exc)
            raise OfficeCOMError(
                f"Failed to open presentation: {resolved.name}. "
                "Check PowerPoint installation."
            ) from exc

        charts_activated = 0
        skipped_linked = 0

        try:
            slide_count = prs_com.Slides.Count
            for slide_i in range(1, slide_count + 1):
                slide = prs_com.Slides(slide_i)
                shape_count = slide.Shapes.Count
                for shape_i in range(1, shape_count + 1):
                    shape = slide.Shapes(shape_i)

                    # B-9: handle both native charts (Type=3) and OLE-embedded (Type=7)
                    if shape.Type == MSO_CHART:
                        # Native PowerPoint chart
                        pass  # falls through to chart_data access below
                    elif shape.Type == MSO_EMBEDDED_OLE:
                        # OLE-embedded object — verify it is an Excel chart
                        try:
                            prog_id = shape.OLEFormat.ProgID
                            if not prog_id.startswith("Excel.Chart"):
                                continue
                        except Exception:
                            continue
                    else:
                        continue

                    # Access chart data
                    try:
                        chart_data = shape.Chart.ChartData
                    except Exception:
                        logger.debug(
                            "Slide %d shape %d: could not access ChartData, skipping",
                            slide_i, shape_i,
                        )
                        continue

                    # Skip externally linked charts — they reference files on disk/UNC
                    try:
                        is_linked = chart_data.IsLinked
                    except Exception:
                        is_linked = False

                    if is_linked:
                        logger.warning(
                            "Slide %d shape %d: ChartData.IsLinked=True — skipping "
                            "(external link; not refreshed for safety)",
                            slide_i, shape_i,
                        )
                        skipped_linked += 1
                        continue

                    # B-2 Option A: Activate only — NO RefreshAll() — SSRF-safe.
                    # Activate() opens the embedded OLE workbook in-process and
                    # recalculates existing formulas without executing Power Query
                    # or any external data connections.
                    try:
                        chart_data.Activate()
                        charts_activated += 1
                    except Exception as exc:
                        logger.warning(
                            "Slide %d shape %d: chart_data.Activate() failed: %s",
                            slide_i, shape_i, exc,
                        )

            prs_com.Save()

        except pythoncom.com_error as exc:
            logger.error("recalculate_charts COM error for %s: %s",
                         resolved.name, exc)
            raise OfficeCOMError("recalculate_charts failed") from exc  # B-8

        finally:
            # Restore original automation security level
            with contextlib.suppress(Exception):
                app.AutomationSecurity = original_security
            with contextlib.suppress(Exception):
                prs_com.Close()

    return {
        "recalculated": True,
        "charts_activated": charts_activated,
        "skipped_linked_charts": skipped_linked,
        "note": (
            "RefreshAll() is disabled for SSRF safety. "
            "Embedded workbook formulas were recalculated via Activate() only. "
            "External data connections (Power Query, web queries) were not executed."
        ),
    }


# -- Re-exports for backward compatibility ----------------------------------
from pptmcp._pptx_com_export import (  # noqa: E402, F401
    _MD5_HEX_RE,
    _compute_slide_hash,
    _make_stamped_export_dir,
    export_changed_slides_only,
    export_deck_as_pdf,
    export_slide_as_png,
    export_slides_to_stamped_dir,
    ppt_export_deck,
    ppt_export_slide,
)
from pptmcp._pptx_com_guards import (  # noqa: E402, F401
    _check_com_enabled,
    _is_slideshow_running,
    _ppt_app_context,
)
from pptmcp._pptx_com_validators import (  # noqa: E402, F401
    _ALLOWED_OUTPUT_EXTENSIONS,
    _DECK_EXT_MAP,
    _SLIDE_EXT_MAP,
    _VALID_DECK_FORMATS,
    _VALID_SLIDE_FORMATS,
    _check_in_allowlist,
    _check_output_dir,
    _check_output_file,
    _check_output_path,
    _check_output_path_deck,
    _check_output_path_slide,
    _validate_dpi,
    _validate_export_format,
)

_SUBMODULE_MAP: dict[str, str] = {
    "_check_com_enabled": "_pptx_com_guards",
    "_ppt_app_context": "_pptx_com_guards",
    "_is_slideshow_running": "_pptx_com_guards",
    "_ALLOWED_OUTPUT_EXTENSIONS": "_pptx_com_validators",
    "_VALID_SLIDE_FORMATS": "_pptx_com_validators",
    "_VALID_DECK_FORMATS": "_pptx_com_validators",
    "_validate_export_format": "_pptx_com_validators",
    "_validate_dpi": "_pptx_com_validators",
    "_check_in_allowlist": "_pptx_com_validators",
    "_check_output_file": "_pptx_com_validators",
    "_check_output_dir": "_pptx_com_validators",
    "_SLIDE_EXT_MAP": "_pptx_com_validators",
    "_check_output_path_slide": "_pptx_com_validators",
    "_DECK_EXT_MAP": "_pptx_com_validators",
    "_check_output_path_deck": "_pptx_com_validators",
    "_check_output_path": "_pptx_com_validators",
    "ppt_export_slide": "_pptx_com_export",
    "export_slide_as_png": "_pptx_com_export",
    "ppt_export_deck": "_pptx_com_export",
    "export_deck_as_pdf": "_pptx_com_export",
    "_make_stamped_export_dir": "_pptx_com_export",
    "export_slides_to_stamped_dir": "_pptx_com_export",
    "_MD5_HEX_RE": "_pptx_com_export",
    "_compute_slide_hash": "_pptx_com_export",
    "export_changed_slides_only": "_pptx_com_export",
}


def __getattr__(name: str):
    """PEP-562 fallback for any import not in explicit re-exports."""
    if name in _SUBMODULE_MAP:
        import importlib
        mod = importlib.import_module(f"pptmcp.{_SUBMODULE_MAP[name]}")
        return getattr(mod, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
