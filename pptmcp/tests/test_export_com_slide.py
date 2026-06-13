"""Tests for ppt_export_slide, ppt_export_deck, _validate_export_format,
_validate_dpi, _check_output_path_slide, _check_output_path_deck,
and deprecated server.py aliases.

All COM calls are mocked — no live PowerPoint required.
TestDeprecatedAliases mocks _pcom; tests skip automatically on platforms where
_COM_AVAILABLE=False (i.e. non-Windows or missing win32com).
Run: pytest powerpoint/tests/test_export_com.py -v
"""
import contextlib
import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from pptx import Presentation as _Prs

from pptmcp.presentation_com import (
    ppt_export_slide,
)
from pptmcp.presentation_pptx import NotAllowedError, OfficeCOMError, ValidationError


# ─────────────────────────────────────────────────────────────────────────────
# Shared fixtures
# ─────────────────────────────────────────────────────────────────────────────


@pytest.fixture
def com_env(tmp_path, monkeypatch):
    """COM-ready environment: env vars + sys.modules mocks + real .pptx in tmp_path."""
    monkeypatch.setenv("PPT_ENABLE_COM", "true")
    monkeypatch.setenv("PPT_ENABLE_WRITE", "true")
    monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))

    # Suppress the Windows-platform check so tests run on Linux/macOS CI
    _mock_check_com = MagicMock()
    monkeypatch.setattr(
        "pptmcp._pptx_com_export._check_com_enabled", _mock_check_com
    )
    monkeypatch.setattr(
        "pptmcp.presentation_com._check_com_enabled", _mock_check_com
    )

    mock_pythoncom = MagicMock()
    mock_pythoncom.com_error = Exception
    # Mock both win32com (parent package) AND win32com.client so that
    # `import win32com.client` inside _ppt_app_context succeeds on Linux CI.
    # CPython returns parent from sys.modules immediately without setting
    # parent.client = child, so we must link them explicitly; that way
    # monkeypatch.setattr(sys.modules["win32com.client"], "Dispatch", ...)
    # reaches the same object that _ppt_app_context accesses as
    # win32com.client.Dispatch after the import.
    mock_win32com_client = MagicMock()
    mock_win32com = MagicMock()
    mock_win32com.client = mock_win32com_client  # explicit link — critical!
    monkeypatch.setitem(sys.modules, "win32com", mock_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", mock_win32com_client)
    monkeypatch.setitem(sys.modules, "pythoncom", mock_pythoncom)

    prs = _Prs()
    for _ in range(3):
        prs.slides.add_slide(prs.slide_layouts[5])
    pptx_path = tmp_path / "test.pptx"
    prs.save(str(pptx_path))

    return {"pptx": str(pptx_path), "tmp": tmp_path}


@pytest.fixture
def com_mocks(com_env, monkeypatch):
    """Full COM mock setup: patched _ppt_app_context + mock app/prs/slide."""
    mock_slide = MagicMock()
    mock_prs = MagicMock()
    mock_prs.Slides.Count = 3
    mock_prs.Slides.return_value = mock_slide
    mock_app = MagicMock()
    mock_app.Presentations.Open.return_value = mock_prs

    @contextlib.contextmanager
    def _fake_ctx(visible=False):
        yield mock_app

    monkeypatch.setattr("pptmcp._pptx_com_export._ppt_app_context", _fake_ctx)
    monkeypatch.setattr("pptmcp.presentation_com._ppt_app_context", _fake_ctx)

    return {
        "app": mock_app,
        "prs": mock_prs,
        "slide": mock_slide,
        "pptx": com_env["pptx"],
        "tmp": com_env["tmp"],
    }


# ─────────────────────────────────────────────────────────────────────────────
# TestValidateExportFormat
# ─────────────────────────────────────────────────────────────────────────────


class TestPptExportSlide:

    # ── helpers ──────────────────────────────────────────────────────────────

    def _out(self, com_mocks, name="out.png"):
        return str(Path(com_mocks["pptx"]).parent / name)

    # ── happy paths ───────────────────────────────────────────────────────────

    def test_png_happy_path(self, com_mocks):
        out = self._out(com_mocks, "slide1.png")
        result = ppt_export_slide(com_mocks["pptx"], 0, "png", out, confirm=True)
        assert result["exported"] is True
        assert result["format"] == "png"
        assert result["slide_number"] == 1
        com_mocks["prs"].Slides.assert_called_with(1)
        com_mocks["slide"].Export.assert_called()

    def test_pdf_happy_path(self, com_mocks):
        out = self._out(com_mocks, "slide1.pdf")
        result = ppt_export_slide(com_mocks["pptx"], 1, "pdf", out, confirm=True)
        assert result["exported"] is True
        assert result["format"] == "pdf"
        com_mocks["prs"].ExportAsFixedFormat.assert_called()

    def test_pdf_single_slide_uses_all_positional_args(self, com_mocks):
        # Regression for pywin32 cannot mix positional
        # args with kwargs around a skipped IDispatch slot (PrintRange).
        # The fix uses all-positional args with None at pos 7. This test
        # asserts the call signature so the SlideStart kwarg pattern cannot
        # silently come back.
        out = self._out(com_mocks, "slide_pos.pdf")
        ppt_export_slide(com_mocks["pptx"], 1, "pdf", out, confirm=True)
        call = com_mocks["prs"].ExportAsFixedFormat.call_args
        assert call.kwargs == {}, (
            f"single-slide PDF must use all-positional args; got kwargs={call.kwargs}"
        )
        args = call.args
        assert len(args) == 13, f"expected 13 positional args, got {len(args)}"
        # pos  1: ppFixedFormatTypePDF=2
        # pos  7: PrintRange=None (Nothing)
        # pos  8: RangeType=4 (ppPrintSlideRange)
        # pos  9: SlideStart=slide_index+1
        # pos 10: SlideEnd=slide_index+1
        # pos 12: IncludeDocProperties=0 (B-7)
        assert args[1] == 2
        assert args[7] is None
        assert args[8] == 4
        assert args[9] == 2  # slide_index 1 → SlideStart 2
        assert args[10] == 2  # SlideEnd same as start for single slide
        assert args[12] == 0

    def test_svg_happy_path(self, com_mocks):
        out = self._out(com_mocks, "slide1.svg")
        result = ppt_export_slide(com_mocks["pptx"], 0, "svg", out, confirm=True)
        assert result["exported"] is True
        assert result["format"] == "svg"

    def test_svg_office_2016_com_error(self, com_mocks):
        """SVG export raises com_error on Office 2016 → OfficeCOMError with upgrade hint."""
        out = self._out(com_mocks, "fail.svg")
        # pythoncom.com_error is mocked as Exception in com_env fixture
        com_mocks["slide"].Export.side_effect = Exception("SVG not supported by this Office")
        with pytest.raises(OfficeCOMError, match="2019|365"):
            ppt_export_slide(com_mocks["pptx"], 0, "svg", out, confirm=True)

    # ── slide bounds ──────────────────────────────────────────────────────────

    def test_slide_index_negative_raises(self, com_mocks):
        out = self._out(com_mocks, "out_neg.png")
        with pytest.raises(ValidationError, match="out of range"):
            ppt_export_slide(com_mocks["pptx"], -1, "png", out, confirm=True)

    def test_slide_number_out_of_range_raises(self, com_mocks):
        out = self._out(com_mocks, "out99.png")
        with pytest.raises(ValidationError, match="out of range"):
            ppt_export_slide(com_mocks["pptx"], 99, "png", out, confirm=True)

    # ── format validation ─────────────────────────────────────────────────────

    def test_invalid_format_raises(self, com_mocks):
        out = self._out(com_mocks, "out.gif")
        with pytest.raises(ValidationError):
            ppt_export_slide(com_mocks["pptx"], 0, "gif", out, confirm=True)

    # ── DPI ───────────────────────────────────────────────────────────────────

    def test_dpi_applied_to_png(self, com_mocks):
        """dpi=150 → Export called with width=1500, height=1125."""
        out = self._out(com_mocks, "dpi.png")
        ppt_export_slide(com_mocks["pptx"], 0, "png", out, dpi=150, confirm=True)
        expected_out = str(Path(out).resolve())
        com_mocks["slide"].Export.assert_called_once_with(
            expected_out, "PNG", 1500, 1125
        )

    # ── security / guard paths ────────────────────────────────────────────────

    def test_confirm_false_raises(self, com_env):
        out = str(Path(com_env["pptx"]).parent / "noconfirm.png")
        with pytest.raises(ValidationError):
            ppt_export_slide(com_env["pptx"], 0, "png", out, confirm=False)

    def test_output_path_traversal_blocked(self, com_env):
        outside = str(Path(com_env["tmp"]).parent / "traversal.png")
        with pytest.raises(ValidationError):
            ppt_export_slide(com_env["pptx"], 0, "png", outside, confirm=True)

    def test_com_not_available_guard(self, tmp_path, monkeypatch):
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        # Allow path through _check_path so the COM gate is actually exercised
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        with pytest.raises(NotAllowedError):
            ppt_export_slide(
                str(tmp_path / "x.pptx"), 0, "png",
                str(tmp_path / "out.png"), confirm=True,
            )

    # ── return shape ──────────────────────────────────────────────────────────

    def test_returns_success_dict(self, com_mocks):
        out = self._out(com_mocks, "retval.png")
        result = ppt_export_slide(com_mocks["pptx"], 0, "png", out, confirm=True)
        assert result == {
            "exported": True,
            "slide_number": 1,
            "format": "png",
            "output_path": str(Path(out).resolve()),
        }

    # ── US-09: explicit pixel dimensions ─────────────────────────────────────

    def test_explicit_width_height_png(self, com_mocks):
        """US-09: explicit width_px/height_px are forwarded to slide.Export."""
        out = self._out(com_mocks, "explicit_dims.png")
        ppt_export_slide(
            com_mocks["pptx"], 0, "png", out,
            width_px=1920, height_px=1080, confirm=True,
        )
        expected_out = str(Path(out).resolve())
        com_mocks["slide"].Export.assert_called_once_with(
            expected_out, "PNG", 1920, 1080
        )

    def test_width_px_boundary_max_passes(self, com_mocks):
        """US-09: width_px=7680, height_px=4320 are at the upper boundary — must not raise."""
        out = self._out(com_mocks, "maxdims.png")
        ppt_export_slide(
            com_mocks["pptx"], 0, "png", out,
            width_px=7680, height_px=4320, confirm=True,
        )
        expected_out = str(Path(out).resolve())
        com_mocks["slide"].Export.assert_called_once_with(
            expected_out, "PNG", 7680, 4320
        )

    def test_width_px_over_boundary_raises(self, com_mocks):
        """US-09: width_px=7681 exceeds the 0–7680 bound → ValidationError."""
        out = self._out(com_mocks, "oob.png")
        with pytest.raises(ValidationError):
            ppt_export_slide(
                com_mocks["pptx"], 0, "png", out,
                width_px=7681, height_px=0, confirm=True,
            )

    def test_height_px_negative_raises(self, com_mocks):
        """US-09: height_px=-1 is negative → ValidationError."""
        out = self._out(com_mocks, "neg_h.png")
        with pytest.raises(ValidationError):
            ppt_export_slide(
                com_mocks["pptx"], 0, "png", out,
                width_px=0, height_px=-1, confirm=True,
            )

    def test_partial_pixel_spec_raises(self, com_mocks):
        """US-09: width_px set but height_px=0 → ValidationError (must both be set or both 0)."""
        out = self._out(com_mocks, "partial.png")
        with pytest.raises(ValidationError):
            ppt_export_slide(
                com_mocks["pptx"], 0, "png", out,
                width_px=1920, height_px=0, confirm=True,
            )

    def test_pixel_dims_on_non_png_raises(self, com_mocks):
        """US-09: width_px/height_px are only valid when fmt='png'."""
        out = self._out(com_mocks, "out.svg")
        with pytest.raises(ValidationError):
            ppt_export_slide(
                com_mocks["pptx"], 0, "svg", out,
                width_px=1920, height_px=1080, confirm=True,
            )

    def test_defaults_zero_preserve_dpi_path(self, com_mocks):
        """US-09: dpi=150, width_px=0, height_px=0 → DPI path still used: Export(w=1500, h=1125)."""
        out = self._out(com_mocks, "dpi_zero_px.png")
        ppt_export_slide(
            com_mocks["pptx"], 0, "png", out,
            dpi=150, width_px=0, height_px=0, confirm=True,
        )
        expected_out = str(Path(out).resolve())
        com_mocks["slide"].Export.assert_called_once_with(
            expected_out, "PNG", 1500, 1125
        )

    def test_dpi_overridden_by_pixel_dims_warning(self, com_mocks, caplog):
        """US-09: dpi + pixel dims both supplied → pixel wins; warning logged."""
        import logging
        out = self._out(com_mocks, "px_wins.png")
        with caplog.at_level(logging.WARNING, logger="pptmcp.presentation_com"):
            ppt_export_slide(
                com_mocks["pptx"], 0, "png", out,
                dpi=150, width_px=1920, height_px=1080, confirm=True,
            )
        expected_out = str(Path(out).resolve())
        com_mocks["slide"].Export.assert_called_once_with(
            expected_out, "PNG", 1920, 1080
        )
        assert any(
            "dpi" in r.message.lower() and "ignored" in r.message.lower()
            for r in caplog.records
        )


# ─────────────────────────────────────────────────────────────────────────────
# TestPptExportDeck
# ─────────────────────────────────────────────────────────────────────────────


