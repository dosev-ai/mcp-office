"""COM tool tests — targets presentation_com.py. Skipped on non-Windows."""

import contextlib
import sys
from unittest.mock import MagicMock

import pytest
from pptx import Presentation as _Prs
from pptmcp.presentation_pptx import (
    PPTMCPError,
    ValidationError,
    OfficeCOMError,
)
# pptmcp.presentation_com imports remain DEFERRED (inside class methods only)
# e.g. `from pptmcp.presentation_com import export_slide_as_png`
# This is intentional: deferred imports prevent import errors on non-Windows
# at collection time, correctly delegating to skipif at execution time.


@pytest.mark.skipif(sys.platform != "win32", reason="COM tests require Windows (win32com)")
@pytest.mark.unit
class TestPptComTools:
    """Unit tests for presentation_com: _check_com_enabled, _check_output_path, export_slide_as_png."""

    # ── shared helpers ────────────────────────────────────────────────────

    def _com_setup(self, monkeypatch, tmp_path):
        monkeypatch.setenv("PPT_ENABLE_COM", "true")
        monkeypatch.setenv("PPT_ENABLE_WRITE", "true")
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        monkeypatch.setitem(sys.modules, "win32com.client", MagicMock())
        mock_pythoncom = MagicMock()
        mock_pythoncom.com_error = Exception
        monkeypatch.setitem(sys.modules, "pythoncom", mock_pythoncom)

    def _make_pptx(self, tmp_path, slide_count=3):
        path = tmp_path / "src.pptx"
        p = _Prs()
        for _ in range(slide_count):
            p.slides.add_slide(p.slide_layouts[5])
        p.save(str(path))
        return str(path)

    def _make_com_mocks(self, slide_count=3):
        mock_slide_obj = MagicMock()
        mock_slides = MagicMock()
        mock_slides.Count = slide_count
        mock_slides.return_value = mock_slide_obj
        mock_prs_com = MagicMock()
        mock_prs_com.Slides = mock_slides
        mock_app = MagicMock()
        mock_app.Presentations.Open.return_value = mock_prs_com
        return mock_app, mock_slides, mock_slide_obj

    def _patch_app_context(self, monkeypatch, mock_app):
        @contextlib.contextmanager
        def _ctx(visible=False):
            try:
                yield mock_app
            finally:
                with contextlib.suppress(Exception):
                    mock_app.Quit()
        monkeypatch.setattr("pptmcp._pptx_com_export._ppt_app_context", _ctx)
        monkeypatch.setattr("pptmcp.presentation_com._ppt_app_context", _ctx)

    # ── _check_com_enabled ────────────────────────────────────────────────

    @pytest.mark.unit
    def test_check_com_enabled_no_env_var(self, monkeypatch):
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        from pptmcp.presentation_com import _check_com_enabled
        with pytest.raises(PPTMCPError):
            _check_com_enabled()

    @pytest.mark.unit
    def test_check_com_enabled_wrong_platform(self, monkeypatch):
        monkeypatch.setenv("PPT_ENABLE_COM", "true")
        monkeypatch.setattr(sys, "platform", "linux")
        from pptmcp.presentation_com import _check_com_enabled
        with pytest.raises(PPTMCPError):
            _check_com_enabled()

    @pytest.mark.unit
    def test_check_com_enabled_no_win32com(self, monkeypatch):
        monkeypatch.setenv("PPT_ENABLE_COM", "true")
        monkeypatch.setitem(sys.modules, "win32com.client", None)
        from pptmcp.presentation_com import _check_com_enabled
        with pytest.raises(PPTMCPError):
            _check_com_enabled()

    @pytest.mark.unit
    def test_check_com_enabled_ok(self, monkeypatch):
        monkeypatch.setenv("PPT_ENABLE_COM", "true")
        monkeypatch.setitem(sys.modules, "win32com.client", MagicMock())
        from pptmcp.presentation_com import _check_com_enabled
        _check_com_enabled()  # must not raise

    # ── _check_output_path ────────────────────────────────────────────────

    @pytest.mark.unit
    def test_check_output_path_null_byte(self, monkeypatch, tmp_path):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        from pptmcp.presentation_com import _check_output_path
        with pytest.raises(ValidationError):
            _check_output_path(str(tmp_path) + "/out\x00.png")

    @pytest.mark.unit
    def test_check_output_path_wrong_ext(self, monkeypatch, tmp_path):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        from pptmcp.presentation_com import _check_output_path
        with pytest.raises(ValidationError):
            _check_output_path(str(tmp_path / "out.txt"))

    @pytest.mark.unit
    def test_check_output_path_outside_allowlist(self, monkeypatch, tmp_path):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", "/nonexistent_XZY_root")
        from pptmcp.presentation_com import _check_output_path
        with pytest.raises(ValidationError):
            _check_output_path(str(tmp_path / "out.png"))

    @pytest.mark.unit
    def test_check_output_path_valid_png(self, monkeypatch, tmp_path):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        from pptmcp.presentation_com import _check_output_path
        result = _check_output_path(str(tmp_path / "out.png"))
        assert result.suffix == ".png"

    @pytest.mark.unit
    def test_check_output_path_valid_pdf(self, monkeypatch, tmp_path):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        from pptmcp.presentation_com import _check_output_path
        result = _check_output_path(str(tmp_path / "out.pdf"))
        assert result.suffix == ".pdf"

    # ── export_slide_as_png ───────────────────────────────────────────────

    @pytest.mark.unit
    def test_export_png_com_disabled(self, monkeypatch, tmp_path):
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(PPTMCPError):
            export_slide_as_png(str(tmp_path / "x.pptx"), 0, str(tmp_path / "out.png"), confirm=True)

    @pytest.mark.unit
    def test_export_png_outside_allowlist(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        # Restrict allowlist to a sub-directory; source pptx is in parent → blocked
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path / "sub"))
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(PPTMCPError):
            export_slide_as_png(
                str(tmp_path / "src.pptx"), 0,
                str(tmp_path / "sub" / "out.png"), confirm=True,
            )

    @pytest.mark.unit
    def test_export_png_slide_index_negative(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        mock_app, _, _ = self._make_com_mocks(slide_count=3)
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(ValidationError):
            export_slide_as_png(pptx, -1, str(tmp_path / "out.png"), confirm=True)

    @pytest.mark.unit
    def test_export_png_slide_index_too_large(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        mock_app, _, _ = self._make_com_mocks(slide_count=3)
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(ValidationError):
            export_slide_as_png(pptx, 5, str(tmp_path / "out.png"), confirm=True)

    @pytest.mark.unit
    def test_export_png_output_exists(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        out = tmp_path / "out.png"
        out.write_bytes(b"exists")
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(ValidationError):
            export_slide_as_png(pptx, 0, str(out), confirm=True)

    @pytest.mark.unit
    def test_export_png_happy_path(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        out = str(tmp_path / "out.png")
        mock_app, mock_slides, mock_slide_obj = self._make_com_mocks(slide_count=3)
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_slide_as_png
        result = export_slide_as_png(pptx, 0, out, confirm=True)
        assert result["exported"] is True
        mock_slides.assert_called_once_with(1)
        mock_slide_obj.Export.assert_called_once()
        mock_app.Quit.assert_called()

    @pytest.mark.unit
    def test_export_png_confirm_false(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(PPTMCPError):
            export_slide_as_png(
                str(tmp_path / "src.pptx"), 0, str(tmp_path / "out.png"), confirm=False,
            )

    @pytest.mark.unit
    def test_export_png_com_error(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        out = str(tmp_path / "out.png")
        mock_app, mock_slides, mock_slide_obj = self._make_com_mocks(slide_count=3)
        mock_slide_obj.Export.side_effect = Exception("COM explosion")
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(OfficeCOMError):
            export_slide_as_png(pptx, 0, out, confirm=True)

    @pytest.mark.unit
    def test_export_png_null_byte_output(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(ValidationError):
            export_slide_as_png(pptx, 0, str(tmp_path) + "/out\x00.png", confirm=True)

    @pytest.mark.unit
    def test_export_png_wrong_output_ext(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(ValidationError):
            export_slide_as_png(pptx, 0, str(tmp_path / "out.bmp"), confirm=True)

    # ── export_deck_as_pdf ────────────────────────────────────────────────

    @pytest.mark.unit
    def test_export_pdf_com_disabled(self, monkeypatch, tmp_path):
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        from pptmcp.presentation_com import export_deck_as_pdf
        with pytest.raises(PPTMCPError):
            export_deck_as_pdf(
                str(tmp_path / "x.pptx"), str(tmp_path / "out.pdf"), confirm=True
            )

    @pytest.mark.unit
    def test_export_pdf_outside_allowlist(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path / "sub"))
        from pptmcp.presentation_com import export_deck_as_pdf
        with pytest.raises(PPTMCPError):
            export_deck_as_pdf(
                str(tmp_path / "src.pptx"),
                str(tmp_path / "sub" / "out.pdf"),
                confirm=True,
            )

    @pytest.mark.unit
    def test_export_pdf_output_exists(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        out = tmp_path / "out.pdf"
        out.write_bytes(b"exists")
        from pptmcp.presentation_com import export_deck_as_pdf
        with pytest.raises(ValidationError):
            export_deck_as_pdf(pptx, str(out), confirm=True)

    @pytest.mark.unit
    def test_export_pdf_happy_path(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        out = str(tmp_path / "out.pdf")
        mock_app, _, _ = self._make_com_mocks(slide_count=2)
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_deck_as_pdf
        result = export_deck_as_pdf(pptx, out, confirm=True)
        assert result["exported"] is True
        prs_com = mock_app.Presentations.Open.return_value
        call_args = prs_com.ExportAsFixedFormat.call_args
        assert call_args[0][1] == 2  # second positional arg = ppFixedFormatTypePDF
        assert call_args[0][12] == 0  # COM-B2/B-7: IncludeDocProperties at positional index 12
        mock_app.Quit.assert_called()

    @pytest.mark.unit
    def test_export_pdf_confirm_false(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        from pptmcp.presentation_com import export_deck_as_pdf
        with pytest.raises(PPTMCPError):
            export_deck_as_pdf(
                str(tmp_path / "src.pptx"), str(tmp_path / "out.pdf"), confirm=False
            )

    @pytest.mark.unit
    def test_export_pdf_com_error(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        out = str(tmp_path / "out.pdf")
        mock_app, _, _ = self._make_com_mocks(slide_count=2)
        prs_com = mock_app.Presentations.Open.return_value
        prs_com.ExportAsFixedFormat.side_effect = Exception("COM boom")
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_deck_as_pdf
        with pytest.raises(OfficeCOMError):
            export_deck_as_pdf(pptx, out, confirm=True)

    @pytest.mark.unit
    def test_export_pdf_null_byte_output(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        from pptmcp.presentation_com import export_deck_as_pdf
        with pytest.raises(ValidationError):
            export_deck_as_pdf(pptx, str(tmp_path) + "/out\x00.pdf", confirm=True)

    @pytest.mark.unit
    def test_export_pdf_wrong_ext(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        from pptmcp.presentation_com import export_deck_as_pdf
        with pytest.raises(ValidationError):
            export_deck_as_pdf(pptx, str(tmp_path / "out.csv"), confirm=True)

    # ── run_slide_show ────────────────────────────────────────────────────

    @pytest.mark.unit
    def test_run_slideshow_com_disabled(self, monkeypatch, tmp_path):
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        from pptmcp.presentation_com import run_slide_show
        with pytest.raises(PPTMCPError):
            run_slide_show(str(tmp_path / "x.pptx"))

    @pytest.mark.unit
    def test_run_slideshow_outside_allowlist(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path / "sub"))
        from pptmcp.presentation_com import run_slide_show
        with pytest.raises(PPTMCPError):
            run_slide_show(str(tmp_path / "src.pptx"))

    @pytest.mark.unit
    def test_run_slideshow_already_running(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        monkeypatch.setattr(
            "pptmcp.presentation_com._is_slideshow_running", lambda _: True
        )
        from pptmcp.presentation_com import run_slide_show
        with pytest.raises(ValidationError, match="already running"):
            run_slide_show(pptx, confirm=True)

    @pytest.mark.unit
    def test_run_slideshow_happy_path(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        monkeypatch.setattr(
            "pptmcp.presentation_com._is_slideshow_running", lambda _: False
        )
        # Set up win32com.client.Dispatch → mock_app
        mock_win32com_client = sys.modules["win32com.client"]
        mock_win32com = MagicMock()
        mock_win32com.client = mock_win32com_client
        monkeypatch.setitem(sys.modules, "win32com", mock_win32com)
        mock_app = MagicMock()
        mock_prs_obj = MagicMock()
        mock_prs_obj.Slides.Count = 3
        mock_app.Presentations.Open.return_value = mock_prs_obj
        mock_win32com_client.Dispatch.return_value = mock_app
        from pptmcp.presentation_com import run_slide_show
        result = run_slide_show(pptx, confirm=True)
        assert result["running"] is True
        assert result["slide_count"] == 3
        # B-6: Saved must be set True immediately after Open
        assert mock_prs_obj.Saved is True
        # B-1: SlideShowSettings.Run() must be called (not 'with' block)
        mock_prs_obj.SlideShowSettings.Run.assert_called_once()
        # app.Quit must NOT be called — window stays open for the slide show
        mock_app.Quit.assert_not_called()

    @pytest.mark.unit
    def test_run_slideshow_com_open_error(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        monkeypatch.setattr(
            "pptmcp.presentation_com._is_slideshow_running", lambda _: False
        )
        mock_win32com_client = sys.modules["win32com.client"]
        mock_win32com = MagicMock()
        mock_win32com.client = mock_win32com_client
        monkeypatch.setitem(sys.modules, "win32com", mock_win32com)
        mock_app = MagicMock()
        mock_app.Presentations.Open.side_effect = Exception("COM boom")
        mock_win32com_client.Dispatch.return_value = mock_app
        from pptmcp.presentation_com import run_slide_show
        with pytest.raises(OfficeCOMError):
            run_slide_show(pptx, confirm=True)

    # ── recalculate_charts ────────────────────────────────────────────────

    @pytest.mark.unit
    def test_recalc_charts_com_disabled(self, monkeypatch, tmp_path):
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        from pptmcp.presentation_com import recalculate_charts
        with pytest.raises(PPTMCPError):
            recalculate_charts(str(tmp_path / "x.pptx"), confirm=True)

    @pytest.mark.unit
    def test_recalc_charts_outside_allowlist(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path / "sub"))
        from pptmcp.presentation_com import recalculate_charts
        with pytest.raises(PPTMCPError):
            recalculate_charts(str(tmp_path / "src.pptx"), confirm=True)

    @pytest.mark.unit
    def test_recalc_charts_confirm_false(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        from pptmcp.presentation_com import recalculate_charts
        with pytest.raises(PPTMCPError):
            recalculate_charts(str(tmp_path / "ok.pptx"), confirm=False)

    @pytest.mark.unit
    def test_recalc_charts_happy_path(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)

        # shape1: native chart (Type=3), not linked → should be activated
        shape1 = MagicMock()
        shape1.Type = 3
        shape1.Chart.ChartData.IsLinked = False

        # shape2: native chart (Type=3), linked → should be skipped
        shape2 = MagicMock()
        shape2.Type = 3
        shape2.Chart.ChartData.IsLinked = True

        mock_slide = MagicMock()
        mock_slide.Shapes.Count = 2
        mock_slide.Shapes.side_effect = lambda i: shape1 if i == 1 else shape2

        mock_prs_com = MagicMock()
        mock_prs_com.Slides.Count = 1
        mock_prs_com.Slides.side_effect = lambda i: mock_slide

        security_at_open: list = []
        mock_app = MagicMock()

        def _mock_open(*args, **kwargs):
            security_at_open.append(mock_app.AutomationSecurity)
            return mock_prs_com

        mock_app.Presentations.Open.side_effect = _mock_open
        self._patch_app_context(monkeypatch, mock_app)

        from pptmcp.presentation_com import recalculate_charts
        result = recalculate_charts(pptx, confirm=True)

        assert result["recalculated"] is True
        assert result["charts_activated"] == 1
        assert result["skipped_linked_charts"] == 1
        # B-2: AutomationSecurity=3 must be set BEFORE app.Presentations.Open
        assert security_at_open == [3]
        shape1.Chart.ChartData.Activate.assert_called_once()
        shape2.Chart.ChartData.Activate.assert_not_called()

    @pytest.mark.unit
    def test_recalc_charts_com_error(self, monkeypatch, tmp_path):
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        mock_app = MagicMock()
        mock_app.Presentations.Open.side_effect = Exception("COM boom")
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import recalculate_charts
        with pytest.raises(OfficeCOMError):
            recalculate_charts(pptx, confirm=True)

    # ── US-09: export_slide_as_png resolution params ──────────────────────

    @pytest.mark.unit
    def test_export_slide_as_png_accepts_resolution_params(self, monkeypatch, tmp_path):
        """US-09: width_px and height_px are passed to COM Slide.Export."""
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path, slide_count=1)
        out = str(tmp_path / "out.png")
        mock_app, mock_slides, mock_slide_obj = self._make_com_mocks(slide_count=1)
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_slide_as_png
        result = export_slide_as_png(pptx, 0, out, width_px=1920, height_px=1080, confirm=True)
        assert result["exported"] is True
        assert result["width_px"] == 1920
        assert result["height_px"] == 1080
        args = mock_slide_obj.Export.call_args[0]
        assert args[1] == "PNG"
        assert args[2] == 1920
        assert args[3] == 1080

    @pytest.mark.unit
    def test_export_slide_as_png_rejects_oversized_width(self, tmp_path):
        """US-09: width_px > 7680 raises ValidationError (fires before COM)."""
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(ValidationError):
            export_slide_as_png(
                str(tmp_path / "src.pptx"), 0, str(tmp_path / "out.png"),
                width_px=9999, height_px=1080, confirm=True,
            )

    @pytest.mark.unit
    def test_export_slide_as_png_rejects_oversized_height(self, tmp_path):
        """US-09: height_px > 7680 raises ValidationError (fires before COM)."""
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(ValidationError):
            export_slide_as_png(
                str(tmp_path / "src.pptx"), 0, str(tmp_path / "out.png"),
                width_px=1920, height_px=9999, confirm=True,
            )

    # ── US-10: export_slides_to_stamped_dir ───────────────────────────────

    @pytest.mark.unit
    def test_make_stamped_export_dir_creates_path(self, tmp_path):
        """_make_stamped_export_dir creates exports/run-<ts>/png/ subdir."""
        from pptmcp.presentation_com import _make_stamped_export_dir
        result = _make_stamped_export_dir(str(tmp_path))
        assert result.exists()
        assert result.name == "png"
        assert result.parent.name.startswith("run-")
        assert result.parent.parent.name == "exports"

    @pytest.mark.unit
    def test_export_slides_to_stamped_dir_creates_expected_structure(self, monkeypatch, tmp_path):
        """US-10: run_dir matches exports/run-<timestamp>/png/ pattern; slide-000.png produced."""
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path, slide_count=2)
        mock_app, mock_slides, mock_slide_obj = self._make_com_mocks(slide_count=2)
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_slides_to_stamped_dir
        result = export_slides_to_stamped_dir(pptx, str(tmp_path), slide_indices=[0], confirm=True)
        from pathlib import Path as _Path
        run_dir_path = _Path(result["run_dir"])
        assert run_dir_path.name == "png"
        assert run_dir_path.parent.name.startswith("run-")
        assert run_dir_path.parent.parent.name == "exports"
        assert len(result["exported_slides"]) == 1
        assert "slide-000.png" in result["exported_slides"][0]
        assert result["slide_count"] == 1

    @pytest.mark.unit
    def test_export_slides_to_stamped_dir_requires_confirm(self, monkeypatch, tmp_path):
        """US-10: confirm=False raises PPTMCPError."""
        self._com_setup(monkeypatch, tmp_path)
        from pptmcp.presentation_com import export_slides_to_stamped_dir
        with pytest.raises(PPTMCPError):
            export_slides_to_stamped_dir(
                str(tmp_path / "src.pptx"), str(tmp_path), confirm=False
            )

    @pytest.mark.unit
    def test_export_slides_to_stamped_dir_requires_com(self, monkeypatch):
        """US-10: without PPT_ENABLE_COM=true, raises PPTMCPError."""
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        from pptmcp.presentation_com import export_slides_to_stamped_dir
        with pytest.raises(PPTMCPError):
            export_slides_to_stamped_dir("/tmp/x.pptx", "/tmp", confirm=True)

    @pytest.mark.unit
    def test_export_slides_to_stamped_dir_write_disabled(self, monkeypatch, tmp_path):
        """M-001: PPT_ENABLE_WRITE=false raises PPTMCPError for export_slides_to_stamped_dir."""
        self._com_setup(monkeypatch, tmp_path)
        monkeypatch.setenv("PPT_ENABLE_WRITE", "false")
        from pptmcp.presentation_com import export_slides_to_stamped_dir
        with pytest.raises(PPTMCPError):
            export_slides_to_stamped_dir(
                str(tmp_path / "src.pptx"), str(tmp_path), confirm=True
            )

    @pytest.mark.unit
    def test_export_slides_to_stamped_dir_all_slides(self, monkeypatch, tmp_path):
        """M-003: slide_indices=None exports all slides; filenames follow slide-NNN.png pattern."""
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path, slide_count=2)
        mock_app, mock_slides, mock_slide_obj = self._make_com_mocks(slide_count=2)
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_slides_to_stamped_dir
        result = export_slides_to_stamped_dir(pptx, str(tmp_path), slide_indices=None, confirm=True)
        assert result["slide_count"] == 2
        assert len(result["exported_slides"]) == 2
        assert "slide-000.png" in result["exported_slides"][0]
        assert "slide-001.png" in result["exported_slides"][1]

    @pytest.mark.unit
    def test_export_slides_to_stamped_null_byte_in_base_dir(self, monkeypatch, tmp_path):
        """L-001a: base_dir with null byte raises ValidationError."""
        self._com_setup(monkeypatch, tmp_path)
        from pptmcp.presentation_com import export_slides_to_stamped_dir
        with pytest.raises((ValidationError, PPTMCPError)):
            export_slides_to_stamped_dir(
                str(tmp_path / "src.pptx"), str(tmp_path) + "/ex\x00ports", confirm=True
            )

    @pytest.mark.unit
    def test_export_slides_to_stamped_base_dir_outside_allowlist(self, monkeypatch, tmp_path):
        """L-001b: base_dir outside PPT_ALLOWLIST_ROOTS raises ValidationError."""
        self._com_setup(monkeypatch, tmp_path)
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path / "allowed"))
        from pptmcp.presentation_com import export_slides_to_stamped_dir
        with pytest.raises((ValidationError, PPTMCPError)):
            export_slides_to_stamped_dir(
                str(tmp_path / "src.pptx"), str(tmp_path / "other"), confirm=True
            )

    @pytest.mark.unit
    def test_export_slides_to_stamped_oob_slide_index(self, monkeypatch, tmp_path):
        """L-001c: slide_indices with out-of-range index raises ValidationError."""
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path, slide_count=2)
        mock_app, mock_slides, mock_slide_obj = self._make_com_mocks(slide_count=2)
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_slides_to_stamped_dir
        with pytest.raises((ValidationError, PPTMCPError)):
            export_slides_to_stamped_dir(pptx, str(tmp_path), slide_indices=[0, 99], confirm=True)

    @pytest.mark.unit
    def test_export_png_rejects_mismatched_px_dimensions(self, monkeypatch, tmp_path):
        """L-003: width_px=1920, height_px=0 raises ValidationError (mismatch)."""
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path)
        from pptmcp.presentation_com import export_slide_as_png
        with pytest.raises(ValidationError):
            export_slide_as_png(pptx, 0, str(tmp_path / "out.png"), width_px=1920, height_px=0, confirm=True)

    # ── US-11: export_deck_as_pdf page_count ─────────────────────────────

    @pytest.mark.unit
    def test_export_deck_as_pdf_returns_page_count(self, monkeypatch, tmp_path):
        """US-11: export_deck_as_pdf return dict contains page_count == slide_count."""
        self._com_setup(monkeypatch, tmp_path)
        pptx = self._make_pptx(tmp_path, slide_count=3)
        out = str(tmp_path / "out.pdf")
        mock_app, _, _ = self._make_com_mocks(slide_count=3)
        self._patch_app_context(monkeypatch, mock_app)
        from pptmcp.presentation_com import export_deck_as_pdf
        result = export_deck_as_pdf(pptx, out, confirm=True)
        assert result["exported"] is True
        assert "page_count" in result
        assert result["page_count"] == result["slide_count"]

    # ── US-12: COM disabled error message ─────────────────────────────────

    @pytest.mark.unit
    def test_com_disabled_error_message_is_agent_actionable(self, monkeypatch):
        """US-12: _check_com_enabled() raises with agent-actionable message."""
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        from pptmcp.presentation_com import _check_com_enabled
        with pytest.raises(Exception) as exc_info:
            _check_com_enabled()
        assert "PPT_ENABLE_COM=true" in str(exc_info.value)
