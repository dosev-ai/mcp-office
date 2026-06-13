"""Tests for ppt_export_slide, ppt_export_deck, _validate_export_format,
_validate_dpi, _check_output_path_slide, _check_output_path_deck,
and deprecated server.py aliases.

All COM calls are mocked — no live PowerPoint required.
TestDeprecatedAliases mocks _pcom; tests skip automatically on platforms where
_COM_AVAILABLE=False (i.e. non-Windows or missing win32com).
Run: pytest powerpoint/tests/test_export_com.py -v
"""
import contextlib
import re
import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from pptx import Presentation as _Prs

from pptmcp.presentation_com import (
    export_slides_to_stamped_dir,
    ppt_export_deck,
)
from pptmcp.presentation_pptx import NotAllowedError, ValidationError


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


class TestPptExportDeck:

    # ── helpers ──────────────────────────────────────────────────────────────

    def _out(self, com_mocks, name):
        return str(Path(com_mocks["pptx"]).parent / name)

    # ── happy paths ───────────────────────────────────────────────────────────

    def test_pdf_happy_path(self, com_mocks):
        out = self._out(com_mocks, "deck.pdf")
        result = ppt_export_deck(com_mocks["pptx"], "pdf", out, confirm=True)
        assert result["exported"] is True
        assert result["format"] == "pdf"
        assert result["slide_count"] == 3
        com_mocks["prs"].ExportAsFixedFormat.assert_called()

    def test_pptx_happy_path(self, com_mocks):
        out = self._out(com_mocks, "copy.pptx")
        result = ppt_export_deck(com_mocks["pptx"], "pptx", out, confirm=True)
        assert result["exported"] is True
        assert result["format"] == "pptx"
        assert result["slide_count"] == 3
        com_mocks["prs"].SaveCopyAs.assert_called()

    def test_images_happy_path(self, com_mocks):
        out = self._out(com_mocks, "imgs")
        result = ppt_export_deck(com_mocks["pptx"], "images", out, confirm=True)
        assert result["exported"] is True
        assert result["format"] == "images"
        assert "output_dir" in result
        assert "files" in result
        assert len(result["files"]) == 3
        com_mocks["prs"].Export.assert_called()

    # ── format validation ─────────────────────────────────────────────────────

    def test_invalid_format_raises(self, com_mocks):
        out = self._out(com_mocks, "bad.zip")
        with pytest.raises(ValidationError):
            ppt_export_deck(com_mocks["pptx"], "zip", out, confirm=True)

    # ── M-4 same-path guard ───────────────────────────────────────────────────

    def test_pptx_same_path_raises(self, com_env):
        """M-4: output path same as source raises ValidationError before any COM call."""
        with pytest.raises(ValidationError, match="same"):
            ppt_export_deck(com_env["pptx"], "pptx", com_env["pptx"], confirm=True)

    # ── DPI ───────────────────────────────────────────────────────────────────

    def test_dpi_applied_to_images(self, com_mocks):
        """dpi=120 → prs.Export called with width=1200, height=900."""
        out = self._out(com_mocks, "dpi_imgs")
        ppt_export_deck(com_mocks["pptx"], "images", out, dpi=120, confirm=True)
        expected_out = str(Path(out).resolve())
        com_mocks["prs"].Export.assert_called_once_with(
            expected_out, "PNG", 1200, 900
        )

    def test_dpi_ignored_for_pdf(self, com_mocks):
        """dpi param is warned and ignored for pdf format (no ValidationError)."""
        out = self._out(com_mocks, "pdf_dpi.pdf")
        ppt_export_deck(com_mocks["pptx"], "pdf", out, dpi=150, confirm=True)
        com_mocks["prs"].ExportAsFixedFormat.assert_called()

    # ── security / guard paths ────────────────────────────────────────────────

    def test_confirm_false_raises(self, com_env):
        out = str(Path(com_env["pptx"]).parent / "noconfirm.pdf")
        with pytest.raises(ValidationError):
            ppt_export_deck(com_env["pptx"], "pdf", out, confirm=False)

    def test_output_path_traversal_blocked(self, com_env):
        outside = str(Path(com_env["tmp"]).parent / "traversal.pdf")
        with pytest.raises(ValidationError):
            ppt_export_deck(com_env["pptx"], "pdf", outside, confirm=True)

    def test_com_not_available_guard(self, tmp_path, monkeypatch):
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        # Allow path through _check_path so the COM gate is actually exercised
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        with pytest.raises(NotAllowedError):
            ppt_export_deck(
                str(tmp_path / "x.pptx"), "pdf",
                str(tmp_path / "out.pdf"), confirm=True,
            )

    # ── return shape ──────────────────────────────────────────────────────────

    def test_returns_success_dict(self, com_mocks):
        out = self._out(com_mocks, "retval.pdf")
        result = ppt_export_deck(com_mocks["pptx"], "pdf", out, confirm=True)
        assert result["exported"] is True
        assert "format" in result
        assert "slide_count" in result
        assert "output_path" in result

    # ── directory creation ────────────────────────────────────────────────────

    def test_images_dir_created(self, com_mocks):
        """ppt_export_deck creates the output directory before the COM call."""
        out = Path(com_mocks["pptx"]).parent / "new_dir"
        assert not out.exists()
        ppt_export_deck(com_mocks["pptx"], "images", str(out), confirm=True)
        assert out.is_dir()

    # ── US-11: page_count in PDF result ──────────────────────────────────────

    def test_pdf_return_contains_page_count(self, com_mocks):
        """US-11: PDF export result contains 'page_count' key."""
        out = self._out(com_mocks, "pc_present.pdf")
        result = ppt_export_deck(com_mocks["pptx"], "pdf", out, confirm=True)
        assert "page_count" in result

    def test_pdf_page_count_equals_slide_count(self, com_mocks):
        """US-11: page_count equals slide_count (one PDF page per slide)."""
        out = self._out(com_mocks, "pc_eq_sc.pdf")
        result = ppt_export_deck(com_mocks["pptx"], "pdf", out, confirm=True)
        assert result["page_count"] == result["slide_count"] == 3

    def test_page_count_absent_for_pptx(self, com_mocks):
        """US-11: page_count is NOT present when exporting as pptx."""
        out = self._out(com_mocks, "no_pc.pptx")
        result = ppt_export_deck(com_mocks["pptx"], "pptx", out, confirm=True)
        assert "page_count" not in result


# ─────────────────────────────────────────────────────────────────────────────
# TestPptExportSlidesToStampedDir
# ─────────────────────────────────────────────────────────────────────────────


class TestPptExportSlidesToStampedDir:
    """US-10: Export selected slides to a timestamped PNG directory."""

    def test_all_slides_happy_path(self, com_mocks):
        """US-10: all 3 slides exported; result has slide_count=3 and 3 .png file paths."""
        result = export_slides_to_stamped_dir(
            com_mocks["pptx"],
            str(com_mocks["tmp"]),
            slide_indices=None,
            width_px=0,
            height_px=0,
            confirm=True,
        )
        assert result["slide_count"] == 3
        assert len(result["exported_slides"]) == 3
        assert all(f.endswith(".png") for f in result["exported_slides"])
        assert com_mocks["slide"].Export.call_count == 3

    def test_specific_slide_indices(self, com_mocks):
        """US-10: only slides 0 and 2 exported; filenames are slide-000.png and slide-002.png."""
        result = export_slides_to_stamped_dir(
            com_mocks["pptx"],
            str(com_mocks["tmp"]),
            slide_indices=[0, 2],
            confirm=True,
        )
        assert result["slide_count"] == 2
        assert len(result["exported_slides"]) == 2
        filenames = [Path(f).name for f in result["exported_slides"]]
        assert "slide-000.png" in filenames
        assert "slide-002.png" in filenames

    def test_run_dir_structure(self, com_mocks):
        """US-10: run_dir path follows exports/run-<UTC-stamp>/png pattern."""
        result = export_slides_to_stamped_dir(
            com_mocks["pptx"],
            str(com_mocks["tmp"]),
            slide_indices=None,
            confirm=True,
        )
        assert "exports" in result["run_dir"]
        assert "png" in result["run_dir"]
        assert re.search(
            r"exports[/\\]run-\d{8}T\d{6}\d+Z[/\\]png",
            result["run_dir"],
        )

    def test_width_height_forwarded(self, com_mocks):
        """US-10: width_px=800, height_px=600 forwarded to every slide.Export call."""
        result = export_slides_to_stamped_dir(
            com_mocks["pptx"],
            str(com_mocks["tmp"]),
            slide_indices=None,
            width_px=800,
            height_px=600,
            confirm=True,
        )
        assert result["slide_count"] == 3
        for call in com_mocks["slide"].Export.call_args_list:
            assert call.args[2:] == (800, 600)

    def test_base_dir_outside_allowlist_raises(self, com_mocks):
        """US-10: base_dir not inside PPT_ALLOWLIST_ROOTS → ValidationError."""
        base_dir = str(com_mocks["tmp"].parent / "notallowed")
        with pytest.raises(ValidationError):
            export_slides_to_stamped_dir(
                com_mocks["pptx"],
                base_dir,
                confirm=True,
            )

    def test_slide_index_out_of_range_raises(self, com_mocks):
        """US-10: index 99 exceeds slide count (3) → ValidationError before any Export."""
        with pytest.raises(ValidationError):
            export_slides_to_stamped_dir(
                com_mocks["pptx"],
                str(com_mocks["tmp"]),
                slide_indices=[0, 99],
                confirm=True,
            )
        assert com_mocks["slide"].Export.call_count == 0

    def test_com_gate_raises(self, tmp_path, monkeypatch):
        """US-10: PPT_ENABLE_COM absent → NotAllowedError before any COM call."""
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        with pytest.raises(NotAllowedError):
            export_slides_to_stamped_dir(
                str(tmp_path / "x.pptx"),
                str(tmp_path),
                confirm=True,
            )

    def test_write_gate_raises(self, tmp_path, monkeypatch):
        """US-10: PPT_ENABLE_WRITE absent → NotAllowedError before any COM call."""
        monkeypatch.setenv("PPT_ENABLE_COM", "true")
        monkeypatch.delenv("PPT_ENABLE_WRITE", raising=False)
        with pytest.raises(NotAllowedError):
            export_slides_to_stamped_dir(
                str(tmp_path / "x.pptx"),
                str(tmp_path),
                confirm=True,
            )

    def test_confirm_gate_raises(self, com_mocks):
        """US-10: confirm=False → ValidationError before any COM call."""
        with pytest.raises(ValidationError):
            export_slides_to_stamped_dir(
                com_mocks["pptx"],
                str(com_mocks["tmp"]),
                confirm=False,
            )


# ─────────────────────────────────────────────────────────────────────────────
# TestInputPathSecurity
# ─────────────────────────────────────────────────────────────────────────────


