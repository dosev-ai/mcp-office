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
from unittest.mock import MagicMock

import pytest
from pptx import Presentation as _Prs

from pptmcp.presentation_com import (
    _VALID_SLIDE_FORMATS,
    _check_output_path_deck,
    _check_output_path_slide,
    _validate_dpi,
    _validate_export_format,
)
from pptmcp.presentation_pptx import ValidationError


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
    monkeypatch.setattr(
        "pptmcp.presentation_com._check_com_enabled", MagicMock()
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


class TestValidateExportFormat:

    def test_valid_png(self):
        assert _validate_export_format("png", _VALID_SLIDE_FORMATS) == "png"

    def test_valid_pdf(self):
        assert _validate_export_format("pdf", _VALID_SLIDE_FORMATS) == "pdf"

    def test_valid_svg(self):
        assert _validate_export_format("svg", _VALID_SLIDE_FORMATS) == "svg"

    def test_invalid_format_raises(self):
        with pytest.raises(ValidationError):
            _validate_export_format("gif", _VALID_SLIDE_FORMATS)

    def test_format_none_raises(self):
        """M-1: isinstance guard prevents AttributeError on None input."""
        with pytest.raises(ValidationError):
            _validate_export_format(None, _VALID_SLIDE_FORMATS)

    def test_format_uppercase_normalised(self):
        """M-2: uppercase input is valid and normalised to lowercase."""
        assert _validate_export_format("PNG", _VALID_SLIDE_FORMATS) == "png"


# ─────────────────────────────────────────────────────────────────────────────
# TestValidateDpi
# ─────────────────────────────────────────────────────────────────────────────


class TestValidateDpi:

    def test_none_passthrough(self):
        assert _validate_dpi(None) is None

    def test_valid_int(self):
        assert _validate_dpi(150) == 150

    def test_too_high_raises(self):
        """M-3: dpi > 600 is rejected."""
        with pytest.raises(ValidationError):
            _validate_dpi(601)

    def test_zero_raises(self):
        """M-3: dpi < 72 (0) is rejected."""
        with pytest.raises(ValidationError):
            _validate_dpi(0)

    def test_negative_raises(self):
        """M-3: negative dpi is rejected."""
        with pytest.raises(ValidationError):
            _validate_dpi(-1)

    def test_float_raises(self):
        """M-3: float dpi rejected even if value is in range (COM VARIANT type issue)."""
        with pytest.raises(ValidationError):
            _validate_dpi(150.0)

    def test_dpi_71_raises(self):
        """dpi=71 is below the minimum of 72."""
        with pytest.raises(ValidationError):
            _validate_dpi(71)

    def test_dpi_72_valid(self):
        """dpi=72 is the minimum valid value (boundary)."""
        assert _validate_dpi(72) == 72

    def test_dpi_600_valid(self):
        """dpi=600 is the maximum valid value (boundary)."""
        assert _validate_dpi(600) == 600


# ─────────────────────────────────────────────────────────────────────────────
# TestCheckOutputPathSlide
# ─────────────────────────────────────────────────────────────────────────────


class TestCheckOutputPathSlide:

    def test_png_valid_ext(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        result = _check_output_path_slide(str(tmp_path / "out.png"), "png")
        assert result.suffix == ".png"

    def test_pdf_valid_ext(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        result = _check_output_path_slide(str(tmp_path / "out.pdf"), "pdf")
        assert result.suffix == ".pdf"

    def test_svg_valid_ext(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        result = _check_output_path_slide(str(tmp_path / "out.svg"), "svg")
        assert result.suffix == ".svg"

    def test_wrong_ext_raises(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        with pytest.raises(ValidationError):
            _check_output_path_slide(str(tmp_path / "out.txt"), "png")

    def test_null_byte_raises(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        with pytest.raises(ValidationError, match="null"):
            _check_output_path_slide(str(tmp_path) + "/out\x00.png", "png")


# ─────────────────────────────────────────────────────────────────────────────
# TestCheckOutputPathDeck
# ─────────────────────────────────────────────────────────────────────────────


class TestCheckOutputPathDeck:

    def test_pdf_valid(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        src = (tmp_path / "src.pptx").resolve()
        result = _check_output_path_deck(str(tmp_path / "out.pdf"), "pdf", src)
        assert result.suffix == ".pdf"

    def test_pptx_valid(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        src = (tmp_path / "src.pptx").resolve()
        result = _check_output_path_deck(str(tmp_path / "copy.pptx"), "pptx", src)
        assert result.suffix == ".pptx"

    def test_images_dir_ok(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        src = (tmp_path / "src.pptx").resolve()
        result = _check_output_path_deck(str(tmp_path / "images"), "images", src)
        assert result.name == "images"

    def test_pptx_same_path_raises(self, tmp_path, monkeypatch):
        """M-4: output path identical to source raises ValidationError."""
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        src = (tmp_path / "src.pptx").resolve()
        with pytest.raises(ValidationError, match="same"):
            _check_output_path_deck(str(src), "pptx", src)

    def test_wrong_ext_raises(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPT_ALLOWLIST_ROOTS", str(tmp_path))
        src = (tmp_path / "src.pptx").resolve()
        with pytest.raises(ValidationError):
            _check_output_path_deck(str(tmp_path / "out.png"), "pdf", src)


# ─────────────────────────────────────────────────────────────────────────────
# TestPptExportSlide
# ─────────────────────────────────────────────────────────────────────────────


