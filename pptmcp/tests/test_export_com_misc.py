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
    _check_com_enabled,
    ppt_export_deck,
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


class TestInputPathSecurity:
    """Verify input file_path traversal attacks are blocked by _check_path."""

    def test_input_traversal_slide(self, com_env):
        """'../evil.pptx' as file_path is rejected by _check_path before any COM call."""
        out = str(Path(com_env["pptx"]).parent / "out_trav.png")
        with pytest.raises(ValidationError, match="allowlist"):
            ppt_export_slide("../evil.pptx", 0, "png", out, confirm=True)

    def test_input_traversal_deck(self, com_env):
        """'../evil.pptx' as file_path for deck export is rejected by _check_path."""
        out = str(Path(com_env["pptx"]).parent / "out_trav.pdf")
        with pytest.raises(ValidationError, match="allowlist"):
            ppt_export_deck("../evil.pptx", "pdf", out, confirm=True)

    def test_path_like_format_raises(self, com_env):
        """B1: format='../evil' is path-like and rejected by _validate_export_format."""
        out = str(Path(com_env["pptx"]).parent / "out_fmt.png")
        with pytest.raises(ValidationError):
            ppt_export_slide(com_env["pptx"], 0, "../evil", out, confirm=True)


# ─────────────────────────────────────────────────────────────────────────────
# TestCheckComEnabled
# ─────────────────────────────────────────────────────────────────────────────


class TestCheckComEnabled:
    """US-12: _check_com_enabled enforces PPT_ENABLE_COM env var."""

    def test_raises_when_env_var_absent(self, monkeypatch):
        """US-12: PPT_ENABLE_COM absent → NotAllowedError matching 'PPT_ENABLE_COM=true'."""
        monkeypatch.delenv("PPT_ENABLE_COM", raising=False)
        with pytest.raises(NotAllowedError, match=r"PPT_ENABLE_COM=true"):
            _check_com_enabled()

    def test_raises_when_env_var_wrong_value(self, monkeypatch):
        """US-12: PPT_ENABLE_COM='yes' is not 'true' → NotAllowedError."""
        monkeypatch.setenv("PPT_ENABLE_COM", "yes")
        with pytest.raises(NotAllowedError, match=r"PPT_ENABLE_COM=true"):
            _check_com_enabled()


# ─────────────────────────────────────────────────────────────────────────────
# TestDeprecatedAliases  (imports pptmcp.server — no live COM required)
# ─────────────────────────────────────────────────────────────────────────────


class TestDeprecatedAliases:
    """Verify deprecated server.py wrappers delegate correctly to the new API.

    These tests mock _pcom — no live COM required.
    They are skipped automatically when export_slide_as_png / export_deck_as_pdf
    are not registered (i.e. _COM_AVAILABLE=False at import time).
    """

    @pytest.fixture(autouse=True)
    def _skip_if_no_server(self):
        pytest.importorskip("pptmcp.server",
                            reason="pptmcp.server not importable — skipping alias tests")

    def test_slide_as_png_calls_export_slide(self, monkeypatch, tmp_path):
        """export_slide_as_png.fn(...) delegates to _pcom.ppt_export_slide with fmt='png'."""
        import pptmcp.server as _srv

        if not hasattr(_srv, "export_slide_as_png"):
            pytest.skip("export_slide_as_png not registered (COM unavailable)")

        mock_fn = MagicMock(return_value={
            "exported": True, "slide_number": 1, "format": "png",
            "output_path": str(tmp_path / "out.png"),
        })
        monkeypatch.setattr(_srv._pcom, "ppt_export_slide", mock_fn)

        _srv.export_slide_as_png(
            file_path="x.pptx", slide_index=0,
            output_path=str(tmp_path / "out.png"), confirm=True,
        )
        mock_fn.assert_called_once()
        _, kwargs = mock_fn.call_args
        assert kwargs["fmt"] == "png"

    def test_slide_index_passthrough(self, monkeypatch, tmp_path):
        """export_slide_as_png(slide_index=0) → ppt_export_slide(slide_index=0)."""
        import pptmcp.server as _srv

        if not hasattr(_srv, "export_slide_as_png"):
            pytest.skip("export_slide_as_png not registered (COM unavailable)")

        mock_fn = MagicMock(return_value={
            "exported": True, "slide_number": 1, "format": "png",
            "output_path": str(tmp_path / "out.png"),
        })
        monkeypatch.setattr(_srv._pcom, "ppt_export_slide", mock_fn)

        _srv.export_slide_as_png(
            file_path="x.pptx", slide_index=0,
            output_path=str(tmp_path / "out.png"), confirm=True,
        )
        _, kwargs = mock_fn.call_args
        assert kwargs["slide_index"] == 0, (
            "0-based slide_index=0 must be passed through unchanged to ppt_export_slide"
        )

    def test_deck_as_pdf_calls_export_deck(self, monkeypatch, tmp_path):
        """export_deck_as_pdf delegates to _pcom.ppt_export_deck with fmt='pdf'."""
        import pptmcp.server as _srv

        if not hasattr(_srv, "export_deck_as_pdf"):
            pytest.skip("export_deck_as_pdf not registered (COM unavailable)")

        mock_fn = MagicMock(return_value={
            "exported": True, "format": "pdf", "slide_count": 3,
            "output_path": str(tmp_path / "out.pdf"),
        })
        monkeypatch.setattr(_srv._pcom, "ppt_export_deck", mock_fn)

        _srv.export_deck_as_pdf(
            file_path="x.pptx",
            output_path=str(tmp_path / "out.pdf"), confirm=True,
        )
        mock_fn.assert_called_once()
        _, kwargs = mock_fn.call_args
        assert kwargs["fmt"] == "pdf"


# ─────────────────────────────────────────────────────────────────────────────
# TestPptAppContextDispatchError  (BUG-2)
# ─────────────────────────────────────────────────────────────────────────────


@pytest.mark.skipif(
    sys.platform != "win32",
    reason="_ppt_app_context import path requires win32com on Windows",
)
class TestPptAppContextDispatchError:
    """BUG-2: com_error thrown by Dispatch() during __enter__ must be wrapped
    into OfficeCOMError — never allowed to reach FastMCP as a raw pywintypes.com_error.
    """

    def test_dispatch_com_error_wrapped_for_slide(self, com_env, monkeypatch):
        """ppt_export_slide: Dispatch() com_error → OfficeCOMError, not raw com_error."""
        # Make win32com.client.Dispatch raise com_error (mocked as plain Exception
        # via com_env fixture which sets pythoncom.com_error = Exception)
        import sys
        _w32 = sys.modules["win32com.client"]  # already mocked by com_env
        monkeypatch.setattr(_w32, "Dispatch", MagicMock(side_effect=Exception("launch fail")))

        out = str(Path(com_env["pptx"]).parent / "bug2_slide.png")
        with pytest.raises(OfficeCOMError, match="PowerPoint.Application"):
            ppt_export_slide(com_env["pptx"], 0, "png", out, confirm=True)

    def test_dispatch_com_error_wrapped_for_deck(self, com_env, monkeypatch):
        """ppt_export_deck: Dispatch() com_error → OfficeCOMError, not raw com_error."""
        import sys
        _w32 = sys.modules["win32com.client"]  # already mocked by com_env
        monkeypatch.setattr(_w32, "Dispatch", MagicMock(side_effect=Exception("launch fail")))

        out = str(Path(com_env["pptx"]).parent / "bug2_deck.pdf")
        with pytest.raises(OfficeCOMError, match="PowerPoint.Application"):
            ppt_export_deck(com_env["pptx"], "pdf", out, confirm=True)

    def test_dispatch_error_is_not_raw_com_error(self, com_env, monkeypatch):
        """The raised exception must NOT be the raw com_error itself."""
        import sys
        _w32 = sys.modules["win32com.client"]  # already mocked by com_env
        raw_exc = Exception("raw com failure")
        monkeypatch.setattr(_w32, "Dispatch", MagicMock(side_effect=raw_exc))

        out = str(Path(com_env["pptx"]).parent / "bug2_notraw.png")
        try:
            ppt_export_slide(com_env["pptx"], 0, "png", out, confirm=True)
        except OfficeCOMError:
            pass  # correct — wrapped
        except Exception as caught:
            pytest.fail(
                f"Expected OfficeCOMError but got {type(caught).__name__}: {caught}"
            )
