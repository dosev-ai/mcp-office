"""Unit tests for server_com_vba.py (US-E-029: list_macros + run_macro).

VBA / COM tests are Windows-only. COM objects are always mocked — no live Office.
Integration class requires @pytest.mark.integration and a live Excel install.
"""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock, PropertyMock, patch

import pytest

from excelmcp._core import NotAllowedError, OfficeCOMError, ValidationError
from excelmcp.server_com_vba import list_macros, run_macro

pytestmark = pytest.mark.skipif(
    sys.platform != "win32",
    reason="VBA COM tools are Windows-only",
)


# ---------------------------------------------------------------------------
# Shared fixture
# ---------------------------------------------------------------------------

@pytest.fixture()
def vba_env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Set environment gates and return a tmp_path registered as allowlist root."""
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_ENABLE_MACROS", "true")
    monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
    return tmp_path


def _make_xlsm(base: Path, name: str = "test_macro.xlsm") -> Path:
    """Create an empty file with .xlsm extension (COM never actually opens it in unit tests)."""
    p = base / name
    p.write_bytes(b"PK\x03\x04")  # minimal fake xlsx/xlsm header
    return p


# ---------------------------------------------------------------------------
# TestListMacros
# ---------------------------------------------------------------------------

class TestListMacros:

    def test_happy_path(self, vba_env: Path) -> None:
        """list_macros returns component names/types from VBProject."""
        wb_path = _make_xlsm(vba_env)

        mock_comp1 = MagicMock()
        mock_comp1.Name = "Module1"
        mock_comp1.Type = 1
        mock_comp2 = MagicMock()
        mock_comp2.Name = "Sheet1"
        mock_comp2.Type = 100

        mock_vbp = MagicMock()
        mock_vbp.VBComponents.Count = 2
        mock_vbp.VBComponents.Item.side_effect = [mock_comp1, mock_comp2]

        mock_wb = MagicMock()
        mock_wb.VBProject = mock_vbp

        mock_excel = MagicMock()

        with (
            patch("excelmcp.server_com_vba._com_excel_app") as mock_ctx,
            patch("excelmcp.server_com_vba._open_workbook", return_value=mock_wb),
            patch("excelmcp.server_com_vba._close_workbook"),
        ):
            mock_ctx.return_value.__enter__ = lambda s: mock_excel
            mock_ctx.return_value.__exit__ = lambda s, *a: None

            result = list_macros(str(wb_path))

        assert result["ok"] is True
        assert result["count"] == 2
        assert result["components"][0]["name"] == "Module1"
        assert result["components"][1]["name"] == "Sheet1"

    def test_com_gate_disabled(
        self, vba_env: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """EXCEL_ENABLE_COM=false raises NotAllowedError before COM is touched."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "false")
        wb_path = _make_xlsm(vba_env)

        with pytest.raises(NotAllowedError):
            list_macros(str(wb_path))

    def test_path_not_in_allowlist(
        self, vba_env: Path, tmp_path: Path
    ) -> None:
        """Path outside allowlist raises ValidationError."""
        outside = tmp_path.parent / "outside_vba"
        outside.mkdir(exist_ok=True)
        wb_path = outside / "macro.xlsm"
        wb_path.write_bytes(b"PK\x03\x04")

        with pytest.raises(ValidationError):
            list_macros(str(wb_path))

    def test_vbide_access_denied(self, vba_env: Path) -> None:
        """VBProject access raising an exception produces ValidationError with trust-center message."""
        wb_path = _make_xlsm(vba_env)

        mock_wb = MagicMock()
        type(mock_wb).VBProject = PropertyMock(side_effect=Exception("Access denied"))

        mock_excel = MagicMock()

        with (
            patch("excelmcp.server_com_vba._com_excel_app") as mock_ctx,
            patch("excelmcp.server_com_vba._open_workbook", return_value=mock_wb),
            patch("excelmcp.server_com_vba._close_workbook"),
        ):
            mock_ctx.return_value.__enter__ = lambda s: mock_excel
            mock_ctx.return_value.__exit__ = lambda s, *a: None

            with pytest.raises(ValidationError, match="Trust Center"):
                list_macros(str(wb_path))

    def test_empty_vb_project(self, vba_env: Path) -> None:
        """An empty VBProject (Count=0) returns an empty components list."""
        wb_path = _make_xlsm(vba_env)

        mock_vbp = MagicMock()
        mock_vbp.VBComponents.Count = 0
        mock_wb = MagicMock()
        mock_wb.VBProject = mock_vbp

        mock_excel = MagicMock()

        with (
            patch("excelmcp.server_com_vba._com_excel_app") as mock_ctx,
            patch("excelmcp.server_com_vba._open_workbook", return_value=mock_wb),
            patch("excelmcp.server_com_vba._close_workbook"),
        ):
            mock_ctx.return_value.__enter__ = lambda s: mock_excel
            mock_ctx.return_value.__exit__ = lambda s, *a: None

            result = list_macros(str(wb_path))

        assert result["count"] == 0
        assert result["components"] == []

    def test_requires_xlsm(self, vba_env: Path) -> None:
        """Passing an .xlsx file raises ValidationError immediately (no COM opened)."""
        xlsx_path = vba_env / "workbook.xlsx"
        xlsx_path.write_bytes(b"PK\x03\x04")

        with (
            patch("excelmcp.server_com_vba._com_excel_app") as mock_ctx,
        ):
            with pytest.raises(ValidationError, match=".xlsm"):
                list_macros(str(xlsx_path))
            mock_ctx.assert_not_called()


# ---------------------------------------------------------------------------
# TestRunMacro
# ---------------------------------------------------------------------------

class TestRunMacro:

    def test_happy_path(self, vba_env: Path) -> None:
        """run_macro calls excel.Run with correct module-qualified reference."""
        wb_path = _make_xlsm(vba_env)

        mock_wb = MagicMock()
        mock_excel = MagicMock()

        with (
            patch("excelmcp.server_com_vba._com_excel_app_macros") as mock_ctx,
            patch("excelmcp.server_com_vba._open_workbook", return_value=mock_wb),
            patch("excelmcp.server_com_vba._close_workbook"),
        ):
            mock_ctx.return_value.__enter__ = lambda s: mock_excel
            mock_ctx.return_value.__exit__ = lambda s, *a: None

            result = run_macro(
                path=str(wb_path),
                macro_name="Module1.RunReport",
                confirm=True,
            )

        assert result["ok"] is True
        assert result["macro_name"] == "Module1.RunReport"
        assert result["elapsed_ms"] >= 0
        expected_ref = f"'{wb_path.name}'!Module1.RunReport"
        mock_excel.Run.assert_called_once_with(expected_ref)

    def test_blocked_no_macros_env(
        self, vba_env: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """EXCEL_ENABLE_MACROS not set raises NotAllowedError."""
        monkeypatch.delenv("EXCEL_ENABLE_MACROS", raising=False)
        wb_path = _make_xlsm(vba_env)

        with pytest.raises(NotAllowedError):
            run_macro(path=str(wb_path), macro_name="MyMacro", confirm=True)

    def test_blocked_macros_env_false(
        self, vba_env: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """EXCEL_ENABLE_MACROS=false raises NotAllowedError."""
        monkeypatch.setenv("EXCEL_ENABLE_MACROS", "false")
        wb_path = _make_xlsm(vba_env)

        with pytest.raises(NotAllowedError):
            run_macro(path=str(wb_path), macro_name="MyMacro", confirm=True)

    def test_blocked_no_com_env(
        self, vba_env: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """EXCEL_ENABLE_COM=false raises NotAllowedError."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "false")
        wb_path = _make_xlsm(vba_env)

        with pytest.raises(NotAllowedError):
            run_macro(path=str(wb_path), macro_name="MyMacro", confirm=True)

    def test_requires_confirm_true(self, vba_env: Path) -> None:
        """confirm=False raises NotAllowedError before any COM usage."""
        wb_path = _make_xlsm(vba_env)

        with pytest.raises((NotAllowedError, ValidationError)):
            run_macro(path=str(wb_path), macro_name="MyMacro", confirm=False)

    def test_rejects_xlsx_extension(self, vba_env: Path) -> None:
        """BLOCKER-4: .xlsx file raises ValidationError before COM is opened."""
        xlsx_path = vba_env / "not_macro.xlsx"
        xlsx_path.write_bytes(b"PK\x03\x04")

        with (
            patch("excelmcp.server_com_vba._com_excel_app_macros") as mock_ctx,
        ):
            with pytest.raises(ValidationError, match=".xlsm"):
                run_macro(path=str(xlsx_path), macro_name="MyMacro", confirm=True)
            mock_ctx.assert_not_called()

    def test_rejects_single_quote_in_path(self, vba_env: Path) -> None:
        """BLOCKER-3: single quote in filename raises ValidationError before COM."""
        bad_path = vba_env / "O'Brian.xlsm"
        bad_path.write_bytes(b"PK\x03\x04")

        with pytest.raises(ValidationError, match="single quote"):
            run_macro(path=str(bad_path), macro_name="MyMacro", confirm=True)

    def test_path_not_in_allowlist(self, vba_env: Path, tmp_path: Path) -> None:
        """Path outside allowlist raises ValidationError."""
        outside = tmp_path.parent / "outside_run"
        outside.mkdir(exist_ok=True)
        bad = outside / "macro.xlsm"
        bad.write_bytes(b"PK\x03\x04")

        with pytest.raises(ValidationError):
            run_macro(path=str(bad), macro_name="MyMacro", confirm=True)

    def test_rejects_invalid_macro_name_double_dot(self, vba_env: Path) -> None:
        """DEFECT-2: double-dot macro name is rejected by _MACRO_NAME_RE."""
        wb_path = _make_xlsm(vba_env)

        with pytest.raises(ValidationError, match="valid VBA identifier"):
            run_macro(path=str(wb_path), macro_name="Mod1..Sub1", confirm=True)

    def test_rejects_macro_name_with_spaces(self, vba_env: Path) -> None:
        """Macro name with spaces is rejected by _MACRO_NAME_RE."""
        wb_path = _make_xlsm(vba_env)

        with pytest.raises(ValidationError, match="valid VBA identifier"):
            run_macro(path=str(wb_path), macro_name="My Macro", confirm=True)

    @pytest.mark.parametrize("macro_name", ["RunSomething", "Module1.MyMacro"])
    def test_accepts_valid_macro_name_formats(
        self, vba_env: Path, macro_name: str
    ) -> None:
        """Both simple and module-qualified identifiers are accepted."""
        wb_path = _make_xlsm(vba_env)

        mock_wb = MagicMock()
        mock_excel = MagicMock()

        with (
            patch("excelmcp.server_com_vba._com_excel_app_macros") as mock_ctx,
            patch("excelmcp.server_com_vba._open_workbook", return_value=mock_wb),
            patch("excelmcp.server_com_vba._close_workbook"),
        ):
            mock_ctx.return_value.__enter__ = lambda s: mock_excel
            mock_ctx.return_value.__exit__ = lambda s, *a: None

            result = run_macro(path=str(wb_path), macro_name=macro_name, confirm=True)

        assert result["ok"] is True
        assert result["macro_name"] == macro_name

    def test_com_error_wrapped_as_office_com_error(self, vba_env: Path) -> None:
        """COM failures in excel.Run are re-raised as OfficeCOMError."""
        wb_path = _make_xlsm(vba_env)

        mock_wb = MagicMock()
        mock_excel = MagicMock()
        mock_excel.Run.side_effect = Exception("COM dispatch error")

        with (
            patch("excelmcp.server_com_vba._com_excel_app_macros") as mock_ctx,
            patch("excelmcp.server_com_vba._open_workbook", return_value=mock_wb),
            patch("excelmcp.server_com_vba._close_workbook"),
        ):
            mock_ctx.return_value.__enter__ = lambda s: mock_excel
            mock_ctx.return_value.__exit__ = lambda s, *a: None

            with pytest.raises(OfficeCOMError):
                run_macro(path=str(wb_path), macro_name="MyMacro", confirm=True)


# ---------------------------------------------------------------------------
# TestVbaIntegration  (requires live Excel — skipped in CI)
# ---------------------------------------------------------------------------

@pytest.mark.integration
class TestVbaIntegration:

    def test_list_macros_live_excel(self, vba_env: Path) -> None:
        """list_macros succeeds against a real .xlsm file with a known macro."""
        pytest.skip("requires live Excel with macro-enabled workbook")

    def test_run_macro_live_excel(self, vba_env: Path) -> None:
        """run_macro executes a real macro and returns ok=True."""
        pytest.skip("requires live Excel with a test macro workbook")

    def test_list_then_run_macro_roundtrip(self, vba_env: Path) -> None:
        """list_macros discovers a macro; run_macro executes it successfully."""
        pytest.skip("requires live Excel with macro-enabled workbook")
