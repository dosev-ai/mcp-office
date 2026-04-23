"""Focused unit tests for Excel COM startup-error normalization."""
from __future__ import annotations

import sys
from unittest.mock import MagicMock

import pytest

from excelmcp._core import OfficeCOMError

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="COM requires Windows")


class _BootstrapExcelStub:
    """Minimal Excel.Application stub for startup-stage failure tests."""

    # xl_app.Ready polling loop breaks immediately when this is True.
    Ready = True

    def __init__(self, fail_stage: str | None = None):
        self.fail_stage = fail_stage
        self.quit_calls = 0
        self.fail_quit_calls: set[int] = set()
        self._visible = None
        self._display_alerts = None
        self._automation_security = None

    @property
    def Visible(self):
        return self._visible

    @Visible.setter
    def Visible(self, value):
        if self.fail_stage == "Visible":
            raise RuntimeError("visible boom\nunsafe detail")
        self._visible = value

    @property
    def DisplayAlerts(self):
        return self._display_alerts

    @DisplayAlerts.setter
    def DisplayAlerts(self, value):
        if self.fail_stage == "DisplayAlerts":
            raise RuntimeError("alerts boom\nunsafe detail")
        self._display_alerts = value

    @property
    def AutomationSecurity(self):
        return self._automation_security

    @AutomationSecurity.setter
    def AutomationSecurity(self, value):
        if self.fail_stage == "AutomationSecurity":
            raise RuntimeError("automation boom\nunsafe detail")
        self._automation_security = value

    def Quit(self):
        self.quit_calls += 1
        if self.quit_calls in self.fail_quit_calls:
            raise RuntimeError("quit boom\nunsafe detail")


def _patch_startup(monkeypatch, *, pythoncom: MagicMock, win32c: MagicMock) -> None:
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setattr(
        "excelmcp._com._ensure_com_available",
        lambda: (pythoncom, win32c),
    )


def _assert_bootstrap_property_failure(
    monkeypatch,
    *,
    stage_name: str,
    expected_message: str,
) -> None:
    mock_pythoncom = MagicMock()
    mock_win32c = MagicMock()
    excel = _BootstrapExcelStub(fail_stage=stage_name)
    mock_win32c.Dispatch.return_value = excel
    _patch_startup(monkeypatch, pythoncom=mock_pythoncom, win32c=mock_win32c)

    from excelmcp._com import _com_excel_app

    with pytest.raises(OfficeCOMError) as exc_info:
        with _com_excel_app():
            pass

    assert str(exc_info.value) == expected_message
    assert "\n" not in str(exc_info.value)
    assert excel.quit_calls == 1
    mock_win32c.Dispatch.assert_called_once_with("Excel.Application")
    mock_pythoncom.CoInitialize.assert_called_once()
    mock_pythoncom.CoUninitialize.assert_called_once()


@pytest.mark.unit
def test_com_excel_app_coinitialize_failure_normalized(monkeypatch):
    """_com_excel_app() surfaces a startup-specific CoInitialize failure."""
    mock_pythoncom = MagicMock()
    mock_pythoncom.CoInitialize.side_effect = RuntimeError("coinit boom\nunsafe detail")
    mock_win32c = MagicMock()
    _patch_startup(monkeypatch, pythoncom=mock_pythoncom, win32c=mock_win32c)

    from excelmcp._com import _com_excel_app

    with pytest.raises(OfficeCOMError) as exc_info:
        with _com_excel_app():
            pass

    assert str(exc_info.value) == "Excel COM startup failed during CoInitialize()"
    assert "\n" not in str(exc_info.value)
    mock_win32c.Dispatch.assert_not_called()
    mock_pythoncom.CoUninitialize.assert_not_called()


@pytest.mark.unit
def test_com_excel_app_dispatch_failure_normalized(monkeypatch):
    """_com_excel_app() surfaces a startup-specific dispatch failure."""
    mock_pythoncom = MagicMock()
    mock_win32c = MagicMock()
    mock_win32c.Dispatch.side_effect = RuntimeError("dispatch boom\nunsafe detail")
    _patch_startup(monkeypatch, pythoncom=mock_pythoncom, win32c=mock_win32c)

    from excelmcp._com import _com_excel_app

    with pytest.raises(OfficeCOMError) as exc_info:
        with _com_excel_app():
            pass

    assert str(exc_info.value) == "Excel COM startup failed creating Excel.Application"
    assert "\n" not in str(exc_info.value)
    mock_win32c.Dispatch.assert_called_once_with("Excel.Application")
    mock_pythoncom.CoInitialize.assert_called_once()
    mock_pythoncom.CoUninitialize.assert_called_once()


@pytest.mark.unit
def test_com_excel_app_visible_bootstrap_failure_normalized(monkeypatch):
    """_com_excel_app() surfaces a startup-specific `Visible` failure and cleans up."""
    _assert_bootstrap_property_failure(
        monkeypatch,
        stage_name="Visible",
        expected_message="Excel COM startup failed setting Excel.Visible",
    )


@pytest.mark.unit
def test_com_excel_app_displayalerts_bootstrap_failure_normalized(monkeypatch):
    """_com_excel_app() surfaces a startup-specific `DisplayAlerts` failure and cleans up."""
    _assert_bootstrap_property_failure(
        monkeypatch,
        stage_name="DisplayAlerts",
        expected_message="Excel COM startup failed setting Excel.DisplayAlerts",
    )


@pytest.mark.unit
def test_com_excel_app_automationsecurity_bootstrap_failure_normalized(monkeypatch):
    """_com_excel_app() surfaces a startup-specific `AutomationSecurity` failure and cleans up."""
    _assert_bootstrap_property_failure(
        monkeypatch,
        stage_name="AutomationSecurity",
        expected_message="Excel COM startup failed setting Excel.AutomationSecurity",
    )


@pytest.mark.unit
def test_com_excel_app_startup_error_survives_couninitialize_failure(monkeypatch):
    """A startup-specific OfficeCOMError is not overridden by CoUninitialize cleanup failure."""
    mock_pythoncom = MagicMock()
    mock_pythoncom.CoUninitialize.side_effect = RuntimeError("uninit boom\nunsafe detail")
    mock_win32c = MagicMock()
    excel = _BootstrapExcelStub(fail_stage="Visible")
    mock_win32c.Dispatch.return_value = excel
    _patch_startup(monkeypatch, pythoncom=mock_pythoncom, win32c=mock_win32c)

    from excelmcp._com import _com_excel_app

    with pytest.raises(OfficeCOMError) as exc_info:
        with _com_excel_app():
            pass

    assert str(exc_info.value) == "Excel COM startup failed setting Excel.Visible"
    assert "\n" not in str(exc_info.value)
    assert excel.quit_calls == 1
    mock_pythoncom.CoInitialize.assert_called_once()
    mock_pythoncom.CoUninitialize.assert_called_once()


@pytest.mark.unit
def test_com_excel_app_property_failure_retries_cleanup_after_quit_failure(monkeypatch):
    """A property-stage startup OfficeCOMError is propagated even when Quit() raises.

    _excel_configure() calls Quit() when a property setter fails.  When that Quit()
    itself raises, _excel_configure swallows the error and then raises OfficeCOMError.
    The outer _com_excel_app finally-block sets excel=None after the configure step
    fails, so exactly 1 Quit() attempt is made (in _excel_configure).
    """
    mock_pythoncom = MagicMock()
    mock_win32c = MagicMock()
    excel = _BootstrapExcelStub(fail_stage="DisplayAlerts")
    excel.fail_quit_calls = {1}
    mock_win32c.Dispatch.return_value = excel
    _patch_startup(monkeypatch, pythoncom=mock_pythoncom, win32c=mock_win32c)

    from excelmcp._com import _com_excel_app

    with pytest.raises(OfficeCOMError) as exc_info:
        with _com_excel_app():
            pass

    assert str(exc_info.value) == "Excel COM startup failed setting Excel.DisplayAlerts"
    assert "\n" not in str(exc_info.value)
    # _excel_configure makes one Quit() attempt; it fails but error is swallowed.
    assert excel.quit_calls == 1
    mock_pythoncom.CoInitialize.assert_called_once()
    mock_pythoncom.CoUninitialize.assert_called_once()


@pytest.mark.unit
def test_recalculate_workbook_startup_specific_message_survives_as_tool_error(
    monkeypatch,
    tmp_path,
):
    """`recalculate_workbook()` preserves startup-specific COM text at the server boundary."""
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    wb_path = tmp_path / "startup-fail.xlsx"
    wb_path.touch()

    mock_pythoncom = MagicMock()
    mock_pythoncom.CoInitialize.side_effect = RuntimeError("coinit boom\nunsafe detail")
    mock_win32c = MagicMock()
    monkeypatch.setattr(
        "excelmcp._com._ensure_com_available",
        lambda: (mock_pythoncom, mock_win32c),
    )

    from fastmcp.exceptions import ToolError
    from excelmcp.server_com import recalculate_workbook

    with pytest.raises(
        ToolError,
        match=r"Recalculation failed — ",
    ):
        recalculate_workbook(path=str(wb_path), confirm=True)

