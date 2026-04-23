"""Unit tests for excelmcp._da_com (COM gate guard on probe_excel_reopen).

All tests run without live Excel. COM interaction is mocked via sys.modules
patching when required.
win32com is NEVER imported at module level in this test file.

Test organisation:
    TestProbeReopenGate  — _da_com.probe_excel_reopen COM-gate guard (3 tests)

Exception type contracts:
    - gate tests call probe_excel_reopen directly → expect NotAllowedError
      (PermissionError subclass from excelmcp._core) when gate is blocked.
"""
from __future__ import annotations

import os
import sys
import tempfile
from unittest.mock import MagicMock

import pytest

from excelmcp._core import NotAllowedError

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="COM requires Windows")


# ---------------------------------------------------------------------------
# TestProbeReopenGate (3 tests)
# ---------------------------------------------------------------------------

class TestProbeReopenGate:
    """COM-gate guard tests for _da_com.probe_excel_reopen.

    Verifies that probe_excel_reopen raises NotAllowedError when
    EXCEL_ENABLE_COM is absent or not 'true', and that it proceeds to call
    win32com.client.Dispatch when the gate passes.
    """

    @pytest.mark.unit
    def test_probe_reopen_gate_blocked_env_absent(self, monkeypatch):
        """probe_excel_reopen raises NotAllowedError when EXCEL_ENABLE_COM is absent."""
        monkeypatch.delenv("EXCEL_ENABLE_COM", raising=False)

        from excelmcp._da_com import probe_excel_reopen

        with pytest.raises(NotAllowedError, match="EXCEL_ENABLE_COM"):
            probe_excel_reopen("/some/path.xlsx")

    @pytest.mark.unit
    def test_probe_reopen_gate_blocked_env_false(self, monkeypatch):
        """probe_excel_reopen raises NotAllowedError when EXCEL_ENABLE_COM=false."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "false")

        from excelmcp._da_com import probe_excel_reopen

        with pytest.raises(NotAllowedError, match="EXCEL_ENABLE_COM"):
            probe_excel_reopen("/some/path.xlsx")

    @pytest.mark.unit
    def test_probe_reopen_gate_passes_dispatches_excel(self, monkeypatch):
        """With EXCEL_ENABLE_COM=true and mocked COM, probe proceeds and calls Dispatch."""
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")

        mock_pythoncom = MagicMock()
        mock_win32client = MagicMock()
        mock_excel = MagicMock()
        mock_wb = MagicMock()
        mock_excel.Workbooks.Open.return_value = mock_wb
        mock_win32client.Dispatch.return_value = mock_excel
        mock_win32com_mod = MagicMock()
        mock_win32com_mod.client = mock_win32client

        monkeypatch.setitem(sys.modules, "pythoncom", mock_pythoncom)
        monkeypatch.setitem(sys.modules, "win32com", mock_win32com_mod)
        monkeypatch.setitem(sys.modules, "win32com.client", mock_win32client)

        # Use a file in the system tempdir — the allowlist guard auto-passes for temp paths.
        fd, tmp_file = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        try:
            from excelmcp._da_com import probe_excel_reopen
            result = probe_excel_reopen(tmp_file)
            assert result["ok"] is True, f"Expected ok=True, got: {result}"
        finally:
            if os.path.exists(tmp_file):
                os.unlink(tmp_file)

        mock_win32client.Dispatch.assert_called_with("Excel.Application")
