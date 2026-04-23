"""Tests for excelmcp.runtime_config — 9 tests."""
from __future__ import annotations

import ast
import inspect

import pytest

from excelmcp.runtime_config import (
    REQUIRED_FIELDS,
    get_effective_policy,
    get_contract_version,
    is_strict,
    load_runtime_config,
)


def test_default_policy_is_strict(monkeypatch):
    """When EXCEL_METADATA_POLICY is unset, get_effective_policy() returns 'strict'."""
    monkeypatch.delenv("EXCEL_METADATA_POLICY", raising=False)
    assert get_effective_policy() == "strict"


def test_lenient_policy_via_env(monkeypatch):
    """EXCEL_METADATA_POLICY=lenient returns 'lenient' and is_strict() is False."""
    monkeypatch.setenv("EXCEL_METADATA_POLICY", "lenient")
    assert get_effective_policy() == "lenient"
    assert is_strict() is False


def test_invalid_policy_raises(monkeypatch):
    """EXCEL_METADATA_POLICY=garbage raises ValueError."""
    monkeypatch.setenv("EXCEL_METADATA_POLICY", "garbage")
    with pytest.raises(ValueError, match="EXCEL_METADATA_POLICY"):
        load_runtime_config()


def test_is_strict_returns_bool(monkeypatch):
    """is_strict() always returns a bool."""
    monkeypatch.delenv("EXCEL_METADATA_POLICY", raising=False)
    result = is_strict()
    assert isinstance(result, bool)
    assert result is True


def test_load_runtime_config_returns_dict(monkeypatch):
    """load_runtime_config() returns a dict with required keys."""
    monkeypatch.delenv("EXCEL_METADATA_POLICY", raising=False)
    monkeypatch.delenv("EXCEL_METADATA_CONTRACT_VERSION", raising=False)
    cfg = load_runtime_config()
    assert isinstance(cfg, dict)
    assert "metadata_policy" in cfg
    assert "metadata_contract_version" in cfg
    assert "required_fields" in cfg


def test_required_fields_constant():
    """REQUIRED_FIELDS is a tuple containing 'path' and 'workbook_name'."""
    assert isinstance(REQUIRED_FIELDS, tuple)
    assert "path" in REQUIRED_FIELDS
    assert "workbook_name" in REQUIRED_FIELDS


def test_contract_version_override_via_env(monkeypatch):
    """EXCEL_METADATA_CONTRACT_VERSION env overrides the compiled-in default."""
    monkeypatch.setenv("EXCEL_METADATA_CONTRACT_VERSION", "2.0")
    v = get_contract_version()
    assert v == "2.0"
    cfg = load_runtime_config()
    assert cfg["metadata_contract_version"] == "2.0"


def test_no_fastmcp_import_in_runtime_module():
    """runtime_config must import neither fastmcp.server nor any _com module."""
    import excelmcp.runtime_config as mod
    source = inspect.getsource(mod)
    tree = ast.parse(source)
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module:
            assert not node.module.startswith("fastmcp"), (
                f"runtime_config must not import FastMCP modules; found: {node.module}"
            )
            assert "_com" not in node.module, (
                f"runtime_config must not import COM modules; found: {node.module}"
            )


def test_invalid_policy_warns_and_falls_back(monkeypatch):
    """get_effective_policy() with invalid env value emits UserWarning and returns 'strict'."""
    import warnings
    monkeypatch.setenv("EXCEL_METADATA_POLICY", "notathing")
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        result = get_effective_policy()
    assert result == "strict", f"Expected 'strict' fallback, got: {result!r}"
    assert len(w) == 1, f"Expected exactly 1 warning, got {len(w)}: {[str(x.message) for x in w]}"
    assert issubclass(w[0].category, UserWarning)
    assert "EXCEL_METADATA_POLICY" in str(w[0].message)


# ---------------------------------------------------------------------------
# D-01 through D-08: get_session_mode() — Sprint H active-workbook bridge
# ---------------------------------------------------------------------------

from excelmcp.runtime_config import get_session_mode  # noqa: E402


def test_get_session_mode_false_by_default(monkeypatch):
    """D-01: get_session_mode() returns False when EXCEL_SESSION_MODE is not set."""
    monkeypatch.delenv("EXCEL_SESSION_MODE", raising=False)
    assert get_session_mode() is False


def test_get_session_mode_true_when_both_enabled(monkeypatch):
    """D-02: get_session_mode() returns True when SESSION=true AND COM=true."""
    monkeypatch.setenv("EXCEL_SESSION_MODE", "true")
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    assert get_session_mode() is True


def test_get_session_mode_false_when_explicitly_false(monkeypatch):
    """D-03: get_session_mode() returns False when SESSION=false."""
    monkeypatch.setenv("EXCEL_SESSION_MODE", "false")
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    assert get_session_mode() is False


def test_get_session_mode_emits_user_warning_when_com_disabled(monkeypatch):
    """D-04: UserWarning emitted when SESSION=true but COM is not enabled."""
    monkeypatch.setenv("EXCEL_SESSION_MODE", "true")
    monkeypatch.setenv("EXCEL_ENABLE_COM", "false")
    with pytest.warns(UserWarning, match="SESSION_MODE"):
        result = get_session_mode()
    assert result is True  # still returns True — enabling the env is user's choice


def test_get_session_mode_no_warning_when_com_enabled(monkeypatch):
    """D-05: No UserWarning when SESSION=true and COM=true."""
    import warnings
    monkeypatch.setenv("EXCEL_SESSION_MODE", "true")
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        get_session_mode()
    user_warnings = [x for x in w if issubclass(x.category, UserWarning)]
    assert user_warnings == [], f"Unexpected UserWarning(s): {[str(x.message) for x in user_warnings]}"


def test_get_session_mode_no_warning_when_session_disabled(monkeypatch):
    """D-06: No UserWarning when SESSION=false (even if COM is also disabled)."""
    import warnings
    monkeypatch.setenv("EXCEL_SESSION_MODE", "false")
    monkeypatch.delenv("EXCEL_ENABLE_COM", raising=False)
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        get_session_mode()
    user_warnings = [x for x in w if issubclass(x.category, UserWarning)]
    assert user_warnings == [], f"Unexpected UserWarning(s): {[str(x.message) for x in user_warnings]}"


def test_get_session_mode_case_insensitive(monkeypatch):
    """D-07: EXCEL_SESSION_MODE=TRUE (uppercase) is treated as True."""
    monkeypatch.setenv("EXCEL_SESSION_MODE", "TRUE")
    monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
    assert get_session_mode() is True


def test_no_com_import_in_runtime_config():
    """D-08: runtime_config.py must not import any _com module or win32com at module level."""
    import excelmcp.runtime_config as mod
    source = inspect.getsource(mod)
    tree = ast.parse(source)
    for node in ast.walk(tree):
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            if isinstance(node, ast.ImportFrom) and node.module:
                assert "_com" not in node.module, (
                    f"runtime_config must not import COM modules; found: {node.module}"
                )
                assert "win32com" not in node.module, (
                    f"runtime_config must not import win32com; found: {node.module}"
                )
            if isinstance(node, ast.Import):
                for alias in node.names:
                    assert "_com" not in alias.name, (
                        f"runtime_config must not import COM modules; found: {alias.name}"
                    )
                    assert "win32com" not in alias.name, (
                        f"runtime_config must not import win32com; found: {alias.name}"
                    )
