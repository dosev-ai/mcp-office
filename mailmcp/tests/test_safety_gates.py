from __future__ import annotations

from pathlib import Path

import pytest
from fastmcp.exceptions import ToolError

from mailmcp import _core, _folders
from mailmcp._mail_compose import compose_mail
from mailmcp._mail_ops import send_mail
from mailmcp._tools import _safe


@pytest.fixture(autouse=True)
def restore_config():
    previous = _core._config
    previous_startup = _core._startup_config
    previous_overrides = dict(_core._account_overrides)
    previous_startup_overrides = None if _core._startup_account_overrides is None else dict(_core._startup_account_overrides)
    try:
        yield
    finally:
        _core._config = previous
        _core._startup_config = previous_startup
        _core._account_overrides = previous_overrides
        _core._startup_account_overrides = previous_startup_overrides


def test_write_gate_is_disabled_by_default() -> None:
    _core.set_config(_core.OutlookConfig())
    with pytest.raises(PermissionError, match="write mutations are disabled"):
        _core._assert_write_enabled()


def test_compose_requires_confirm_before_any_com_call() -> None:
    _core.set_config(_core.OutlookConfig(enable_write=True))
    with pytest.raises(ValueError, match="confirm=True"):
        compose_mail(to=["person@example.com"], subject="Synthetic", body="Synthetic", confirm=False)


def test_send_requires_explicit_send_enable_before_com() -> None:
    _core.set_config(_core.OutlookConfig(enable_send=False))
    with pytest.raises(PermissionError, match="send_mail is disabled"):
        send_mail("synthetic-entry-id", confirm=True)


def test_domain_allowlist_fails_closed() -> None:
    _core.set_config(_core.OutlookConfig(allowlist_domains=["example.com"]))
    _core._assert_domains_allowed(["person@example.com"])
    with pytest.raises(PermissionError, match="not in OUTLOOK_ALLOWLIST_DOMAINS"):
        _core._assert_domains_allowed(["person@other.test"])


def test_attachment_guard_rejects_missing_path(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="does not exist"):
        _folders._validate_attachment_path(str(tmp_path / "missing.pdf"))


def test_safe_translates_validation_errors() -> None:
    def fail() -> None:
        raise ValueError("synthetic invalid input")

    with pytest.raises(ToolError, match="synthetic invalid input"):
        _safe(fail)
