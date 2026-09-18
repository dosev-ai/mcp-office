from __future__ import annotations

import importlib.metadata
from types import SimpleNamespace

import pytest

from mailmcp import _context, _core, _mail_compose, _message_actions, _message_fetch


def test_mail_context_rejects_negative_caller_bounds():
    _core.set_config(_core.OutlookConfig(allowlist_folders=["Inbox"]))

    with pytest.raises(ValueError, match="max_items must be a non-negative integer"):
        _context.get_mail_context(folder="Inbox", max_items=-1)

    with pytest.raises(ValueError, match="body_preview_chars must be a non-negative integer"):
        _context.get_mail_context(folder="Inbox", body_preview_chars=-1)


def test_reply_recipient_resolution_uses_address_entry(monkeypatch):
    address_entry = object()
    recipient = SimpleNamespace(AddressEntry=address_entry)

    class Recipients:
        Count = 1

        @staticmethod
        def Item(index):
            assert index == 1
            return recipient

    seen: list[object] = []

    def fake_resolve(entry):
        seen.append(entry)
        return "person@example.com"

    monkeypatch.setattr(_mail_compose, "_resolve_smtp_from_entry", fake_resolve)

    assert _mail_compose._resolved_reply_addresses(
        SimpleNamespace(Recipients=Recipients())
    ) == ["person@example.com"]
    assert seen == [address_entry]


def test_bulk_message_action_normalizes_operation_before_dispatch(monkeypatch):
    dispatched: list[str] = []

    def fake_handle_message_action(**kwargs):
        dispatched.append(kwargs["operation"])
        return {"ok": True}

    monkeypatch.setattr(
        _message_actions,
        "handle_message_action",
        fake_handle_message_action,
    )

    result = _message_actions.outlook_bulk_message_action(
        entry_ids=["synthetic-message"],
        operation=" MOVE ",
        target_folder="Archive",
        confirm=True,
    )

    assert result["succeeded"] == 1
    assert dispatched == ["move"]


def test_health_does_not_fail_when_distribution_metadata_is_missing(monkeypatch):
    _core.set_config(_core.OutlookConfig())
    monkeypatch.setattr(
        _message_fetch._core,
        "_get_outlook",
        lambda: SimpleNamespace(Name="Microsoft Outlook", Version="synthetic"),
    )

    def missing_version(_name):
        raise importlib.metadata.PackageNotFoundError("mcp-office")

    monkeypatch.setattr(_message_fetch.importlib.metadata, "version", missing_version)

    result = _message_fetch.health()

    assert result["status"] == "ok"
    assert result["server_version"] == "unknown"
