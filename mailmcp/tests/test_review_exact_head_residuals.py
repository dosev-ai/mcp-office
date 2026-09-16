from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from mailmcp import (
    _calendar,
    _calendar_ops,
    _core,
    _mail_calendar,
    _mail_compose,
    _mail_ops,
    _message_fetch,
    _message_recipients,
    _messages,
)


def _install_account_override(*, redact_mode: str | None = None, max_items: int | None = None, allowlist_domains: list[str] | None = None) -> str:
    email = "owner-a@example.com"
    _core._account_overrides[email] = _core.OutlookAccountOverride(
        email=email,
        allowlist_folders=["Inbox", "Contacts", "Calendar", "Drafts"],
        redact_mode=redact_mode,
        max_items=max_items,
        allowlist_domains=allowlist_domains,
    )
    return email


def test_recipient_search_requires_account_scope_before_gal_resolution():
    _install_account_override()

    with pytest.raises(PermissionError, match="account_email is required"):
        _message_recipients.search_recipients(name="Synthetic Person")


def test_default_mailbox_stats_require_account_scope():
    _install_account_override()

    with pytest.raises(PermissionError, match="account_email is required"):
        _calendar.get_mailbox_stats()


def test_freebusy_uses_effective_account_policy(monkeypatch):
    email = _install_account_override(
        redact_mode="emails",
        max_items=1,
        allowlist_domains=["allowed.example"],
    )

    with pytest.raises(PermissionError, match="account_email is required"):
        _mail_calendar.check_freebusy(
            ["person@allowed.example"],
            "2026-09-16T09:00:00+00:00",
            "2026-09-16T10:00:00+00:00",
        )

    with pytest.raises(ValueError, match="max_items"):
        _mail_calendar.check_freebusy(
            ["one@allowed.example", "two@allowed.example"],
            "2026-09-16T09:00:00+00:00",
            "2026-09-16T10:00:00+00:00",
            account_email=email,
        )

    with pytest.raises(PermissionError, match="Domain"):
        _mail_calendar.check_freebusy(
            ["person@blocked.example"],
            "2026-09-16T09:00:00+00:00",
            "2026-09-16T10:00:00+00:00",
            account_email=email,
        )

    captured = {}

    def _fake_freebusy(**kwargs):
        captured.update(kwargs)
        return {"attendees": []}

    monkeypatch.setattr(_calendar_ops.ol, "check_freebusy", _fake_freebusy)
    _calendar_ops.outlook_check_freebusy(
        ["person@allowed.example"],
        "2026-09-16T09:00:00+00:00",
        "2026-09-16T10:00:00+00:00",
        account_email=email,
    )
    assert captured["account_email"] == email


def test_account_listing_redacts_with_each_store_policy(mailbox):
    _core.set_config(_core.OutlookConfig(redact_mode="none"))
    _install_account_override(redact_mode="emails")

    accounts = _message_fetch.list_accounts()

    assert accounts[0]["display_name"] == "[email]"
    assert accounts[1]["display_name"] == "owner-b@example.com"


def test_conversation_rejects_negative_max_items():
    with pytest.raises(ValueError, match="non-negative integer"):
        _mail_ops.get_conversation_thread(conversation_id="synthetic-conversation", max_items=-1)


def test_conversation_expands_wildcard_before_folder_resolution(monkeypatch):
    _core.set_config(_core.OutlookConfig(allowlist_folders=["*"]))
    message = SimpleNamespace(
        ConversationID="synthetic-conversation",
        ConversationTopic="Synthetic topic",
    )

    class Items(list):
        def Restrict(self, _query):
            return []

    resolved_names: list[str] = []

    def _resolve(name):
        resolved_names.append(name)
        return SimpleNamespace(Items=Items([message]))

    monkeypatch.setattr(_messages, "_all_folder_search_names", lambda cfg, account_email: ["Inbox"])
    monkeypatch.setattr(_mail_ops._folders, "_folder_by_name", _resolve)

    result = _mail_ops.get_conversation_thread(conversation_id="synthetic-conversation")

    assert result["count"] == 0
    assert resolved_names
    assert set(resolved_names) == {"Inbox"}


def test_forward_draft_redacts_outlook_derived_subject(mailbox):
    _core.set_config(_core.OutlookConfig(
        allowlist_folders=["Inbox", "Drafts"],
        enable_write=True,
        redact_mode="emails",
    ))
    mailbox.item.Parent = mailbox.folders[("a", "Inbox")]

    class Recipients:
        def __init__(self):
            self.rows = []

        @property
        def Count(self):
            return len(self.rows)

        def Remove(self, index):
            del self.rows[index - 1]

        def Add(self, address):
            row = SimpleNamespace(Type=1, Address=address)
            self.rows.append(row)
            return row

        def ResolveAll(self):
            return True

    fwd = SimpleNamespace(
        Recipients=Recipients(),
        Attachments=SimpleNamespace(Add=Mock()),
        HTMLBody="<html><body>Original</body></html>",
        Subject="Fwd: person@example.com",
        EntryID="forward-draft",
        Save=Mock(),
    )
    mailbox.item.Forward = Mock(return_value=fwd)

    result = _mail_compose.forward_mail(
        "synthetic-item",
        to=["recipient@example.com"],
        body="Synthetic forward",
        confirm=True,
    )

    assert result["subject"] == "Fwd: [email]"
