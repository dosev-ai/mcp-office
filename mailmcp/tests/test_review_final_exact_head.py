from __future__ import annotations

from datetime import datetime
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from mailmcp import _calendar, _core, _folders, _mail_calendar


def test_account_override_rejects_malformed_identifier(monkeypatch):
    monkeypatch.setenv("OUTLOOK_ACCOUNT_1_EMAIL", "ownerexample.com")

    with pytest.raises(ValueError, match="OUTLOOK_ACCOUNT_1_EMAIL must be a valid email address"):
        _core._parse_account_overrides()


def test_freebusy_rejects_unknown_explicit_account_before_recipient_lookup(mailbox):
    _core._account_overrides["owner-a@example.com"] = _core.OutlookAccountOverride(
        email="owner-a@example.com",
        allowlist_domains=["example.com"],
    )

    with pytest.raises(ValueError, match="Account not found"):
        _mail_calendar.check_freebusy(
            ["person@example.com"],
            "2026-01-01T09:00:00",
            "2026-01-01T10:00:00",
            account_email="unknown@example.com",
        )

    mailbox.mapi.CreateRecipient.assert_not_called()


def test_freebusy_converts_requested_range_to_local_wall_time(mailbox):
    recipient = SimpleNamespace(
        Resolve=Mock(),
        Resolved=True,
        AddressEntry=SimpleNamespace(Address="person@example.com"),
        Name="Synthetic Person",
        FreeBusy=Mock(return_value="0" * 200),
    )
    mailbox.mapi.CreateRecipient.return_value = recipient

    start = "2026-01-01T09:00:00+02:00"
    end = "2026-01-01T10:00:00+02:00"
    expected_start = datetime.fromisoformat(start).astimezone().replace(tzinfo=None)
    expected_midnight = expected_start.replace(hour=0, minute=0, second=0, microsecond=0)

    result = _mail_calendar.check_freebusy(["person@example.com"], start, end)

    recipient.FreeBusy.assert_called_once_with(expected_midnight, 30, True)
    assert result["attendees"][0]["slots"][0]["start"] == expected_start.isoformat()


def test_calendar_listing_uses_interval_overlap_filter(monkeypatch):
    items = SimpleNamespace(Sort=Mock(), Restrict=Mock(return_value=[]))
    calendar_folder = SimpleNamespace(Items=items)
    monkeypatch.setattr(_folders, "_get_calendar_folder", lambda account_email=None: calendar_folder)

    _calendar.list_calendar_events(
        start="2026-01-01T09:00:00",
        end="2026-01-01T11:00:00",
    )

    filter_text = items.Restrict.call_args.args[0]
    assert "[Start] <=" in filter_text
    assert "[End] >=" in filter_text
    assert "[Start] >=" not in filter_text


def test_calendar_update_rechecks_resolved_attendee_domain_before_save(monkeypatch, mailbox):
    mailbox.item.Class = 26
    mailbox.item.Parent = mailbox.folders[("a", "Calendar")]
    monkeypatch.setattr(
        _core,
        "_resolve_smtp_from_entry",
        lambda _entry: "resolved@blocked.example",
    )

    with pytest.raises(PermissionError, match="blocked.example"):
        _calendar.update_calendar_event(
            "synthetic-item",
            required_attendees=["alias@example.com"],
            confirm=True,
        )

    mailbox.item.Save.assert_not_called()
