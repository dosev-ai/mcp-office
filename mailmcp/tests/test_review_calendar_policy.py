from __future__ import annotations

from dataclasses import replace
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from mailmcp import _calendar, _core, _mail_calendar


@pytest.mark.parametrize("field", ["required_attendees", "optional_attendees"])
def test_calendar_domain_rejection_precedes_any_mutation(mailbox, field):
    mailbox.item.Class = 26
    mailbox.item.Parent = mailbox.folders[("a", "Calendar")]
    with pytest.raises(PermissionError):
        _calendar.update_calendar_event(
            "synthetic-item", subject="Changed", confirm=True,
            **{field: ["outside@other.test"]},
        )
    assert mailbox.item.Subject == "Synthetic original"
    assert mailbox.item.Recipients.mutations == []
    mailbox.item.Save.assert_not_called()


@pytest.mark.parametrize("reader", ["list", "get"])
def test_calendar_is_not_an_implicit_allowlist_exception(mailbox, reader):
    mailbox.item.Class = 26
    mailbox.item.Parent = mailbox.folders[("a", "Calendar")]
    _core.set_config(replace(_core.get_config(), allowlist_folders=["Inbox", "Contacts"]))
    with pytest.raises(PermissionError):
        if reader == "list":
            _calendar.list_calendar_events()
        else:
            _calendar.get_calendar_event("synthetic-item")


@pytest.mark.parametrize("end", ["2026-01-01T09:00:00", "2026-01-01T08:00:00"])
def test_freebusy_invalid_range_rejected_before_com(mailbox, end):
    with pytest.raises(ValueError):
        _mail_calendar.check_freebusy(["person@example.com"], "2026-01-01T09:00:00", end)
    mailbox.get_mapi.assert_not_called()


def test_freebusy_domain_guard_checks_whole_batch_before_com(mailbox):
    with pytest.raises(PermissionError):
        _mail_calendar.check_freebusy(
            ["person@example.com", "outside@other.test"],
            "2026-01-01T09:00:00", "2026-01-01T10:00:00",
        )
    mailbox.get_mapi.assert_not_called()


def test_freebusy_positive_preserves_response_shape(mailbox):
    recipient = SimpleNamespace(Resolve=Mock(), Resolved=True, Name="Synthetic", FreeBusy=Mock(return_value="0" * 100))
    mailbox.mapi.CreateRecipient.return_value = recipient
    result = _mail_calendar.check_freebusy(["person@example.com"], "2026-01-01T09:00:00", "2026-01-01T10:00:00")
    assert len(result["attendees"][0]["slots"]) == 2
    assert result["attendees"][0]["status"] == "ok"
