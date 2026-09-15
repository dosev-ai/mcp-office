from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from mailmcp import _calendar, _core, _mail_calendar, _message_fetch, _message_save, _tasks
from mailmcp import _folders


def test_config_rejects_negative_limits_and_malformed_domains(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MAX_BODY_CHARS", "-1")
    with pytest.raises(ValueError, match="non-negative"):
        _core.OutlookConfig.from_env()

    monkeypatch.setenv("OUTLOOK_MAX_BODY_CHARS", "4000")
    monkeypatch.setenv("OUTLOOK_ALLOWLIST_DOMAINS", "example,com")
    with pytest.raises(ValueError, match="domain allowlist"):
        _core.OutlookConfig.from_env()

    monkeypatch.delenv("OUTLOOK_ALLOWLIST_DOMAINS")
    monkeypatch.setenv("OUTLOOK_ACCOUNT_1_EMAIL", "owner-a@example.com")
    monkeypatch.setenv("OUTLOOK_ACCOUNT_1_ALLOWLIST_DOMAINS", "example.com,")
    with pytest.raises(ValueError, match="domain allowlist"):
        _core._parse_account_overrides()


def test_default_task_paths_require_tasks_allowlist(mailbox):
    _core.set_config(_core.OutlookConfig(allowlist_folders=["Inbox", "Contacts"], enable_write=True))
    with pytest.raises(PermissionError, match="Tasks"):
        _tasks.list_tasks()
    with pytest.raises(PermissionError, match="Tasks"):
        _tasks.create_task("Synthetic", confirm=True)

    mailbox.item.Class = _folders._OL_TASK_CLASS
    mailbox.item.Parent = SimpleNamespace(Name="Tasks")
    with pytest.raises(PermissionError, match="Tasks"):
        _tasks.complete_task("synthetic-item", confirm=True)


def test_meeting_draft_requires_calendar_allowlist():
    _core.set_config(_core.OutlookConfig(allowlist_folders=["Inbox"], enable_write=True))
    with pytest.raises(PermissionError, match="Calendar"):
        _mail_calendar.create_meeting_draft(
            subject="Synthetic",
            start_iso="2026-09-16T09:00:00+00:00",
            end_iso="2026-09-16T10:00:00+00:00",
            required=["person@example.com"],
            confirm=True,
        )


def test_folder_discovery_is_account_scoped_and_allowlisted(mailbox):
    _core.set_config(_core.OutlookConfig(allowlist_folders=["Inbox"]))
    for folder in mailbox.folders.values():
        folder.UnReadItemCount = 0
        folder.Items.Count = 0
        folder.Folders = []

    result = _message_fetch.list_folders(account_email="owner-a@example.com", depth=2)
    assert result == [{"name": "Inbox", "unread_count": 0, "item_count": 0}]


def test_attachment_save_uses_collision_free_name(mailbox, monkeypatch, tmp_path):
    _core.set_config(_core.OutlookConfig(allowlist_folders=["Inbox"], attachment_max_mb=10))
    mailbox.item.Parent = mailbox.folders[("a", "Inbox")]
    monkeypatch.setattr(_folders, "_ALLOWED_ATTACHMENT_ROOTS", (tmp_path.resolve(),))

    existing = tmp_path / "report.docx"
    existing.write_bytes(b"keep")

    class Attachment:
        Size = 3
        FileName = "report.docx"
        Index = 1

        @staticmethod
        def SaveAsFile(path):
            Path(path).write_bytes(b"new")

    mailbox.item.Attachments = [Attachment()]
    saved = _message_save.save_attachments("synthetic-item", str(tmp_path))

    assert existing.read_bytes() == b"keep"
    assert saved[0]["name"] == "report_1.docx"
    assert (tmp_path / "report_1.docx").read_bytes() == b"new"


def test_calendar_update_preserves_omitted_optional_attendees(mailbox, monkeypatch):
    _core.set_config(_core.OutlookConfig(
        allowlist_folders=["Calendar"],
        enable_write=True,
        allowlist_domains=["example.com"],
    ))
    mailbox.item.Class = _folders._OL_APPOINTMENT_CLASS
    mailbox.item.Parent = mailbox.folders[("a", "Calendar")]
    mailbox.item.Recipients.rows[0].Type = 1
    mailbox.item.Recipients.rows.append(SimpleNamespace(
        AddressEntry=SimpleNamespace(Address="optional@example.com"),
        Type=2,
    ))
    monkeypatch.setattr(_calendar._folders, "_appointment_to_dict", lambda *args, **kwargs: {})

    _calendar.update_calendar_event(
        "synthetic-item",
        required_attendees=["required-new@example.com"],
        confirm=True,
    )

    optional = [
        row.AddressEntry.Address
        for row in mailbox.item.Recipients.rows
        if row.Type == 2
    ]
    required = [
        row.AddressEntry.Address
        for row in mailbox.item.Recipients.rows
        if row.Type == 1
    ]
    assert optional == ["optional@example.com"]
    assert required == ["required-new@example.com"]


def test_task_restrict_failure_does_not_return_unfiltered_tasks(monkeypatch):
    _core.set_config(_core.OutlookConfig(allowlist_folders=["Tasks"]))
    items = SimpleNamespace(
        Sort=Mock(),
        Restrict=Mock(side_effect=RuntimeError("synthetic restrict failure")),
    )
    folder = SimpleNamespace(Items=items)
    monkeypatch.setattr(_tasks, "_mapi", lambda: SimpleNamespace(GetDefaultFolder=lambda _: folder))

    with pytest.raises(RuntimeError, match="no unfiltered tasks"):
        _tasks.list_tasks()


def test_meeting_restrict_failure_does_not_return_arbitrary_mail(monkeypatch):
    _core.set_config(_core.OutlookConfig(allowlist_folders=["Inbox"]))
    items = SimpleNamespace(
        Sort=Mock(),
        Restrict=Mock(side_effect=RuntimeError("synthetic restrict failure")),
    )
    folder = SimpleNamespace(Items=items)
    monkeypatch.setattr(_tasks, "_mapi", lambda: SimpleNamespace(GetDefaultFolder=lambda _: folder))

    with pytest.raises(RuntimeError, match="no unfiltered messages"):
        _tasks.list_meeting_requests()
