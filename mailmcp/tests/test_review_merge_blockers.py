from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
import mailmcp
import mailmcp.outlook_com as outlook_com

from mailmcp import (
    _calendar,
    _core,
    _folders,
    _mail_ops,
    _message_actions,
    _message_fetch,
    _message_save,
    _messages,
    _tasks,
)


def test_send_rejects_cross_account_sendusingaccount(mailbox):
    _core._account_overrides.update({
        "owner-a@example.com": _core.OutlookAccountOverride(
            email="owner-a@example.com",
            allowlist_folders=["Drafts"],
            enable_send=True,
            allowlist_domains=["example.com"],
        ),
        "owner-b@example.com": _core.OutlookAccountOverride(
            email="owner-b@example.com",
            allowlist_folders=["Drafts"],
            enable_send=False,
            allowlist_domains=["blocked.example"],
        ),
    })
    mailbox.item.SendUsingAccount = mailbox.mapi.Accounts[1]

    with pytest.raises(PermissionError, match="sending account"):
        _mail_ops.send_mail(
            "synthetic-item",
            confirm=True,
            account_email="owner-a@example.com",
        )

    mailbox.item.Send.assert_not_called()


def test_wildcard_discovery_keeps_raw_names_internal_and_redacts_public(
    mailbox, monkeypatch
):
    _core.set_config(
        replace(
            _core.get_config(),
            allowlist_folders=["*"],
            redact_mode="emails",
        )
    )
    child = SimpleNamespace(
        Name="person@example.com",
        UnReadItemCount=1,
        Items=SimpleNamespace(Count=2),
        Folders=[],
    )
    monkeypatch.setattr(
        mailbox.stores[0],
        "GetRootFolder",
        lambda: SimpleNamespace(Folders=[child]),
    )

    public_rows = _message_fetch.list_folders(depth=2)
    raw_rows = _message_fetch._list_folders_raw(depth=2)

    assert public_rows[0]["name"] == "[email]"
    assert raw_rows[0]["name"] == "person@example.com"
    assert _messages._all_folder_search_names(None) == ["person@example.com"]


def test_read_only_task_extraction_honors_zero_max_items(mailbox, monkeypatch):
    _core.set_config(
        replace(
            _core.get_config(),
            allowlist_folders=["Inbox"],
            max_items=0,
        )
    )
    mailbox.item.Parent = mailbox.folders[("a", "Inbox")]
    mailbox.item.Body = "TODO: Reply to person@example.com"
    monkeypatch.setattr(_tasks, "_mapi", mailbox.get_mapi)

    result = _tasks.extract_tasks_from_message("synthetic-item")

    assert result["candidates"] == []
    assert result["created"] == []


def test_denied_folder_error_does_not_echo_mailbox_folder_name():
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["Inbox"],
            redact_mode="emails",
        )
    )

    with pytest.raises(PermissionError) as exc_info:
        _core._assert_allowed("person@example.com")

    assert "person@example.com" not in str(exc_info.value)


@pytest.mark.parametrize(
    "filename",
    ["invoice-person@example.com.pdf", "person@example.pdf"],
)
def test_attachment_filename_and_saved_path_follow_redaction(
    mailbox, monkeypatch, tmp_path, filename
):
    _core.set_config(
        replace(
            _core.get_config(),
            allowlist_folders=["Inbox"],
            redact_mode="emails",
        )
    )
    mailbox.item.Parent = mailbox.folders[("a", "Inbox")]

    def save_as_file(path):
        Path(path).write_bytes(b"abc")

    attachment = SimpleNamespace(
        FileName=filename,
        Index=1,
        Size=3,
        SaveAsFile=save_as_file,
    )
    mailbox.item.Attachments = [attachment]
    monkeypatch.setattr(
        _folders,
        "_ALLOWED_ATTACHMENT_ROOTS",
        (tmp_path.resolve(),),
    )

    result = _message_save.save_attachments(
        "synthetic-item",
        str(tmp_path),
    )

    assert result[0]["name"] == "[email].pdf"
    assert "person@example" not in result[0]["path"]
    assert Path(result[0]["path"]).exists()


@pytest.mark.parametrize(
    ("redact_mode", "expected_fragment", "forbidden_fragment"),
    [
        ("none", "blocked.example", None),
        ("emails", "domain policy", "blocked.example"),
    ],
)
def test_calendar_update_validates_preserved_attendees(
    mailbox, monkeypatch, redact_mode, expected_fragment, forbidden_fragment
):
    _core.set_config(
        replace(
            _core.get_config(),
            allowlist_folders=["Calendar"],
            allowlist_domains=["example.com"],
            redact_mode=redact_mode,
        )
    )
    mailbox.item.Class = 26
    mailbox.item.Parent = mailbox.folders[("a", "Calendar")]

    class Recipients:
        def __init__(self):
            self.rows = [
                SimpleNamespace(
                    Type=2,
                    AddressEntry=SimpleNamespace(Address="kept@blocked.example"),
                )
            ]

        @property
        def Count(self):
            return len(self.rows)

        def Item(self, index):
            return self.rows[index - 1]

        def Remove(self, index):
            del self.rows[index - 1]

        def Add(self, address):
            row = SimpleNamespace(
                Type=1,
                AddressEntry=SimpleNamespace(Address=address),
            )
            self.rows.append(row)
            return row

        def ResolveAll(self):
            return True

    mailbox.item.Recipients = Recipients()
    monkeypatch.setattr(
        _core,
        "_resolve_smtp_from_entry",
        lambda entry: entry.Address,
    )

    with pytest.raises(PermissionError) as exc_info:
        _calendar.update_calendar_event(
            "synthetic-item",
            required_attendees=["new@example.com"],
            confirm=True,
        )

    message = str(exc_info.value)
    assert expected_fragment in message
    if forbidden_fragment is not None:
        assert forbidden_fragment not in message
        assert "kept@blocked.example" not in message
    mailbox.item.Save.assert_not_called()


def test_conversation_response_redacts_raw_folder_name(monkeypatch):
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["*"],
            redact_mode="emails",
        )
    )
    message = SimpleNamespace(
        EntryID="message-1",
        ConversationID="conversation-1",
        ConversationTopic="Synthetic",
        Subject="Synthetic",
        SenderName="Synthetic Sender",
        SenderEmailAddress="sender@example.com",
        SenderEmailType="SMTP",
        ReceivedTime=None,
        Body="Synthetic body",
    )

    class Items:
        def __iter__(self):
            return iter([message])

        def Restrict(self, _query):
            return [message]

    folder = SimpleNamespace(Items=Items())
    monkeypatch.setattr(
        _messages,
        "_all_folder_search_names",
        lambda _account_email: ["person@example.com"],
    )
    monkeypatch.setattr(
        _mail_ops._folders,
        "_folder_by_name",
        lambda _name: folder,
    )

    result = _mail_ops.get_conversation_thread(
        conversation_id="conversation-1",
        max_items=1,
    )

    assert result["messages"][0]["folder_name"] == "[email]"


def test_account_listing_redacts_configured_folder_names(mailbox):
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["*"],
            redact_mode="none",
        )
    )
    _core._account_overrides["owner-a@example.com"] = (
        _core.OutlookAccountOverride(
            email="owner-a@example.com",
            allowlist_folders=["Inbox", "person@example.com"],
            redact_mode="emails",
        )
    )

    accounts = _message_fetch.list_accounts()

    assert accounts[0]["config_profile"]["allowlist_folders"] == [
        "Inbox",
        "[email]",
    ]


def test_runtime_set_config_is_not_public():
    assert not hasattr(mailmcp, "set_config")
    assert not hasattr(outlook_com, "set_config")
    assert callable(_core.set_config)


def test_mailbox_stats_redacts_configured_folder_name(monkeypatch):
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["person@example.com"],
            redact_mode="emails",
        )
    )
    monkeypatch.setattr(_core, "_account_overrides", {})
    folder = SimpleNamespace(
        UnReadItemCount=1,
        Items=SimpleNamespace(Count=2),
    )
    monkeypatch.setattr(_folders, "_folder_by_name", lambda _name: folder)

    result = _calendar.get_mailbox_stats(folder_path="person@example.com")

    assert result["folders"] == [{"folder": "[email]", "unread": 1, "total": 2}]


def test_public_create_folder_requires_backend_confirmation(monkeypatch):
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["Inbox"],
            enable_write=True,
        )
    )
    monkeypatch.setattr(_core, "_account_overrides", {})
    resolver = Mock()
    monkeypatch.setattr(_folders, "_folder_by_name_for_account", resolver)

    with pytest.raises(ValueError, match="confirm=True"):
        outlook_com.create_folder("Inbox", "Synthetic")

    resolver.assert_not_called()


def test_bulk_permission_failure_returns_explicit_partial_result(monkeypatch):
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["Inbox"],
            redact_mode="emails",
            enable_write=True,
        )
    )
    monkeypatch.setattr(_core, "_account_overrides", {})
    calls: list[str] = []

    def fake_handle_message_action(*, entry_id: str, **_kwargs):
        calls.append(entry_id)
        if entry_id == "blocked":
            raise PermissionError("person@example.com is denied")
        return {"ok": True, "entry_id": entry_id}

    monkeypatch.setattr(
        _message_actions,
        "handle_message_action",
        fake_handle_message_action,
    )

    result = _message_actions.outlook_bulk_message_action(
        ["first", "blocked", "never-called"],
        operation="mark_read",
        confirm=True,
    )

    assert calls == ["first", "blocked"]
    assert result["succeeded"] == 1
    assert result["failed"] == 1
    assert result["skipped"] == 1
    assert result["partial"] is True
    assert result["aborted"] is True
    assert result["results"][1]["permission_denied"] is True
    assert "person@example.com" not in result["results"][1]["error"]
    assert result["results"][2]["skipped"] is True
