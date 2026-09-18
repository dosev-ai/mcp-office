from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from mailmcp import (
    _calendar,
    _categories,
    _core,
    _folders,
    _mail_calendar,
    _mail_compose,
    _mail_edit,
    _mail_ops,
    _message_fetch,
    _messages,
    _task_category_ops,
    _tasks,
)


class _Recipients:
    def __init__(self, addresses=()):
        self.rows = [
            SimpleNamespace(
                AddressEntry=SimpleNamespace(Address=address),
                Type=1,
            )
            for address in addresses
        ]
        self.resolve_all = Mock(return_value=True)

    @property
    def Count(self):
        return len(self.rows)

    def Item(self, index):
        return self.rows[index - 1]

    def Add(self, address):
        row = SimpleNamespace(
            AddressEntry=SimpleNamespace(Address=address),
            Type=1,
        )
        self.rows.append(row)
        return row

    def Remove(self, index):
        del self.rows[index - 1]

    def ResolveAll(self):
        return self.resolve_all()


def _draft_item(addresses=("alias@example.com",)):
    return SimpleNamespace(
        Recipients=_Recipients(addresses),
        Attachments=SimpleNamespace(Add=Mock()),
        HTMLBody="",
        Subject="Synthetic",
        To="alias@example.com",
        EntryID="synthetic-draft",
        Save=Mock(),
    )


def test_shared_resolved_recipient_guard_blocks_alias_domain(monkeypatch, mailbox):
    _core.set_config(
        _core.OutlookConfig(
            allowlist_domains=["allowed.example"],
        )
    )
    mailbox.item.Recipients = _Recipients(["alias@allowed.example"])
    monkeypatch.setattr(
        _mail_compose,
        "_resolve_smtp_from_entry",
        lambda _entry: "resolved@blocked.example",
    )

    with pytest.raises(PermissionError, match="blocked.example"):
        _mail_compose._resolve_and_validate_recipients(mailbox.item)

    mailbox.item.Recipients.resolve_all.assert_called_once_with()


def test_compose_calls_resolved_recipient_guard_before_save(monkeypatch):
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["Drafts"],
            allowlist_domains=["example.com"],
            enable_write=True,
        )
    )
    draft = _draft_item()
    drafts = SimpleNamespace(Items=SimpleNamespace(Add=Mock(return_value=draft)))
    monkeypatch.setattr(
        _mail_compose._folders,
        "_get_drafts_folder",
        lambda account_email=None: drafts,
    )
    guard = Mock(side_effect=PermissionError("resolved recipient blocked"))
    monkeypatch.setattr(
        _mail_compose,
        "_resolve_and_validate_recipients",
        guard,
    )

    with pytest.raises(PermissionError, match="resolved recipient blocked"):
        _mail_compose.compose_mail(
            to=["alias@example.com"],
            subject="Synthetic",
            body="Synthetic",
            confirm=True,
        )

    guard.assert_called_once_with(draft, None)
    draft.Save.assert_not_called()


@pytest.mark.parametrize(
    ("api_name", "factory_name"),
    [
        ("reply_draft", "Reply"),
        ("reply_all_draft", "ReplyAll"),
    ],
)
def test_reply_paths_recheck_resolved_recipients_before_save(
    monkeypatch,
    mailbox,
    api_name,
    factory_name,
):
    mailbox.item.Parent = mailbox.folders[("a", "Inbox")]
    reply = _draft_item()
    setattr(mailbox.item, factory_name, Mock(return_value=reply))
    guard = Mock(side_effect=PermissionError("resolved recipient blocked"))
    monkeypatch.setattr(
        _mail_compose,
        "_resolve_and_validate_recipients",
        guard,
    )

    api = getattr(_mail_compose, api_name)
    with pytest.raises(PermissionError, match="resolved recipient blocked"):
        api(
            "synthetic-item",
            body="Synthetic",
            confirm=True,
        )

    guard.assert_called_once_with(reply, None)
    reply.Save.assert_not_called()


def test_forward_rechecks_resolved_recipients_before_save(
    monkeypatch,
    mailbox,
):
    mailbox.item.Parent = mailbox.folders[("a", "Inbox")]
    fwd = _draft_item([])
    mailbox.item.Forward = Mock(return_value=fwd)
    guard = Mock(side_effect=PermissionError("resolved recipient blocked"))
    monkeypatch.setattr(
        _mail_compose,
        "_resolve_and_validate_recipients",
        guard,
    )

    with pytest.raises(PermissionError, match="resolved recipient blocked"):
        _mail_compose.forward_mail(
            "synthetic-item",
            to=["alias@example.com"],
            body="Synthetic",
            confirm=True,
        )

    guard.assert_called_once_with(fwd, None)
    fwd.Save.assert_not_called()


def test_edit_rechecks_resolved_recipients_before_save(
    monkeypatch,
    mailbox,
):
    mailbox.item.Parent = mailbox.folders[("a", "Drafts")]
    mailbox.item.Recipients = _Recipients(["alias@example.com"])
    guard = Mock(side_effect=PermissionError("resolved recipient blocked"))
    monkeypatch.setattr(
        _mail_edit,
        "_resolve_and_validate_recipients",
        guard,
    )

    with pytest.raises(PermissionError, match="resolved recipient blocked"):
        _mail_edit.edit_draft(
            "synthetic-item",
            to=["alias@example.com"],
            confirm=True,
        )

    guard.assert_called_once_with(mailbox.item, None)
    mailbox.item.Save.assert_not_called()


def test_freebusy_combines_every_overlapped_provider_bucket(
    monkeypatch,
    mailbox,
):
    _core.set_config(
        _core.OutlookConfig(
            allowlist_domains=["example.com"],
            max_items=10,
        )
    )
    monkeypatch.setattr(
        _mail_calendar,
        "_freebusy_local_naive",
        lambda value: value.replace(tzinfo=None),
    )
    freebusy = Mock(return_value=("0" * 19) + "2" + ("0" * 100))
    recipient = SimpleNamespace(
        Resolve=Mock(),
        Resolved=True,
        AddressEntry=SimpleNamespace(Address="person@example.com"),
        Name="Synthetic",
        FreeBusy=freebusy,
    )
    mailbox.mapi.CreateRecipient.return_value = recipient

    result = _mail_calendar.check_freebusy(
        ["person@example.com"],
        "2026-01-01T09:15:00",
        "2026-01-01T09:45:00",
        interval_minutes=30,
    )

    assert result["attendees"][0]["slots"] == [
        {
            "start": "2026-01-01T09:15:00",
            "end": "2026-01-01T09:45:00",
            "status": "Busy",
        }
    ]


def test_account_listing_fails_closed_on_shared_delivery_store(mailbox):
    mailbox.mapi.Accounts[1].DeliveryStore = mailbox.stores[0]

    with pytest.raises(PermissionError, match="uniquely resolve"):
        _message_fetch.list_accounts()


def test_calendar_categories_use_account_redaction():
    _core.set_config(
        _core.OutlookConfig(
            redact_mode="emails+domains",
        )
    )
    appointment = SimpleNamespace(
        EntryID="synthetic-appointment",
        Categories="person@example.com / example.net",
    )

    result = _folders._appointment_to_dict(appointment)

    assert result["categories"] == "[email] / [domain]"


class _Items(list):
    def Sort(self, *args, **kwargs):
        return None

    def Restrict(self, *args, **kwargs):
        return self


def test_task_and_meeting_request_categories_use_account_redaction(monkeypatch):
    task = SimpleNamespace(
        EntryID="synthetic-task",
        Subject="Synthetic",
        Status=0,
        Importance=1,
        DueDate=None,
        Complete=False,
        Categories="person@example.com / example.net",
        Body="Synthetic",
    )
    task_folder = SimpleNamespace(Items=_Items([task]))
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["Tasks"],
            redact_mode="emails+domains",
        )
    )
    monkeypatch.setattr(
        _tasks,
        "_mapi",
        lambda: SimpleNamespace(GetDefaultFolder=lambda _folder_id: task_folder),
    )

    task_result = _tasks.list_tasks()

    assert task_result["tasks"][0]["categories"] == "[email] / [domain]"

    request = SimpleNamespace(
        EntryID="synthetic-request",
        Subject="Synthetic",
        SenderName="Synthetic",
        SenderEmailType="SMTP",
        SenderEmailAddress="sender@example.com",
        ReceivedTime="",
        Start=None,
        End=None,
        Location="",
        ResponseRequested=False,
        Categories="person@example.com / example.net",
    )
    request_folder = SimpleNamespace(Items=_Items([request]))
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["Inbox"],
            redact_mode="emails+domains",
        )
    )
    monkeypatch.setattr(
        _tasks,
        "_mapi",
        lambda: SimpleNamespace(
            GetDefaultFolder=lambda _folder_id: request_folder
        ),
    )

    request_result = _tasks.list_meeting_requests()

    assert request_result["requests"][0]["categories"] == "[email] / [domain]"


def test_send_rechecks_resolved_recipients_before_send(
    monkeypatch,
    mailbox,
):
    mailbox.item.Parent = mailbox.folders[("a", "Drafts")]
    mailbox.item.Recipients = _Recipients(["alias@example.com"])
    guard = Mock(side_effect=PermissionError("resolved recipient blocked"))
    monkeypatch.setattr(
        _mail_ops,
        "_resolve_and_validate_recipients",
        guard,
    )

    with pytest.raises(PermissionError, match="resolved recipient blocked"):
        _mail_ops.send_mail(
            "synthetic-item",
            confirm=True,
        )

    guard.assert_called_once_with(mailbox.item, None)
    mailbox.item.Send.assert_not_called()


def test_unscoped_move_stays_in_source_store(mailbox):
    mailbox.item.Parent = mailbox.folders[("b", "Inbox")]

    _mail_ops.move_message(
        "synthetic-item",
        target_folder="Deleted Items",
        confirm=True,
    )

    mailbox.item.Move.assert_called_once_with(
        mailbox.folders[("b", "Deleted Items")]
    )


def test_unscoped_mark_junk_stays_in_source_store(mailbox):
    mailbox.item.Parent = mailbox.folders[("b", "Inbox")]

    _mail_ops.mark_junk(
        "synthetic-item",
        confirm=True,
    )

    mailbox.item.Move.assert_called_once_with(
        mailbox.folders[("b", "Junk Email")]
    )


def test_search_all_folders_zero_max_items_short_circuits(monkeypatch):
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["Inbox"],
            max_items=0,
        )
    )
    discovery = Mock(side_effect=AssertionError("folder search should not start"))
    monkeypatch.setattr(
        _messages,
        "_all_folder_search_names",
        discovery,
    )

    result = _messages.search_all_folders_detailed(
        query="Synthetic",
    )

    assert result == {
        "messages": [],
        "errors": [],
        "partial_results": False,
        "any_folder_capped": False,
        "folders_searched": 0,
        "folders_total": 0,
    }
    discovery.assert_not_called()


def test_composite_category_list_uses_verified_account_redaction(
    monkeypatch,
    mailbox,
):
    _core.set_config(
        _core.OutlookConfig(
            redact_mode="none",
        )
    )
    _core._account_overrides["owner-a@example.com"] = (
        _core.OutlookAccountOverride(
            email="owner-a@example.com",
            redact_mode="emails",
        )
    )
    mailbox.mapi.Categories = [
        SimpleNamespace(
            Name="person@example.com",
            Color=1,
            ShortcutKey=0,
        )
    ]
    monkeypatch.setattr(
        _categories,
        "_mapi",
        lambda: mailbox.mapi,
    )

    with pytest.raises(PermissionError, match="account_email is required"):
        _task_category_ops.outlook_category("list")

    with pytest.raises(ValueError, match="Account not found"):
        _task_category_ops.outlook_category(
            "list",
            account_email="unknown@example.com",
        )

    result = _task_category_ops.outlook_category(
        "list",
        account_email="owner-a@example.com",
    )

    assert result["categories"][0]["name"] == "[email]"


def test_shared_delivery_store_rejected_for_scoped_operations(mailbox):
    mailbox.mapi.Accounts[1].DeliveryStore = mailbox.stores[0]

    with pytest.raises(PermissionError, match="Cannot resolve Outlook store"):
        _folders._find_store_for_account(
            "owner-a@example.com",
            mapi=mailbox.mapi,
        )

    with pytest.raises(PermissionError, match="does not belong"):
        _folders._assert_object_belongs_to_account(
            mailbox.item,
            "owner-a@example.com",
            object_label="message",
        )


def test_duplicate_account_override_identifier_is_rejected(monkeypatch):
    monkeypatch.setenv(
        "OUTLOOK_ACCOUNT_1_EMAIL",
        "Owner-A@example.com",
    )
    monkeypatch.setenv(
        "OUTLOOK_ACCOUNT_2_EMAIL",
        "owner-a@example.com",
    )

    with pytest.raises(ValueError, match="duplicates an earlier"):
        _core._parse_account_overrides()


def test_folder_discovery_and_wildcard_stats_redact_names(mailbox):
    sensitive = mailbox.folders[("a", "Inbox")]
    sensitive.Name = "person@example.com"
    for suffix in ("a", "b"):
        for name in (
            "Deleted Items",
            "Inbox",
            "Calendar",
            "Contacts",
            "Drafts",
            "Junk Email",
        ):
            folder = mailbox.folders[(suffix, name)]
            folder.UnReadItemCount = 0
            folder.Items.Count = 0

    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["person@example.com"],
            redact_mode="emails",
        )
    )

    discovered = _message_fetch.list_folders(
        account_email="owner-a@example.com",
    )

    assert discovered == [
        {
            "name": "[email]",
            "unread_count": 0,
            "item_count": 0,
        }
    ]

    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["*"],
            redact_mode="emails",
        )
    )
    stats = _calendar.get_mailbox_stats(
        account_email="owner-a@example.com",
    )

    assert any(row["folder"] == "[email]" for row in stats["folders"])
    assert all(
        "person@example.com" not in row["folder"]
        for row in stats["folders"]
    )


def test_conversation_result_scan_has_exact_global_cap(mailbox, monkeypatch):
    _core.set_config(
        _core.OutlookConfig(
            allowlist_folders=["Inbox"],
            max_items=2,
        )
    )
    mailbox.item.Parent = mailbox.folders[("a", "Inbox")]
    mailbox.item.ConversationID = "synthetic-conversation"
    mailbox.item.ConversationTopic = "Synthetic topic"

    class Restricted:
        def __init__(self, total):
            self.total = total
            self.yielded = 0

        def __iter__(self):
            for index in range(self.total):
                self.yielded += 1
                yield SimpleNamespace(
                    ConversationID="synthetic-conversation",
                    EntryID=f"message-{index}",
                    Subject="Synthetic",
                    SenderName="Synthetic",
                    SenderEmailType="SMTP",
                    SenderEmailAddress="person@example.com",
                    ReceivedTime=None,
                    Body="Synthetic body",
                )

    restricted = Restricted(600)

    class Items:
        def Restrict(self, _query):
            return restricted

    monkeypatch.setattr(
        _mail_ops._folders,
        "_folder_by_name",
        lambda _name: SimpleNamespace(Items=Items()),
    )

    result = _mail_ops.get_conversation_thread(
        entry_id="synthetic-item",
        max_items=2,
    )

    assert restricted.yielded == _mail_ops._MAX_THREAD_SCAN_ITEMS
    assert result["count"] == 2
