from __future__ import annotations

from dataclasses import replace
from types import SimpleNamespace

import pytest

from mailmcp import _admin, _contacts_com, _core, _mail_edit, _mail_ops


def assert_no_mutation(item):
    assert item.Subject == "Synthetic original"
    item.Save.assert_not_called()
    item.Send.assert_not_called()
    item.Delete.assert_not_called()
    item.Move.assert_not_called()
    assert item.Recipients.mutations == []


@pytest.mark.parametrize("item_class", [None, 26, 40, 48, 69])
def test_edit_rejects_non_mail_items(mailbox, item_class):
    mailbox.item.Class = item_class
    with pytest.raises((ValueError, PermissionError)):
        _mail_edit.edit_draft("synthetic-item", subject="Changed", confirm=True)
    assert_no_mutation(mailbox.item)


@pytest.mark.parametrize("folder", ["Inbox", "Contacts", "Calendar"])
@pytest.mark.parametrize("operation", ["edit", "send"])
def test_draft_operations_reject_non_draft_folders(mailbox, folder, operation):
    mailbox.item.Parent = mailbox.folders[("a", folder)]
    with pytest.raises((ValueError, PermissionError)):
        if operation == "edit":
            _mail_edit.edit_draft("synthetic-item", subject="Changed", confirm=True)
        else:
            _mail_ops.send_mail("synthetic-item", confirm=True)
    assert_no_mutation(mailbox.item)


@pytest.mark.parametrize("operation", ["edit", "send"])
@pytest.mark.parametrize("sent", [True, None, "false"])
def test_draft_operations_fail_closed_on_sent_state(mailbox, operation, sent):
    mailbox.item.Sent = sent
    with pytest.raises((ValueError, PermissionError)):
        if operation == "edit":
            _mail_edit.edit_draft("synthetic-item", subject="Changed", confirm=True)
        else:
            _mail_ops.send_mail("synthetic-item", confirm=True)
    assert_no_mutation(mailbox.item)


@pytest.mark.parametrize("operation", ["edit", "send"])
def test_default_drafts_not_a_same_named_other_store(mailbox, operation):
    mailbox.item.Parent = mailbox.folders[("b", "Drafts")]
    with pytest.raises((ValueError, PermissionError)):
        if operation == "edit":
            _mail_edit.edit_draft("synthetic-item", subject="Changed", confirm=True)
        else:
            _mail_ops.send_mail("synthetic-item", confirm=True)
    assert_no_mutation(mailbox.item)


@pytest.mark.parametrize("operation", ["edit", "send"])
def test_drafts_access_is_explicit_not_implicitly_granted(mailbox, operation):
    _core.set_config(replace(_core.get_config(), allowlist_folders=["Inbox", "Contacts"]))
    with pytest.raises(PermissionError):
        if operation == "edit":
            _mail_edit.edit_draft("synthetic-item", subject="Changed", confirm=True)
        else:
            _mail_ops.send_mail("synthetic-item", confirm=True)
    assert_no_mutation(mailbox.item)


def test_edit_and_send_valid_draft_with_all_explicit_gates(mailbox):
    result = _mail_edit.edit_draft("synthetic-item", subject="Changed", confirm=True)
    assert result["status"] == "draft_updated"
    mailbox.item.Save.assert_called_once()
    assert _mail_ops.send_mail("synthetic-item", confirm=True)["status"] == "sent"
    mailbox.item.Send.assert_called_once()


@pytest.mark.parametrize("operation", ["send", "move", "flag", "read", "junk", "delete"])
def test_message_mutators_reject_contacts(mailbox, operation):
    mailbox.item.Class = 40
    calls = {
        "send": lambda: _mail_ops.send_mail("synthetic-item", confirm=True),
        "move": lambda: _mail_ops.move_message("synthetic-item", "Inbox", confirm=True),
        "flag": lambda: _mail_ops.flag_message("synthetic-item", "flagged", confirm=True),
        "read": lambda: _mail_ops.mark_read("synthetic-item", confirm=True),
        "junk": lambda: _mail_ops.mark_junk("synthetic-item", confirm=True),
        "delete": lambda: _admin.delete_message("synthetic-item", confirm=True),
    }
    with pytest.raises((ValueError, PermissionError)):
        calls[operation]()
    assert_no_mutation(mailbox.item)


def test_contact_cross_account_is_denied_before_return(mailbox):
    mailbox.item.Class = 40
    mailbox.item.Parent = mailbox.folders[("b", "Contacts")]
    with pytest.raises(PermissionError):
        _contacts_com.get_contact("synthetic-item", account_email="owner-a@example.com")


def test_contact_in_disallowed_actual_folder_is_denied(mailbox):
    mailbox.item.Class = 40
    mailbox.item.Parent = mailbox.folders[("a", "Inbox")]
    with pytest.raises(PermissionError):
        _contacts_com.get_contact("synthetic-item")


def test_contact_default_folder_positive(mailbox):
    mailbox.item.Class = 40
    mailbox.item.Parent = mailbox.folders[("a", "Contacts")]
    assert _contacts_com.get_contact("synthetic-item")["entry_id"] == "synthetic-item"


def test_draft_folder_missing_identity_fails_closed(mailbox):
    mailbox.item.Parent = SimpleNamespace(Name="Drafts", EntryID=None, StoreID=None)
    with pytest.raises(PermissionError):
        _mail_edit.edit_draft("synthetic-item", subject="Changed", confirm=True)
    assert_no_mutation(mailbox.item)
