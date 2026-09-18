from __future__ import annotations

import pytest

from mailmcp import _admin


def test_soft_delete_cannot_purge_deleted_items(mailbox):
    mailbox.item.Parent = mailbox.folders[("a", "Deleted Items")]
    with pytest.raises(ValueError, match="permanent=True"):
        _admin.delete_message("synthetic-item", confirm=True)
    mailbox.item.Delete.assert_not_called()
    mailbox.item.Move.assert_not_called()


def test_soft_delete_normal_mail_preserves_trash_semantics(mailbox):
    mailbox.item.Parent = mailbox.folders[("a", "Inbox")]
    result = _admin.delete_message("synthetic-item", confirm=True)
    assert result["permanent"] is False
    mailbox.item.Delete.assert_called_once_with()


def test_hard_delete_uses_source_store_trash(mailbox):
    mailbox.item.Parent = mailbox.folders[("b", "Inbox")]
    result = _admin.delete_message("synthetic-item", permanent=True, confirm=True)
    mailbox.item.Move.assert_called_once_with(mailbox.folders[("b", "Deleted Items")])
    assert result["permanent"] is True


def test_missing_store_identity_blocks_delete(mailbox):
    mailbox.item.Parent.Store = None
    with pytest.raises(PermissionError):
        _admin.delete_message("synthetic-item", confirm=True)
    mailbox.item.Delete.assert_not_called()
    mailbox.item.Move.assert_not_called()
