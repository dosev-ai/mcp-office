"""mailmcp._admin — Destructive message operations."""
from __future__ import annotations

import logging

from mailmcp import _core
from mailmcp import _folders
from mailmcp._core import get_effective_config, _assert_allowed

logger = logging.getLogger(__name__)


def delete_message(
    entry_id: str,
    permanent: bool = False,
    confirm: bool = False,
    account_email: str | None = None,
) -> dict:
    """
    Delete a message from Outlook.

    Soft delete (permanent=False, default):
      Calls msg.Delete() which moves the item to Deleted Items. Outlook's
      default behaviour. Requires confirm=True only.

    Hard delete (permanent=True):
      Moves the item to Deleted Items, then immediately calls Delete() on the
      relocated item to permanently purge it from the mailbox. This is the
      correct two-step COM pattern for permanent deletion — calling Delete()
      on an Inbox item always moves it to Deleted Items first; to bypass the
      recycle bin we must delete from Deleted Items.
      Requires OUTLOOK_ENABLE_DELETE=true AND confirm=True (BLK-03).

    The Deleted Items destination uses GetDefaultFolder(3) directly and is
    NEVER validated through _assert_allowed() — it is a system folder, not
    a user-facing allowlisted folder (BLK-05).

    Args:
        entry_id: EntryID of the message to delete.
        permanent: If True, permanently purge (requires enable_delete config).
        confirm: Must be True for any delete operation (safety gate).
        account_email: Optional account email for per-account config enforcement.
    """
    # BLK-03: confirm gate applies to ALL delete modes
    if not confirm:
        raise ValueError(
            "confirm=True is required to delete a message. This is a safety gate."
        )

    cfg = get_effective_config(account_email)

    # BLK-03: all delete modes (soft and permanent) require enable_delete config gate
    if not cfg.enable_delete:
        raise PermissionError(
            "Delete is disabled. "
            "Set OUTLOOK_ENABLE_DELETE=true to enable message deletion."
        )

    mapi = _core._mapi()
    try:
        msg = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        hresult = getattr(exc, 'hresult', None)
        if hresult is not None:
            raise ValueError(f"Message not found (entry_id={entry_id!r}) — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc

    _folders._assert_object_belongs_to_account(msg, account_email, object_label="message")

    try:
        _parent_name = msg.Parent.Name
    except Exception as _exc:
        _hresult = getattr(_exc, "hresult", None)
        if _hresult is not None:
            raise ValueError(
                f"Cannot determine source folder for entry_id={entry_id!r} — COM HRESULT {_hresult:#010x}"
            ) from _exc
        raise ValueError(
            f"Cannot determine source folder for entry_id={entry_id!r}"
        ) from _exc
    _assert_allowed(_parent_name, account_email)  # CBR-002

    if not permanent:
        # Soft delete: Outlook moves item to Deleted Items automatically
        msg.Delete()
        return {"deleted": True, "permanent": False, "entry_id": entry_id}

    # Hard (permanent) delete — two-step to bypass the recycle bin (BLK-05)
    # BLK-05: use GetDefaultFolder(3) directly, never _assert_allowed
    if account_email is not None:
        deleted_folder = _folders._get_deleted_items_folder(account_email=account_email)
    else:
        deleted_folder = _folders._get_deleted_items_folder()

    # Check whether the item is already in the Deleted Items folder.
    # If it is, one Delete() call is sufficient. If not, we must move first
    # then delete from the destination.
    already_in_deleted = False
    try:
        parent = msg.Parent
        if parent is not None and parent.EntryID == deleted_folder.EntryID:
            already_in_deleted = True
    except Exception:
        pass  # if we can't determine parent, fall through to Move() path

    if already_in_deleted:
        msg.Delete()
    else:
        moved = msg.Move(deleted_folder)
        # msg reference is now invalid after Move(); use the returned item
        moved.Delete()

    return {"deleted": True, "permanent": True, "entry_id": entry_id}
