"""Fail-closed item and folder checks shared by Outlook operations."""
from __future__ import annotations

from typing import Any

from mailmcp import _core, _folders


def _assert_mail_item(item: Any) -> None:
    if getattr(item, "Class", None) != 43:
        raise ValueError("The requested item is not a MailItem.")


def _folder_identity(folder: Any) -> tuple[str, str]:
    try:
        entry_id = folder.EntryID
        store_id = folder.StoreID
    except Exception as exc:  # noqa: BLE001 - COM identity access must fail closed.
        raise PermissionError("Cannot verify Outlook folder identity.") from exc
    if not all(isinstance(value, str) and value.strip() for value in (entry_id, store_id)):
        raise PermissionError("Cannot verify Outlook folder identity.")
    return entry_id.strip().lower(), store_id.strip().lower()


def _assert_default_item_folder(
    item: Any, folder_id: int, folder_name: str, account_email: str | None,
) -> None:
    if account_email is not None and (
        not isinstance(account_email, str) or not account_email.strip()
    ):
        raise ValueError("account_email must be a non-empty email address")
    _core._assert_allowed(folder_name, account_email)
    _folders._assert_object_belongs_to_account(item, account_email)
    try:
        # No display-name fallback: security checks require the real default folder.
        expected = _folders._get_default_folder_for_account(folder_id, account_email)
        parent = item.Parent
    except Exception as exc:  # noqa: BLE001 - COM identity access must fail closed.
        raise PermissionError("Cannot verify the item's default folder.") from exc
    if _folder_identity(parent) != _folder_identity(expected):
        raise PermissionError(f"The item is not in the selected account's {folder_name} folder.")


def _assert_draft_item(item: Any, account_email: str | None = None) -> None:
    _assert_mail_item(item)
    sent = getattr(item, "Sent", None)
    if type(sent) not in (bool, int) or sent != 0:
        raise PermissionError("Only a verified unsent mail draft may be edited or sent.")
    _assert_default_item_folder(item, _folders._OL_DRAFTS, "Drafts", account_email)


def _assert_calendar_item_folder(item: Any, account_email: str | None = None) -> None:
    _folders._assert_object_belongs_to_account(item, account_email, object_label="calendar item")
    try:
        parent = item.Parent
        parent_identity = _folder_identity(parent)
        default = _folders._get_default_folder_for_account(_folders._OL_CALENDAR, account_email)
        default_identity = _folder_identity(default)
        name = "Calendar" if parent_identity == default_identity else parent.Name
    except Exception as exc:  # noqa: BLE001 - COM identity access must fail closed.
        raise PermissionError("Cannot verify calendar item folder.") from exc
    if not isinstance(name, str) or not name.strip():
        raise PermissionError("Cannot verify calendar item folder.")
    _core._assert_allowed(name, account_email)
