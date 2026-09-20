"""Message listing, reading, folder discovery, stats and thread retrieval."""
from __future__ import annotations

from mailmcp import outlook_com as ol


def outlook_health() -> dict:
    return ol.health()


def outlook_list_accounts() -> dict:
    return {"accounts": ol.list_accounts()}


def outlook_list_folders(
    store_name: str | None = None,
    depth: int = 1,
    account_email: str | None = None,
) -> dict:
    return {
        "folders": ol.list_folders(
            store_name=store_name,
            depth=depth,
            account_email=account_email,
        )
    }


def outlook_list_messages(
    folder_name: str = "Inbox", top: int = 20, unread_only: bool = False,
    since: str | None = None, until: str | None = None,
    has_attachments: bool | None = None, sort_desc: bool = True,
    account_email: str | None = None,
) -> dict:
    result = ol.list_messages(
        folder_name=folder_name, top=top, unread_only=unread_only,
        since=since, until=until, has_attachments=has_attachments,
        sort_desc=sort_desc, account_email=account_email,
    )
    return {"messages": result, "count": len(result)}


def outlook_get_message(entry_id: str, include_body: bool = False, account_email: str | None = None) -> dict:
    return ol.get_message(entry_id=entry_id, include_body=include_body, account_email=account_email)


def outlook_get_mailbox_stats(folder_path: str | None = None, account_email: str | None = None) -> dict:
    return ol.get_mailbox_stats(folder_path=folder_path, account_email=account_email)


def outlook_get_conversation_thread(
    conversation_id: str | None = None,
    max_items: int = 50,
    entry_id: str | None = None,
    account_email: str | None = None,
) -> dict:
    return ol.get_conversation_thread(
        conversation_id=conversation_id,
        max_items=max_items,
        entry_id=entry_id,
        account_email=account_email,
    )
