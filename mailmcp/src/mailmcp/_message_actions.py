"""Compound reply, meeting-request and message-action handlers."""
from __future__ import annotations

import os
from mailmcp import outlook_com as ol


def outlook_reply(
    operation: str, entry_id: str, body: str, html_body: str | None = None,
    cc: list[str] | None = None, bcc: list[str] | None = None,
    account_email: str | None = None, confirm: bool = False,
) -> dict:
    if not confirm:
        raise ValueError("confirm=True is required to create a reply draft.")
    op = operation.lower()
    if op == "all":
        return ol.reply_all_draft(entry_id=entry_id, body=body, html_body=html_body, cc=cc, bcc=bcc, account_email=account_email, confirm=confirm)
    if op == "sender":
        return ol.reply_draft(entry_id=entry_id, body=body, html_body=html_body, cc=cc, bcc=bcc, account_email=account_email, confirm=confirm)
    raise ValueError(f"Unknown operation: {operation!r}. Must be one of: all, sender")


def handle_meeting_request(
    operation: str, entry_id: str | None = None, response: str | None = None,
    top: int = 20, folder_name: str | None = None, confirm: bool = False,
    account_email: str | None = None,
) -> dict:
    op = operation.lower()
    if op == "list":
        return ol.list_meeting_requests(top=top, folder_name=folder_name, account_email=account_email)
    if op in ("respond", "accept", "decline", "tentative"):
        if entry_id is None:
            raise ValueError("entry_id is required to respond to a meeting request")
        actual_response = op if op != "respond" else response
        if actual_response is None:
            raise ValueError("response is required for operation='respond'")
        return ol.respond_to_meeting(entry_id=entry_id, response=actual_response, confirm=confirm, account_email=account_email)
    raise ValueError(f"Unknown operation: {operation!r}. Must be one of: list, respond, accept, decline, tentative")


def outlook_meeting_request(
    operation: str, entry_id: str | None = None, response: str | None = None,
    top: int = 20, folder_name: str | None = None, confirm: bool = False,
    account_email: str | None = None,
) -> dict:
    return handle_meeting_request(
        operation=operation, entry_id=entry_id, response=response, top=top,
        folder_name=folder_name, confirm=confirm, account_email=account_email,
    )


def handle_message_action(
    operation: str, entry_id: str, target_folder: str | None = None,
    flag_status: str | None = None, permanent: bool = False,
    confirm: bool = False, account_email: str | None = None,
) -> dict:
    op = operation.lower()
    if op == "move":
        if target_folder is None:
            raise ValueError("target_folder is required for operation='move'")
        return ol.move_message(entry_id=entry_id, target_folder=target_folder, confirm=confirm, account_email=account_email)
    if op == "flag":
        if flag_status is None:
            raise ValueError("flag_status is required for operation='flag'")
        return ol.flag_message(entry_id=entry_id, flag_status=flag_status, confirm=confirm, account_email=account_email)
    if op == "mark_read":
        return ol.mark_read(entry_id=entry_id, read=True, confirm=confirm, account_email=account_email)
    if op == "mark_unread":
        return ol.mark_read(entry_id=entry_id, read=False, confirm=confirm, account_email=account_email)
    if op == "delete":
        if not confirm:
            raise ValueError("confirm=True is required for operation='delete'")
        return ol.delete_message(entry_id=entry_id, permanent=permanent, confirm=confirm, account_email=account_email)
    if op == "mark_junk":
        if not confirm:
            raise ValueError("confirm=True is required for operation='mark_junk'")
        return ol.mark_junk(entry_id=entry_id, confirm=confirm, account_email=account_email)
    raise ValueError(f"Unknown operation: {operation!r}. Must be one of: move, flag, mark_read, mark_unread, delete, mark_junk")


def outlook_message_action(
    operation: str, entry_id: str, target_folder: str | None = None,
    flag_status: str | None = None, read: bool = True, permanent: bool = False,
    confirm: bool = False, account_email: str | None = None,
) -> dict:
    del read
    return handle_message_action(
        operation=operation, entry_id=entry_id, target_folder=target_folder,
        flag_status=flag_status, permanent=permanent, confirm=confirm,
        account_email=account_email,
    )


def outlook_bulk_message_action(
    entry_ids: list[str], operation: str, target_folder: str | None = None,
    flag_status: str | None = None, permanent: bool = False,
    confirm: bool = False, account_email: str | None = None,
) -> dict:
    if not confirm:
        raise ValueError("confirm=True is required for bulk message actions. This safety gate prevents accidental mutations.")
    if not entry_ids:
        raise ValueError("entry_ids must be a non-empty list of EntryIDs.")
    operation_normalized = operation.strip().lower()
    valid_ops = {"move", "flag", "mark_read", "mark_unread", "delete", "mark_junk"}
    if operation_normalized not in valid_ops:
        raise ValueError(f"Unknown operation: {operation!r}. Valid: {', '.join(sorted(valid_ops))}")
    if operation_normalized == "move" and target_folder is None:
        raise ValueError("target_folder is required for operation='move'")
    if operation_normalized == "flag" and flag_status is None:
        raise ValueError("flag_status is required for operation='flag'")
    try:
        bulk_limit = max(1, min(int(os.environ.get("OUTLOOK_BULK_LIMIT", "50")), 200))
    except (ValueError, TypeError):
        bulk_limit = 50
    if len(entry_ids) > bulk_limit:
        raise ValueError(f"entry_ids length {len(entry_ids)} exceeds bulk limit {bulk_limit}. Set OUTLOOK_BULK_LIMIT env var to increase (max 200).")
    results: list[dict] = []
    succeeded = 0
    failed = 0
    for eid in entry_ids:
        try:
            result = handle_message_action(
                operation=operation_normalized, entry_id=eid, target_folder=target_folder,
                flag_status=flag_status, permanent=permanent, confirm=confirm,
                account_email=account_email,
            )
            results.append({"entry_id": eid, "ok": True, "result": result})
            succeeded += 1
        except PermissionError:
            raise
        except Exception as exc:
            results.append({"entry_id": eid, "ok": False, "error": str(exc)})
            failed += 1
    return {"total": len(entry_ids), "succeeded": succeeded, "failed": failed, "results": results}
