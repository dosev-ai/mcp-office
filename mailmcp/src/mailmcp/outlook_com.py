"""Public Outlook COM shim — re-exports only."""
from __future__ import annotations

import concurrent.futures
import logging
import os

from mailmcp._admin import delete_message
from mailmcp._calendar import get_calendar_event, get_mailbox_stats, list_calendar_events, update_calendar_event
from mailmcp._categories import get_messages_by_category, list_categories, set_message_category
from mailmcp._contacts_com import get_contact, list_contacts, search_contacts
from mailmcp._context import get_mail_context
from mailmcp._core import _DOMAIN_RE, _EMAIL_RE, _EMAIL_VALIDATE_RE, OutlookConfig, _assert_allowed, _assert_write_enabled, _get_outlook, _mapi, _redact, get_config, get_startup_config, reload_config
from mailmcp._folders import _ALLOWED_ATTACHMENT_EXTENSIONS, _ALLOWED_ATTACHMENT_ROOTS, _FLAG_STATUS_MAP, _OL_APPOINTMENT_CLASS, _OL_CALENDAR, _OL_DELETED_ITEMS, _OL_FLAG_COMPLETE, _OL_FLAG_MARKED, _OL_INBOX, _OL_NO_FLAG, _appointment_to_dict, _folder_by_name, _get_calendar_folder, _get_deleted_items_folder, _validate_attachment_path, create_folder
from mailmcp._formatters import _fmt_date, _format_outlook_date, _msg_header, _parse_date, _sql_escape, _text_to_html
from mailmcp._mail_compose import compose_mail, forward_mail, reply_all_draft, reply_draft
from mailmcp._mail_edit import edit_draft
from mailmcp._mail_ops import _FB_STATUS, check_freebusy, create_forwarding_rule, create_meeting_draft, delete_forwarding_rule, flag_message, get_conversation_thread, list_forwarding_rules, mark_junk, mark_read, move_message, send_mail
from mailmcp._messages import get_message as _messages_get_message
from mailmcp._messages import health, list_accounts, list_folders, list_messages, save_attachments, search_all_folders, search_all_folders_detailed, search_messages, search_recipients
from mailmcp._tasks import complete_task, create_task, extract_tasks_from_message, list_meeting_requests, list_tasks, respond_to_meeting

logger = logging.getLogger(__name__)


def get_message(entry_id: str, include_body: bool = False, account_email: str | None = None) -> dict:
    try:
        raw = int(os.getenv("OUTLOOK_GET_MESSAGE_TIMEOUT_SECS", "30"))
    except ValueError:
        raw = 30
    timeout_secs = max(5, min(300, raw))

    def _worker() -> dict:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            return _messages_get_message(entry_id=entry_id, include_body=include_body, account_email=account_email)
        finally:
            pythoncom.CoUninitialize()

    executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
    future = executor.submit(_worker)
    try:
        return future.result(timeout=timeout_secs)
    except concurrent.futures.TimeoutError:
        executor.shutdown(wait=False)
        raise TimeoutError(f"outlook_get_message timed out after {timeout_secs}s for entry_id={entry_id!r}. Outlook may be syncing or unresponsive. Retry or open Outlook manually.")
    except Exception:
        executor.shutdown(wait=False)
        raise
    else:
        executor.shutdown(wait=False)


class _FolderResult(list):
    scan_capped: bool
    def __init__(self, messages: list, scan_capped: bool) -> None:
        super().__init__(messages)
        self.scan_capped = scan_capped


def search_all_folders_worker(folder_name: str, subject: str | None = None, sender: str | None = None, body_contains: str | None = None, top: int | None = None, scan_limit: int | None = None, account_email: str | None = None) -> _FolderResult:
    import pythoncom
    pythoncom.CoInitialize()
    try:
        from mailmcp import _messages
        raw = _messages.search_messages(folder_name=folder_name, subject=subject, sender=sender, body_contains=body_contains, top=top, account_email=account_email, scan_limit=scan_limit, include_scan_metadata=True)
        if isinstance(raw, dict):
            found = list(raw.get("messages") or [])
            scan_capped = bool(raw.get("scan_capped", False))
        else:
            found = list(raw)
            scan_capped = False
        return _FolderResult([dict(msg, folder_name=folder_name) for msg in found], scan_capped)
    finally:
        pythoncom.CoUninitialize()


__all__ = [
    "_FolderResult", "OutlookConfig", "get_config", "get_startup_config", "_EMAIL_RE", "_DOMAIN_RE", "_EMAIL_VALIDATE_RE", "_redact", "_get_outlook", "_mapi", "_assert_allowed", "_assert_write_enabled", "reload_config", "_msg_header", "_fmt_date", "_parse_date", "_format_outlook_date", "_sql_escape", "_text_to_html", "_OL_INBOX", "_OL_DELETED_ITEMS", "_OL_CALENDAR", "_OL_APPOINTMENT_CLASS", "_OL_FLAG_MARKED", "_OL_FLAG_COMPLETE", "_OL_NO_FLAG", "_FLAG_STATUS_MAP", "_ALLOWED_ATTACHMENT_ROOTS", "_ALLOWED_ATTACHMENT_EXTENSIONS", "_folder_by_name", "_get_calendar_folder", "_get_deleted_items_folder", "_appointment_to_dict", "_validate_attachment_path", "health", "list_accounts", "list_folders", "list_messages", "get_message", "search_messages", "save_attachments", "search_recipients", "search_all_folders", "search_all_folders_detailed", "search_all_folders_worker", "_FB_STATUS", "check_freebusy", "create_meeting_draft", "compose_mail", "reply_all_draft", "reply_draft", "edit_draft", "send_mail", "move_message", "flag_message", "mark_read", "get_conversation_thread", "forward_mail", "mark_junk", "list_forwarding_rules", "create_forwarding_rule", "delete_forwarding_rule", "list_calendar_events", "get_calendar_event", "update_calendar_event", "get_mailbox_stats", "delete_message", "list_tasks", "create_task", "complete_task", "list_meeting_requests", "respond_to_meeting", "extract_tasks_from_message", "list_categories", "set_message_category", "get_messages_by_category", "list_contacts", "get_contact", "search_contacts", "create_folder", "get_mail_context",
]
