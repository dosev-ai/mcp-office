"""MailMCP message handlers and public Outlook composition wrappers."""
from __future__ import annotations

from mailmcp import outlook_com as ol
from mailmcp._msg_rules import handle_forwarding_rule, outlook_forwarding_rule
from mailmcp._message_list import (
    outlook_health,
    outlook_list_accounts,
    outlook_list_folders,
    outlook_list_messages,
    outlook_get_message,
    outlook_get_mailbox_stats,
    outlook_get_conversation_thread,
)
from mailmcp._message_search import (
    outlook_search_messages,
    outlook_search_recipients,
    outlook_search_all_folders,
)
from mailmcp._message_actions import (
    outlook_reply,
    handle_meeting_request,
    outlook_meeting_request,
    handle_message_action,
    outlook_message_action,
    outlook_bulk_message_action,
)

__all__ = [
    "handle_forwarding_rule", "outlook_forwarding_rule",
    "outlook_health", "outlook_list_accounts", "outlook_list_folders",
    "outlook_list_messages", "outlook_get_message", "outlook_get_mailbox_stats",
    "outlook_get_conversation_thread", "outlook_search_messages",
    "outlook_search_recipients", "outlook_search_all_folders", "outlook_reply",
    "handle_meeting_request", "outlook_meeting_request", "handle_message_action",
    "outlook_message_action", "outlook_bulk_message_action", "outlook_compose_mail",
    "outlook_send_mail", "outlook_save_attachments", "outlook_forward_mail",
    "outlook_edit_draft", "outlook_reload_config",
]


def outlook_compose_mail(
    to: list[str], subject: str, body: str, html_body: str | None = None,
    cc: list[str] | None = None, attachment_paths: list[str] | None = None,
    account_email: str | None = None, confirm: bool = False,
    source_entry_id: str | None = None,
) -> dict:
    return ol.compose_mail(
        to=to, subject=subject, body=body, html_body=html_body, cc=cc,
        attachment_paths=attachment_paths, account_email=account_email,
        confirm=confirm, source_entry_id=source_entry_id,
    )


def outlook_edit_draft(
    entry_id: str, subject: str | None = None, body: str | None = None,
    html_body: str | None = None, to: list[str] | None = None,
    cc: list[str] | None = None, bcc: list[str] | None = None,
    account_email: str | None = None, confirm: bool = False,
) -> dict:
    return ol.edit_draft(
        entry_id=entry_id, subject=subject, body=body, html_body=html_body,
        to=to, cc=cc, bcc=bcc, account_email=account_email, confirm=confirm,
    )


def outlook_send_mail(entry_id: str, confirm: bool = False, account_email: str | None = None) -> dict:
    return ol.send_mail(entry_id=entry_id, confirm=confirm, account_email=account_email)


def outlook_save_attachments(entry_id: str, save_dir: str) -> dict:
    result = ol.save_attachments(entry_id=entry_id, save_dir=save_dir)
    return {"saved": result, "count": len(result)}


def outlook_forward_mail(
    entry_id: str, to: list[str], body: str, html_body: str | None = None,
    cc: list[str] | None = None, attachment_paths: list[str] | None = None,
    confirm: bool = False, account_email: str | None = None,
) -> dict:
    return ol.forward_mail(
        entry_id=entry_id, to=to, body=body, html_body=html_body, cc=cc,
        attachment_paths=attachment_paths, confirm=confirm, account_email=account_email,
    )


def outlook_reload_config() -> dict:
    return ol.reload_config()
