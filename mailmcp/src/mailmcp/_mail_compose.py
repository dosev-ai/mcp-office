"""Draft composition: compose, reply, and forward operations."""
from __future__ import annotations

import logging
import re

from mailmcp import _core
from mailmcp import _folders
from mailmcp._core import _assert_allowed, _EMAIL_VALIDATE_RE, _assert_domains_allowed, _resolve_smtp_from_entry
from mailmcp._formatters import _text_to_html

_ARTIFACT_ID_PREFIX_MAIL_DRAFT = "mail:draft-"
logger = logging.getLogger(__name__)


def compose_mail(to: list[str], subject: str, body: str, html_body: str | None = None, cc: list[str] | None = None, attachment_paths: list[str] | None = None, account_email: str | None = None, confirm: bool = False, source_entry_id: str | None = None) -> dict:
    _core._assert_write_enabled(account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True is required to create a draft. This is a safety gate.")
    if not to:
        raise ValueError("'to' must contain at least one recipient")
    for addr in to + (cc or []):
        if not _EMAIL_VALIDATE_RE.match(addr):
            raise ValueError(f"Invalid email address: {addr!r}")
    _assert_domains_allowed(to + (cc or []), account_email=account_email)
    validated_attachment_paths = [_folders._validate_attachment_path(p) for p in (attachment_paths or [])]
    drafts = _folders._get_drafts_folder(account_email=account_email)
    mail = drafts.Items.Add()
    mail.Subject = subject
    mail.To = "; ".join(to)
    if cc:
        mail.CC = "; ".join(cc)
    for path in validated_attachment_paths:
        mail.Attachments.Add(str(path))
    mail.HTMLBody = html_body if html_body is not None else _text_to_html(body)
    mail.Save()
    if not mail.EntryID:
        raise ValueError("mail.Save() did not produce an EntryID; cannot construct artifact_id.")
    result = {"status": "draft_saved", "entry_id": mail.EntryID, "artifact_id": f"{_ARTIFACT_ID_PREFIX_MAIL_DRAFT}{mail.EntryID}", "subject": subject, "to": to}
    if source_entry_id:
        result["source_entry_id"] = source_entry_id
    return result


def _inject_reply_html(item, body: str, html_body: str | None) -> None:
    new_html = html_body if html_body is not None else _text_to_html(body)
    existing = item.HTMLBody or ""
    match = re.search(r"<body[^>]*>", existing, re.IGNORECASE)
    if match:
        insert_pos = match.end()
        item.HTMLBody = existing[:insert_pos] + new_html + "<br><br>" + existing[insert_pos:]
    else:
        item.HTMLBody = new_html + "<br><br>" + existing


def _validate_optional_recipients(cc: list[str] | None, bcc: list[str] | None, account_email: str | None) -> None:
    for addr in (cc or []) + (bcc or []):
        if not _EMAIL_VALIDATE_RE.match(addr):
            raise ValueError(f"Invalid email address: {addr!r}")
    _assert_domains_allowed((cc or []) + (bcc or []), account_email=account_email)


def _resolved_reply_addresses(reply) -> list[str]:
    addresses: list[str] = []
    for i in range(1, reply.Recipients.Count + 1):
        recipient = reply.Recipients.Item(i)
        address_entry = getattr(recipient, "AddressEntry", None)
        if address_entry is None:
            raise PermissionError(
                "Cannot resolve a reply recipient address entry. Draft creation blocked for safety."
            )
        addresses.append(_resolve_smtp_from_entry(address_entry))
    return addresses


def reply_all_draft(entry_id: str, body: str, html_body: str | None = None, cc: list[str] | None = None, bcc: list[str] | None = None, account_email: str | None = None, confirm: bool = False) -> dict:
    if not confirm:
        raise ValueError("confirm=True is required to create a reply draft.")
    _validate_optional_recipients(cc, bcc, account_email)
    _core._assert_write_enabled(account_email=account_email)
    mapi = _core._mapi()
    try:
        orig = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    _folders._assert_object_belongs_to_account(orig, account_email, object_label="message")
    try:
        parent_name = orig.Parent.Name
    except Exception as exc:
        raise ValueError(f"Cannot determine source folder for entry_id={entry_id!r}") from exc
    _assert_allowed(parent_name, account_email)
    reply = orig.ReplyAll()
    auto_addrs = _resolved_reply_addresses(reply)
    if auto_addrs:
        _assert_domains_allowed(auto_addrs, account_email=account_email)
    for addr in cc or []:
        recip = reply.Recipients.Add(addr)
        recip.Type = 2
    for addr in bcc or []:
        recip = reply.Recipients.Add(addr)
        recip.Type = 3
    if cc or bcc:
        reply.Recipients.ResolveAll()
    _inject_reply_html(reply, body, html_body)
    reply.Save()
    if not reply.EntryID:
        raise ValueError("reply.Save() did not produce an EntryID; cannot construct artifact_id.")
    result = {"status": "draft_saved", "entry_id": reply.EntryID, "artifact_id": f"{_ARTIFACT_ID_PREFIX_MAIL_DRAFT}{reply.EntryID}", "subject": reply.Subject, "to": [reply.To] if reply.To else []}
    if cc:
        result["cc"] = cc
    if bcc:
        result["bcc"] = bcc
    return result


def reply_draft(entry_id: str, body: str, html_body: str | None = None, cc: list[str] | None = None, bcc: list[str] | None = None, account_email: str | None = None, confirm: bool = False) -> dict:
    if not confirm:
        raise ValueError("confirm=True is required to create a reply draft.")
    _validate_optional_recipients(cc, bcc, account_email)
    _core._assert_write_enabled(account_email=account_email)
    mapi = _core._mapi()
    try:
        msg = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    _folders._assert_object_belongs_to_account(msg, account_email, object_label="message")
    try:
        parent_name = msg.Parent.Name
    except Exception as exc:
        raise ValueError(f"Cannot determine source folder for entry_id={entry_id!r}") from exc
    _assert_allowed(parent_name, account_email)
    reply = msg.Reply()
    auto_addrs = _resolved_reply_addresses(reply)
    if auto_addrs:
        _assert_domains_allowed(auto_addrs, account_email=account_email)
    for addr in cc or []:
        recip = reply.Recipients.Add(addr)
        recip.Type = 2
    for addr in bcc or []:
        recip = reply.Recipients.Add(addr)
        recip.Type = 3
    if cc or bcc:
        reply.Recipients.ResolveAll()
    _inject_reply_html(reply, body, html_body)
    reply.Save()
    if not reply.EntryID:
        raise ValueError("reply.Save() did not produce an EntryID; cannot construct artifact_id.")
    result = {"status": "reply_draft_saved", "entry_id": reply.EntryID, "artifact_id": f"{_ARTIFACT_ID_PREFIX_MAIL_DRAFT}{reply.EntryID}", "subject": reply.Subject, "to": [reply.To] if reply.To else []}
    if cc:
        result["cc"] = cc
    if bcc:
        result["bcc"] = bcc
    return result


def forward_mail(entry_id: str, to: list[str], body: str, html_body: str | None = None, cc: list[str] | None = None, attachment_paths: list[str] | None = None, confirm: bool = False, account_email: str | None = None) -> dict:
    _core._assert_write_enabled(account_email=account_email)
    if not to:
        raise ValueError("'to' must contain at least one recipient")
    for addr in to + (cc or []):
        if not _EMAIL_VALIDATE_RE.match(addr):
            raise ValueError(f"Invalid email address: {addr!r}")
    _assert_domains_allowed(to + (cc or []), account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True is required to forward a message. This is a safety gate.")
    validated_paths = [_folders._validate_attachment_path(p) for p in (attachment_paths or [])]
    mapi = _core._mapi()
    try:
        msg = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    _folders._assert_object_belongs_to_account(msg, account_email, object_label="message")
    try:
        parent_name = msg.Parent.Name
    except Exception as exc:
        raise ValueError(f"Cannot determine source folder for entry_id={entry_id!r}") from exc
    _assert_allowed(parent_name, account_email)
    try:
        fwd = msg.Forward()
    except Exception as exc:
        raise RuntimeError(f"Could not create forward draft for message {entry_id!r}: {exc}") from exc
    if fwd is None:
        raise ValueError(f"Outlook returned no forward draft for message {entry_id!r}")
    while fwd.Recipients.Count > 0:
        fwd.Recipients.Remove(1)
    for addr in to:
        recip = fwd.Recipients.Add(addr)
        recip.Type = 1
    for addr in cc or []:
        recip = fwd.Recipients.Add(addr)
        recip.Type = 2
    fwd.Recipients.ResolveAll()
    _inject_reply_html(fwd, body, html_body)
    for path in validated_paths:
        fwd.Attachments.Add(str(path))
    fwd.Save()
    if not fwd.EntryID:
        raise ValueError("fwd.Save() did not produce an EntryID; cannot construct artifact_id.")
    return {"status": "draft_saved", "entry_id": fwd.EntryID, "artifact_id": f"{_ARTIFACT_ID_PREFIX_MAIL_DRAFT}{fwd.EntryID}", "subject": fwd.Subject, "to": to, "cc": cc or []}
