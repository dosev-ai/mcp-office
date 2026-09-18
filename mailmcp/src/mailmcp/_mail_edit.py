"""mailmcp._mail_edit — In-place edit of existing Outlook draft items."""
from __future__ import annotations

import logging

from mailmcp import _core
from mailmcp._core import _EMAIL_VALIDATE_RE, _assert_domains_allowed
from mailmcp._formatters import _text_to_html
from mailmcp._item_guards import _assert_draft_item
from mailmcp._mail_compose import _resolve_and_validate_recipients

logger = logging.getLogger(__name__)


def edit_draft(
    entry_id: str,
    subject: str | None = None,
    body: str | None = None,
    html_body: str | None = None,
    to: list[str] | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    account_email: str | None = None,
    confirm: bool = False,
) -> dict:
    _core._assert_write_enabled(account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True is required to edit a draft. This is a safety gate.")
    all_new_recipients: list[str] = []
    for addr in (to or []):
        if not _EMAIL_VALIDATE_RE.match(addr):
            raise ValueError(f"Invalid email address: {addr!r}")
        all_new_recipients.append(addr)
    for addr in (cc or []):
        if not _EMAIL_VALIDATE_RE.match(addr):
            raise ValueError(f"Invalid email address: {addr!r}")
        all_new_recipients.append(addr)
    for addr in (bcc or []):
        if not _EMAIL_VALIDATE_RE.match(addr):
            raise ValueError(f"Invalid email address: {addr!r}")
        all_new_recipients.append(addr)
    if all_new_recipients:
        _assert_domains_allowed(all_new_recipients, account_email=account_email)
    mapi = _core._mapi()
    try:
        mail = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            raise ValueError(f"Draft not found (entry_id={entry_id!r}) — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Draft not found (entry_id={entry_id!r})") from exc
    _assert_draft_item(mail, account_email)
    if subject is not None:
        mail.Subject = subject
    if html_body is not None:
        mail.HTMLBody = html_body
    elif body is not None:
        mail.HTMLBody = _text_to_html(body)
    if to is not None:
        mail.To = "; ".join(to)
    if cc is not None:
        mail.CC = "; ".join(cc)
    if bcc is not None:
        mail.BCC = "; ".join(bcc)
    _resolve_and_validate_recipients(mail, account_email)
    mail.Save()
    return {"status": "draft_updated", "entry_id": mail.EntryID, "subject": _core._redact(mail.Subject, account_email)}
