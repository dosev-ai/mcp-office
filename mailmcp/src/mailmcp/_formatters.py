"""mailmcp._formatters — Message formatting, date, SQL escape, HTML helpers."""
from __future__ import annotations

import html as _html_lib
import re
from datetime import datetime, timezone
from typing import Any

from mailmcp._core import _redact


def _resolve_sender_email(msg: Any) -> str:
    email_type = getattr(msg, "SenderEmailType", "")
    if email_type == "EX":
        try:
            sender = msg.Sender
            if sender is not None:
                exuser = sender.GetExchangeUser()
                if exuser is not None:
                    smtp = exuser.PrimarySmtpAddress
                    if smtp:
                        return smtp
        except Exception:
            pass
    return getattr(msg, "SenderEmailAddress", "") or ""


def _msg_header(msg: Any, account_email: str | None = None) -> dict:
    return {
        "entry_id": msg.EntryID,
        "subject": _redact(msg.Subject or "", account_email=account_email),
        "sender_name": _redact(getattr(msg, "SenderName", "") or "", account_email=account_email),
        "sender_email": _redact(_resolve_sender_email(msg), account_email=account_email),
        "received_time": _fmt_date(getattr(msg, "ReceivedTime", None)),
        "unread": getattr(msg, "UnRead", False) is not False,
        "has_attachments": bool(getattr(msg, "Attachments", None) and msg.Attachments.Count > 0),
        "importance": getattr(msg, "Importance", 1),
        "conversation_topic": _redact(getattr(msg, "ConversationTopic", "") or "", account_email=account_email),
        "conversation_id": getattr(msg, "ConversationID", None) or None,
    }


def _fmt_date(dt: Any) -> str | None:
    if dt is None:
        return None
    try:
        if isinstance(dt, datetime):
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc).isoformat()
        return str(dt)[:19]
    except Exception:
        return None


def _format_outlook_date(dt: datetime) -> str:
    ampm = "PM" if dt.hour >= 12 else "AM"
    hour12 = dt.hour % 12 or 12
    return f"{dt.month}/{dt.day}/{dt.year} {hour12}:{dt.minute:02d} {ampm}"


def _parse_date(s: str) -> tuple[datetime, bool]:
    if '\x00' in s:
        raise ValueError("Date string must not contain null bytes.")
    try:
        date_only = "T" not in s and len(s) == 10
        d = datetime.fromisoformat(s)
        if d.tzinfo is not None:
            d = d.astimezone().replace(tzinfo=None)
        return d, date_only
    except (ValueError, TypeError):
        raise ValueError(f"Invalid date format: {s!r}. Use ISO 8601 (YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS)")


def _sql_escape(s: str) -> str:
    if "\x00" in s:
        raise ValueError("Search/filter string contains null bytes which are not permitted.")
    s = s.replace("[", "[[]")
    s = s.replace("%", "[%]")
    s = s.replace("_", "[_]")
    s = s.replace("'", "''")
    return s


def _text_to_html(text: str) -> str:
    escaped = _html_lib.escape(text)
    escaped = re.sub(r"(?m)^-{3,}\s*$", "<hr>", escaped)
    escaped = re.sub(r"\n{2,}", "</p><p>", escaped)
    escaped = escaped.replace("\n", "<br>")
    return f"<p>{escaped}</p>"
