"""mailmcp._context_ops — Tool handler for outlook_get_mail_context.

No COM logic here. Delegates entirely to the outlook_com shim.
No MCP/FastMCP imports — MCP wiring lives in server.py.
"""
from __future__ import annotations

from mailmcp import outlook_com as _ol


def outlook_get_mail_context(
    folder: str = "Inbox",
    subject_filter: str | None = None,
    date_range: dict | None = None,
    max_items: int = 50,
    body_preview_chars: int = 200,
    account_email: str | None = None,
) -> dict:
    """Return a structured mail context packet for a folder, grouped by conversation thread.

    Retrieves recent messages from *folder*, optionally filtered by subject
    and/or date range, grouped by conversation thread with lightweight previews.

    Use for 'catch me up on my inbox' or 'what's in the Projects folder?'.
    For a full single message body, use outlook_get_message instead.

    Args:
        folder: Outlook folder name; must be in OUTLOOK_ALLOWLIST_FOLDERS.
        subject_filter: Optional case-insensitive substring to match against Subject.
        date_range: Optional dict with keys "since" and/or "until" (ISO 8601 strings).
                    Example: {"since": "2026-03-01", "until": "2026-04-05"}
                    Invalid keys raise ValueError. Non-string values raise ValueError.
        max_items: Max messages to return. Capped by OUTLOOK_MAX_ITEMS env var.
        body_preview_chars: Body preview length. Capped by OUTLOOK_MAX_BODY_CHARS.
        account_email: Optional SMTP address for multi-account Outlook setups.

    Returns:
        {
            "folder": str,
            "returned_count": int,
            "limit": int,
            "maybe_more": bool,
            "threads": [{"conversation_id": str|None, "conversation_topic": str,
                         "messages": [<msg dict with body_preview>]}]
        }
    """
    return _ol.get_mail_context(
        folder=folder,
        subject_filter=subject_filter,
        date_range=date_range,
        max_items=max_items,
        body_preview_chars=body_preview_chars,
        account_email=account_email,
    )
