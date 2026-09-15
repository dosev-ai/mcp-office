"""mailmcp._context — get_mail_context: per-folder message context with thread grouping.

No MCP/FastMCP imports permitted in this module (layer-separation rule).
COM-only backend; all public functions are called via the outlook_com shim.
"""
from __future__ import annotations

from datetime import timedelta
import logging

from mailmcp._core import _assert_allowed, get_effective_config, _redact
from mailmcp._folders import _folder_by_name_for_account
from mailmcp._formatters import _parse_date, _format_outlook_date, _msg_header

logger = logging.getLogger(__name__)

_ALLOWED_DATE_RANGE_KEYS: frozenset[str] = frozenset({"since", "until"})


def get_mail_context(
    folder: str = "Inbox",
    subject_filter: str | None = None,
    date_range: dict | None = None,
    max_items: int = 50,
    body_preview_chars: int = 200,
    account_email: str | None = None,
) -> dict:
    """Return a structured mail context packet for *folder*, grouped by conversation thread.

    Messages without a ConversationID (e.g. some automated notifications)
    are grouped under a single thread with conversation_id=None.

    Returns:
        dict with keys:
            folder (str): echoed folder name.
            returned_count (int): number of messages collected.
            limit (int): effective cap applied (min of max_items and cfg.max_items).
            maybe_more (bool): True if the cap was reached (more messages may exist).
            filter_applied (bool): True when a date-range COM Restrict filter was
                successfully applied; False when no date range was requested.
            threads (list[dict]): messages grouped by ConversationID. Messages without
                a ConversationID are grouped under thread["conversation_id"] = None.
    """
    _assert_allowed(folder, account_email)

    if date_range is not None:
        if not isinstance(date_range, dict):
            raise ValueError("date_range must be a dict or None.")
        invalid_keys = set(date_range.keys()) - _ALLOWED_DATE_RANGE_KEYS
        if invalid_keys:
            raise ValueError(
                f"date_range contains invalid keys: {sorted(invalid_keys)!r}. "
                f"Allowed keys: {sorted(_ALLOWED_DATE_RANGE_KEYS)!r}."
            )
        for k, v in date_range.items():
            if not isinstance(v, str):
                raise ValueError(
                    f"date_range[{k!r}] must be a string (ISO 8601), "
                    f"got {type(v).__name__!r}."
                )

    cfg = get_effective_config(account_email)
    limit = min(max_items, cfg.max_items)
    preview_chars = min(body_preview_chars, cfg.max_body_chars)

    filters: list[str] = []
    if date_range:
        since_raw = date_range.get("since")
        until_raw = date_range.get("until")
        if since_raw:
            since_dt, _since_date_only = _parse_date(since_raw)
            filters.append(f"([ReceivedTime] >= '{_format_outlook_date(since_dt)}')")
        if until_raw:
            until_dt, until_date_only = _parse_date(until_raw)
            if until_date_only:
                next_day = until_dt + timedelta(days=1)
                filters.append(f"([ReceivedTime] < '{_format_outlook_date(next_day)}')")
            else:
                filters.append(f"([ReceivedTime] <= '{_format_outlook_date(until_dt)}')")

    ol_folder = _folder_by_name_for_account(folder, account_email=account_email)
    items = ol_folder.Items
    items.Sort("[ReceivedTime]", True)

    filter_applied = bool(filters)
    if filters:
        restriction = " AND ".join(filters)
        try:
            items = items.Restrict(restriction)
            items.Sort("[ReceivedTime]", True)
        except Exception as exc:
            raise RuntimeError(
                "Outlook could not apply the requested mail-context date filter; no unfiltered messages were returned."
            ) from exc

    messages: list[dict] = []
    subject_lower = subject_filter.lower() if subject_filter else None
    for msg in items:
        if len(messages) >= limit:
            break
        try:
            hdr = _msg_header(msg, account_email=account_email)
            if subject_lower is not None:
                if subject_lower not in (hdr.get("subject") or "").lower():
                    continue
            body_raw: str = ""
            try:
                body_raw = getattr(msg, "Body", "") or ""
            except Exception as body_exc:
                logger.debug("Body fetch failed for message: %s", body_exc)
            hdr["body_preview"] = _redact(
                body_raw[:preview_chars], account_email=account_email
            )
            messages.append(hdr)
        except Exception as exc:
            logger.warning("Skipping message in get_mail_context: %s", exc)

    maybe_more = len(messages) >= limit
    thread_map: dict[str | None, dict] = {}
    for msg in messages:
        cid = msg.get("conversation_id")
        if cid not in thread_map:
            thread_map[cid] = {
                "conversation_id": cid,
                "conversation_topic": msg.get("conversation_topic") or "",
                "messages": [],
            }
        thread_map[cid]["messages"].append(msg)

    return {
        "folder": folder,
        "returned_count": len(messages),
        "limit": limit,
        "maybe_more": maybe_more,
        "filter_applied": filter_applied,
        "threads": list(thread_map.values()),
    }
