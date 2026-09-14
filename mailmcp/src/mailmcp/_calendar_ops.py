"""mailmcp._calendar_ops — Tool handler definitions for calendar event operations.

Calendar events, meeting drafts, free/busy checks, and event updates.
Tasks, categories, and meeting requests live in _task_category_ops.py.

Functions here contain the full MCP tool signatures and docstrings; server.py
registers them via the _register() helper which wraps each in _safe() error handling.

No MCP SDK imports here — this module is MCP-framework-agnostic.
"""
from __future__ import annotations

from mailmcp import outlook_com as ol


# ---------------------------------------------------------------------------
# Calendar events
# ---------------------------------------------------------------------------


def outlook_list_calendar_events(
    start: str | None = None,
    end: str | None = None,
    top: int = 20,
    include_cancelled: bool = False,
) -> dict:
    """
    List calendar events (appointments and meetings) in a date/time range.

    Expands recurring events into individual occurrences. Default range is
    from now to +7 days when no start/end are provided.

    Args:
        start: ISO-8601 start datetime, e.g. '2026-03-01T08:00:00'.
               Defaults to the current date/time if omitted.
        end: ISO-8601 end datetime, e.g. '2026-03-07T23:59:59'.
             Defaults to start + 7 days if omitted.
        top: Maximum events to return (default 20, capped by server config).
        include_cancelled: Whether to include cancelled meetings (default False).

    Returns:
        {"events": [...], "count": int}
    """
    result = ol.list_calendar_events(
        start=start,
        end=end,
        top=top,
        include_cancelled=include_cancelled,
    )
    return {"events": result, "count": len(result)}


def outlook_get_calendar_event(
    entry_id: str,
    include_body: bool = False,
    account_email: str | None = None,
) -> dict:
    """
    Fetch a single calendar event by EntryID.

    Verifies the item is an AppointmentItem (Class == 26) — passing a mail
    message EntryID will return a clear error rather than silently returning
    wrong data.

    Args:
        entry_id: The Outlook EntryID of the calendar event (from
                  list_calendar_events).
        include_body: Whether to include the event body/notes text.
        account_email: Optional account email for per-account config enforcement.

    Returns:
        Full AppointmentItem dict:
        {entry_id, subject, start, end, location, organizer, all_day_event,
         meeting_status, is_recurring, required_attendees, optional_attendees,
         categories, sensitivity, importance, [body, body_truncated]}
    """
    return ol.get_calendar_event(
        entry_id=entry_id,
        include_body=include_body,
        account_email=account_email,
    )


# ---------------------------------------------------------------------------
# Meeting scheduling
# ---------------------------------------------------------------------------


def outlook_create_meeting_draft(
    subject: str,
    start: str,
    end: str,
    required: list[str],
    optional: list[str] | None = None,
    body: str | None = None,
    location: str | None = None,
    is_teams: bool = True,
    confirm: bool = False,
    account_email: str | None = None,
) -> dict:
    """
    Create a draft meeting / Teams invite (never sends — saves to calendar as draft).

    Requires OUTLOOK_ENABLE_WRITE=true.

    The invite appears in your Outlook calendar as a draft meeting request.
    When is_teams=True, the Teams add-in attaches a Teams link on save.
    Send it manually from Outlook when ready.

    Args:
        subject: Meeting subject line.
        start: ISO-8601 start datetime (UTC), e.g. '2026-02-24T09:00:00' (= 10:00 CET).
        end: ISO-8601 end datetime (UTC), e.g. '2026-02-24T09:30:00'.
        required: List of required attendee email addresses.
        optional: Optional list of optional attendee email addresses.
        body: Optional plain-text agenda / body.
        location: Optional location string (overridden by Teams link when is_teams=True).
        is_teams: Whether to create a Teams online meeting (default True).
        confirm: Must be True to proceed (safety gate — prevents accidental calendar writes).
        account_email: Optional account email for per-account config enforcement.
    """
    return ol.create_meeting_draft(
        subject=subject,
        start_iso=start,
        end_iso=end,
        required=required,
        optional=optional,
        body=body,
        location=location,
        is_teams=is_teams,
        confirm=confirm,
        account_email=account_email,
    )


def outlook_update_calendar_event(
    entry_id: str,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    body: str | None = None,
    location: str | None = None,
    required_attendees: list[str] | None = None,
    optional_attendees: list[str] | None = None,
    all_day_event: bool | None = None,
    confirm: bool = False,
    account_email: str | None = None,
) -> dict:
    """
    Update an existing calendar event in-place (preserves Exchange tracking).

    Requires OUTLOOK_ENABLE_WRITE=true. Only the provided fields are updated;
    omitted fields remain unchanged.

    Args:
        entry_id: The Outlook EntryID of the calendar event to update.
        subject: New subject line (optional).
        start: New ISO-8601 start datetime, e.g. '2026-03-20T09:00:00' (optional).
        end: New ISO-8601 end datetime (optional).
        body: New plain-text body/notes (optional).
        location: New location string (optional).
        required_attendees: New list of required attendee emails — replaces all existing required attendees (optional).
        optional_attendees: New list of optional attendee emails — replaces all existing optional attendees (optional).
        all_day_event: Whether this is an all-day event (optional).
        confirm: Must be True to proceed (safety gate).
        account_email: Optional account email for per-account config enforcement.
    """
    return ol.update_calendar_event(
        entry_id=entry_id,
        subject=subject,
        start_iso=start,
        end_iso=end,
        body=body,
        location=location,
        required_attendees=required_attendees,
        optional_attendees=optional_attendees,
        all_day_event=all_day_event,
        confirm=confirm,
        account_email=account_email,
    )


def outlook_check_freebusy(
    emails: list[str],
    start: str,
    end: str,
    interval_minutes: int = 30,
) -> dict:
    """
    Check free/busy calendar availability for a list of attendees.

    Uses Exchange free/busy data (Recipient.FreeBusy). Works for internal Exchange
    users and federated external tenants. Returns a status
    per time slot for each attendee.

    Args:
        emails: List of email addresses to check.
        start: ISO-8601 start datetime (UTC), e.g. '2026-02-24T09:00:00' (= 10:00 CET).
        end: ISO-8601 end datetime (UTC), e.g. '2026-02-24T10:00:00'.
        interval_minutes: Granularity per slot (default 30).
    """
    if interval_minutes <= 0:
        raise ValueError("interval_minutes must be a positive integer (got %d)" % interval_minutes)
    return ol.check_freebusy(
        emails=emails,
        start_iso=start,
        end_iso=end,
        interval_minutes=interval_minutes,
    )
