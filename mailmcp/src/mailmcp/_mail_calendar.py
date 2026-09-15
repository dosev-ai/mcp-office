"""mailmcp._mail_calendar — Calendar/free-busy operations."""
from __future__ import annotations

from datetime import datetime, timezone, timedelta

from mailmcp import _core
from mailmcp._core import _EMAIL_VALIDATE_RE, _redact, _assert_domains_allowed
from mailmcp import _folders

_FB_STATUS = {"0": "Free", "1": "Tentative", "2": "Busy", "3": "OOO", "4": "Working elsewhere"}


def check_freebusy(emails: list[str], start_iso: str, end_iso: str, interval_minutes: int = 30) -> dict:
    try:
        start_dt = datetime.fromisoformat(start_iso)
        end_dt = datetime.fromisoformat(end_iso)
    except ValueError as exc:
        raise ValueError(f"Invalid datetime format: {exc}. Use ISO-8601, e.g. '2026-02-24T09:00:00'") from exc
    if isinstance(interval_minutes, bool) or not isinstance(interval_minutes, int) or interval_minutes <= 0:
        raise ValueError("interval_minutes must be a positive integer")
    try:
        duration = (end_dt - start_dt).total_seconds()
    except TypeError as exc:
        raise ValueError("start_iso and end_iso must use compatible timezones") from exc
    if duration <= 0:
        raise ValueError("end_iso must be after start_iso")
    if any(not isinstance(email, str) for email in emails):
        raise ValueError("Each email must be a string.")
    valid_addresses = [email for email in emails if _EMAIL_VALIDATE_RE.match(email)]
    _assert_domains_allowed(valid_addresses)
    mapi = _core._mapi()
    results = []
    total_minutes = int(duration / 60)
    num_slots = max(1, total_minutes // interval_minutes)
    for email in emails:
        if not _EMAIL_VALIDATE_RE.match(email):
            results.append({"email": _redact(email), "status": "invalid", "slots": []})
            continue
        try:
            recip = mapi.CreateRecipient(email)
            recip.Resolve()
            if not recip.Resolved:
                results.append({"email": _redact(email), "status": "unresolved", "slots": []})
                continue
            midnight = start_dt.replace(hour=0, minute=0, second=0, microsecond=0)
            offset_slots = int((start_dt - midnight).total_seconds() / 60) // interval_minutes
            fb_str = recip.FreeBusy(midnight, interval_minutes, True)
            slots = []
            for i in range(num_slots):
                fb_index = offset_slots + i
                ch = fb_str[fb_index] if fb_index < len(fb_str) else "0"
                slot_start = start_dt + timedelta(minutes=i * interval_minutes)
                slot_end = min(end_dt, slot_start + timedelta(minutes=interval_minutes))
                slots.append({"start": slot_start.isoformat(), "end": slot_end.isoformat(), "status": _FB_STATUS.get(ch, f"?({ch})")})
            results.append({"email": _redact(email), "name": _redact(recip.Name), "status": "ok", "slots": slots})
        except Exception as exc:
            results.append({"email": _redact(email), "status": "error", "error": str(exc)[:120], "slots": []})
    return {"start": start_iso, "end": end_iso, "interval_minutes": interval_minutes, "attendees": results}


def create_meeting_draft(
    subject: str,
    start_iso: str,
    end_iso: str,
    required: list[str],
    optional: list[str] | None = None,
    body: str | None = None,
    location: str | None = None,
    is_teams: bool = True,
    confirm: bool = False,
    account_email: str | None = None,
) -> dict:
    _core._assert_write_enabled(account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True is required to create a meeting draft.")
    try:
        start_utc = datetime.fromisoformat(start_iso)
        end_utc = datetime.fromisoformat(end_iso)
        if start_utc.tzinfo is None:
            start_utc = start_utc.replace(tzinfo=timezone.utc)
        if end_utc.tzinfo is None:
            end_utc = end_utc.replace(tzinfo=timezone.utc)
        start_local = start_utc.astimezone()
        end_local = end_utc.astimezone()
    except ValueError as exc:
        raise ValueError(f"Invalid datetime format: {exc}. Use ISO-8601, e.g. '2026-02-24T09:00:00'") from exc
    if end_utc <= start_utc:
        raise ValueError("end_iso must be after start_iso.")
    for addr in required:
        if not _EMAIL_VALIDATE_RE.match(addr):
            raise ValueError(f"Invalid email address: {addr!r}")
    for addr in (optional or []):
        if not _EMAIL_VALIDATE_RE.match(addr):
            raise ValueError(f"Invalid email address: {addr!r}")
    _assert_domains_allowed(required + (optional or []), account_email=account_email)
    if account_email is not None:
        calendar_folder = _folders._get_calendar_folder(account_email=account_email)
        appt = calendar_folder.Items.Add()
    else:
        app = _core._get_outlook()
        appt = app.CreateItem(1)
    appt.Subject = subject
    appt.Start = start_local.strftime("%Y-%m-%d %H:%M")
    appt.End = end_local.strftime("%Y-%m-%d %H:%M")
    appt.MeetingStatus = 1
    if body:
        appt.Body = body
    if location:
        appt.Location = location
    for email in required:
        recip = appt.Recipients.Add(email)
        recip.Type = 1
    for email in (optional or []):
        recip = appt.Recipients.Add(email)
        recip.Type = 2
    if is_teams:
        try:
            appt.IsOnlineMeeting = True
            appt.OnlineMeetingProvider = 2
        except Exception:
            pass
    appt.Save()
    teams_url = None
    if is_teams:
        try:
            teams_url = appt.OnlineMeetingUrl or None
        except Exception:
            pass
    return {
        "status": "draft_saved",
        "entry_id": appt.EntryID,
        "subject": appt.Subject,
        "start_local": str(appt.Start)[:19],
        "end_local": str(appt.End)[:19],
        "is_teams": is_teams,
        "teams_url": teams_url,
        "required": required,
        "optional": optional or [],
    }
