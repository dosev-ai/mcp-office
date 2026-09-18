"""mailmcp._mail_calendar — Calendar/free-busy operations."""
from __future__ import annotations

from datetime import datetime, timezone, timedelta
import math

from mailmcp import _core
from mailmcp._core import _EMAIL_VALIDATE_RE, _redact, _assert_domains_allowed
from mailmcp import _folders

_FB_STATUS = {"0": "Free", "1": "Tentative", "2": "Busy", "3": "OOO", "4": "Working elsewhere"}
_MAX_FREEBUSY_DURATION_DAYS = 30
_MAX_FREEBUSY_TOTAL_SLOTS = 10_000


def _freebusy_local_naive(value: datetime) -> datetime:
    """Interpret naive inputs as UTC and return the equivalent local wall time."""
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone().replace(tzinfo=None)


def check_freebusy(
    emails: list[str],
    start_iso: str,
    end_iso: str,
    interval_minutes: int = 30,
    account_email: str | None = None,
) -> dict:
    if account_email is None and _core._account_overrides:
        raise PermissionError(
            "account_email is required when per-account Outlook policy is configured."
        )
    try:
        start_input = datetime.fromisoformat(start_iso)
        end_input = datetime.fromisoformat(end_iso)
    except ValueError as exc:
        raise ValueError(f"Invalid datetime format: {exc}. Use ISO-8601, e.g. '2026-02-24T09:00:00'") from exc
    start_aware = start_input if start_input.tzinfo is not None else start_input.replace(tzinfo=timezone.utc)
    end_aware = end_input if end_input.tzinfo is not None else end_input.replace(tzinfo=timezone.utc)
    duration = (end_aware - start_aware).total_seconds()
    if duration <= 0:
        raise ValueError("end_iso must be after start_iso")
    if isinstance(interval_minutes, bool) or not isinstance(interval_minutes, int) or interval_minutes <= 0:
        raise ValueError("interval_minutes must be a positive integer")
    if duration > _MAX_FREEBUSY_DURATION_DAYS * 24 * 60 * 60:
        raise ValueError(f"Free/busy range cannot exceed {_MAX_FREEBUSY_DURATION_DAYS} days.")
    if any(not isinstance(email, str) for email in emails):
        raise ValueError("Each email must be a string.")

    mapi = None
    if account_email is not None:
        mapi = _core._mapi()
        _folders._find_store_for_account(account_email, mapi=mapi)
    cfg = _core.get_effective_config(account_email)
    max_attendees = cfg.max_items
    if len(emails) > max_attendees:
        raise ValueError(f"Free/busy attendee count exceeds configured max_items ({max_attendees}).")
    num_slots = max(1, math.ceil(duration / (interval_minutes * 60)))
    if num_slots * len(emails) > _MAX_FREEBUSY_TOTAL_SLOTS:
        raise ValueError(
            f"Free/busy request exceeds the {_MAX_FREEBUSY_TOTAL_SLOTS}-slot work budget. "
            "Reduce attendees, duration, or increase interval_minutes."
        )
    valid_addresses = [email for email in emails if _EMAIL_VALIDATE_RE.match(email)]
    _assert_domains_allowed(valid_addresses, account_email=account_email)
    if mapi is None:
        mapi = _core._mapi()
    start_dt = _freebusy_local_naive(start_input)
    end_dt = _freebusy_local_naive(end_input)
    results = []
    for email in emails:
        if not _EMAIL_VALIDATE_RE.match(email):
            results.append({"email": _redact(email, account_email), "status": "invalid", "slots": []})
            continue
        try:
            recip = mapi.CreateRecipient(email)
            recip.Resolve()
            if not recip.Resolved:
                results.append({"email": _redact(email, account_email), "status": "unresolved", "slots": []})
                continue
            address_entry = getattr(recip, "AddressEntry", None)
            if address_entry is None:
                raise PermissionError("Cannot verify the resolved free/busy recipient.")
            resolved_smtp = _core._resolve_smtp_from_entry(address_entry)
            _assert_domains_allowed([resolved_smtp], account_email=account_email)
            midnight = start_dt.replace(hour=0, minute=0, second=0, microsecond=0)
            offset_slots = int((start_dt - midnight).total_seconds() / 60) // interval_minutes
            fb_str = recip.FreeBusy(midnight, interval_minutes, True)
            if offset_slots + num_slots > len(fb_str):
                raise RuntimeError("Outlook returned incomplete free/busy data for the requested range.")
            slots = []
            for i in range(num_slots):
                fb_index = offset_slots + i
                ch = fb_str[fb_index]
                slot_start = start_dt + timedelta(minutes=i * interval_minutes)
                slot_end = min(end_dt, slot_start + timedelta(minutes=interval_minutes))
                slots.append({"start": slot_start.isoformat(), "end": slot_end.isoformat(), "status": _FB_STATUS.get(ch, f"?({ch})")})
            results.append({
                "email": _redact(email, account_email),
                "name": _redact(recip.Name, account_email),
                "status": "ok",
                "slots": slots,
            })
        except PermissionError:
            results.append({
                "email": _redact(email, account_email),
                "status": "error",
                "error": "Resolved recipient failed the configured Outlook policy check.",
                "slots": [],
            })
        except Exception as exc:
            results.append({
                "email": _redact(email, account_email),
                "status": "error",
                "error": _redact(str(exc), account_email)[:120],
                "slots": [],
            })
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
    _core._assert_allowed("Calendar", account_email)
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
    try:
        if not appt.Recipients.ResolveAll():
            raise PermissionError("Cannot resolve all meeting recipients.")
    except PermissionError:
        raise
    except Exception as exc:
        raise PermissionError("Cannot resolve all meeting recipients.") from exc
    resolved_addresses = []
    for index in range(1, appt.Recipients.Count + 1):
        recipient = appt.Recipients.Item(index)
        address_entry = getattr(recipient, "AddressEntry", None)
        if address_entry is None:
            raise PermissionError("Cannot verify a resolved meeting recipient.")
        resolved_addresses.append(_core._resolve_smtp_from_entry(address_entry))
    _assert_domains_allowed(resolved_addresses, account_email=account_email)
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
