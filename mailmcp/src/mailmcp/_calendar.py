"""Calendar listing, event fetch, update, and mailbox statistics."""
from __future__ import annotations

import logging
from datetime import datetime as _dt, timedelta as _td, timezone as _tz
from typing import Any

from mailmcp import _core
from mailmcp import _folders
from mailmcp._core import get_effective_config, _assert_allowed, _EMAIL_VALIDATE_RE
from mailmcp._folders import _OL_APPOINTMENT_CLASS
from mailmcp._formatters import _format_outlook_date
from mailmcp._item_guards import _assert_calendar_item_folder

logger = logging.getLogger(__name__)


def _local_naive(dt: _dt) -> _dt:
    return dt.astimezone().replace(tzinfo=None) if dt.tzinfo is not None else dt


def list_calendar_events(
    start: str | None = None,
    end: str | None = None,
    top: int | None = None,
    include_cancelled: bool = False,
    account_email: str | None = None,
) -> list[dict]:
    _assert_allowed("Calendar", account_email)
    cfg = get_effective_config(account_email)
    if top is not None and (
        isinstance(top, bool) or not isinstance(top, int) or top <= 0
    ):
        raise ValueError("top must be a positive integer or None.")
    limit = min(top, cfg.max_items) if top is not None else cfg.max_items
    if limit == 0:
        return []
    now = _dt.now()
    if start is None:
        start_dt = now
    else:
        try:
            start_dt = _dt.fromisoformat(start)
        except ValueError as exc:
            raise ValueError(f"Invalid 'start' datetime: {exc}") from exc
    if end is None:
        end_dt = start_dt + _td(days=7)
    else:
        try:
            end_dt = _dt.fromisoformat(end)
        except ValueError as exc:
            raise ValueError(f"Invalid 'end' datetime: {exc}") from exc
    start_dt = _local_naive(start_dt)
    end_dt = _local_naive(end_dt)
    if end_dt <= start_dt:
        raise ValueError("'end' must be after 'start'.")
    start_str = _format_outlook_date(start_dt)
    end_str = _format_outlook_date(end_dt)
    filter_str = f"[Start] <= '{end_str}' AND [End] >= '{start_str}'"
    cal = _folders._get_calendar_folder(account_email=account_email)
    items = cal.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    restricted = items.Restrict(filter_str)
    result: list[dict] = []
    for apt in restricted:
        if len(result) >= limit:
            break
        try:
            status = getattr(apt, "MeetingStatus", 0)
            if not include_cancelled and status in (5, 7):
                continue
            result.append(_folders._appointment_to_dict(apt, account_email=account_email))
        except Exception as exc:
            logger.debug("Skipping calendar item: %s", exc)
    return result


def get_calendar_event(entry_id: str, include_body: bool = False, account_email: str | None = None) -> dict:
    mapi = _core._mapi()
    try:
        item = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            raise ValueError(f"Calendar event not found (entry_id={entry_id!r}) — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Calendar event not found (entry_id={entry_id!r})") from exc
    _folders._assert_object_belongs_to_account(item, account_email, object_label="calendar item")
    item_class = getattr(item, "Class", None)
    if item_class != _OL_APPOINTMENT_CLASS:
        raise ValueError(f"Item entry_id={entry_id!r} is not an AppointmentItem (item.Class={item_class!r}, expected {_OL_APPOINTMENT_CLASS}).")
    _assert_calendar_item_folder(item, account_email)
    return _folders._appointment_to_dict(item, include_body=include_body, account_email=account_email)


def update_calendar_event(entry_id: str, subject: str | None = None, start_iso: str | None = None, end_iso: str | None = None, body: str | None = None, location: str | None = None, required_attendees: list[str] | None = None, optional_attendees: list[str] | None = None, all_day_event: bool | None = None, confirm: bool = False, account_email: str | None = None) -> dict:
    if not confirm:
        raise ValueError("confirm=True is required to update a calendar event.")
    _core._assert_write_enabled(account_email=account_email)
    for address in (required_attendees or []) + (optional_attendees or []):
        if not isinstance(address, str) or not _EMAIL_VALIDATE_RE.match(address):
            raise ValueError("Invalid attendee email address.")
    _core._assert_domains_allowed(
        (required_attendees or []) + (optional_attendees or []), account_email=account_email,
    )
    mapi = _core._mapi()
    try:
        item = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            raise ValueError(f"Calendar event not found (entry_id={entry_id!r}) — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Calendar event not found (entry_id={entry_id!r})") from exc
    _folders._assert_object_belongs_to_account(item, account_email, object_label="calendar item")
    if getattr(item, "Class", None) != _OL_APPOINTMENT_CLASS:
        raise ValueError(f"Item entry_id={entry_id!r} is not an AppointmentItem.")
    _assert_calendar_item_folder(item, account_email)
    if (start_iso is not None) != (end_iso is not None):
        try:
            raw_start = item.Start
            raw_end = item.End
        except Exception:
            raise PermissionError("Cannot read existing Start/End from calendar item — access denied.") from None
        try:
            existing_start = raw_start if isinstance(raw_start, _dt) else _dt.fromisoformat(str(raw_start).replace(" ", "T"))
            existing_end = raw_end if isinstance(raw_end, _dt) else _dt.fromisoformat(str(raw_end).replace(" ", "T"))
            existing_start = existing_start.replace(tzinfo=None)
            existing_end = existing_end.replace(tzinfo=None)
        except (ValueError, OverflowError) as exc:
            raise ValueError(f"Cannot parse existing Start/End: {exc}") from exc
        if start_iso is not None:
            new_start = _dt.fromisoformat(start_iso)
            if new_start.tzinfo is None:
                new_start = new_start.replace(tzinfo=_tz.utc)
            if new_start.astimezone().replace(tzinfo=None) >= existing_end:
                raise ValueError("'start_iso' must be before the existing event end. Provide 'end_iso' to reschedule both.")
        else:
            new_end = _dt.fromisoformat(end_iso)
            if new_end.tzinfo is None:
                new_end = new_end.replace(tzinfo=_tz.utc)
            if new_end.astimezone().replace(tzinfo=None) <= existing_start:
                raise ValueError("'end_iso' must be after the existing event start. Provide 'start_iso' to reschedule both.")
    if start_iso is not None and end_iso is not None:
        start_check = _dt.fromisoformat(start_iso)
        end_check = _dt.fromisoformat(end_iso)
        if start_check.tzinfo is None:
            start_check = start_check.replace(tzinfo=_tz.utc)
        if end_check.tzinfo is None:
            end_check = end_check.replace(tzinfo=_tz.utc)
        if end_check <= start_check:
            raise ValueError("'end_iso' must be after 'start_iso'.")
    if subject is not None:
        item.Subject = subject
    if start_iso is not None:
        start_dt = _dt.fromisoformat(start_iso)
        if start_dt.tzinfo is None:
            start_dt = start_dt.replace(tzinfo=_tz.utc)
        item.Start = start_dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
    if end_iso is not None:
        end_dt = _dt.fromisoformat(end_iso)
        if end_dt.tzinfo is None:
            end_dt = end_dt.replace(tzinfo=_tz.utc)
        item.End = end_dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
    if body is not None:
        item.Body = body
    if location is not None:
        item.Location = location
    if all_day_event is not None:
        item.AllDayEvent = all_day_event
    if required_attendees is not None or optional_attendees is not None:
        try:
            recipient_types_to_replace: set[int] = set()
            if required_attendees is not None:
                recipient_types_to_replace.add(1)
            if optional_attendees is not None:
                recipient_types_to_replace.add(2)
            for index in range(item.Recipients.Count, 0, -1):
                recipient = item.Recipients.Item(index)
                if getattr(recipient, "Type", None) in recipient_types_to_replace:
                    item.Recipients.Remove(index)
            for email in required_attendees or []:
                recip = item.Recipients.Add(email)
                recip.Type = 1
            for email in optional_attendees or []:
                recip = item.Recipients.Add(email)
                recip.Type = 2
            resolved = item.Recipients.ResolveAll()
        except Exception as exc:
            raise RuntimeError(f"Failed to update calendar attendees: {exc}") from exc
        if not resolved:
            raise RuntimeError("One or more attendee addresses could not be resolved.")
        resolved_addresses: list[str] = []
        for index in range(1, item.Recipients.Count + 1):
            recipient = item.Recipients.Item(index)
            address_entry = getattr(recipient, "AddressEntry", None)
            if address_entry is None:
                raise PermissionError("Cannot verify a resolved calendar attendee.")
            resolved_addresses.append(_core._resolve_smtp_from_entry(address_entry))
        try:
            _core._assert_domains_allowed(
                resolved_addresses, account_email=account_email
            )
        except PermissionError as exc:
            if get_effective_config(account_email).redact_mode != "none":
                raise PermissionError(
                    "One or more calendar attendees are blocked by domain policy."
                ) from exc
            raise
    item.Save()
    result = _folders._appointment_to_dict(item, include_body=True, account_email=account_email)
    result["status"] = "updated"
    return result


def _batch_resolve_allowlist_folders(names: list[str]) -> dict[str, Any]:
    from mailmcp._folders import _can_use_folder_cache, _normalized_store_key
    target = {n.lower() for n in names}
    found: dict[str, Any] = {}
    if _can_use_folder_cache():
        snapshot = _core._get_folder_cache_snapshot()
        if snapshot is not None:
            for key, folder in snapshot.items():
                parts = key.rsplit("/", 1)
                if len(parts) == 2 and parts[1] in target and parts[1] not in found:
                    found[parts[1]] = folder
    remaining = target - set(found)
    if not remaining:
        return found
    mapi = _core._mapi()
    for store in mapi.Stores:
        if not remaining:
            break
        try:
            store_key = _normalized_store_key(store)
            root = store.GetRootFolder()
            for folder in root.Folders:
                name_l = (getattr(folder, "Name", "") or "").lower()
                if name_l in remaining:
                    found[name_l] = folder
                    remaining.discard(name_l)
                    if _can_use_folder_cache():
                        _core._upsert_folder_cache_entry(f"{store_key}/{name_l}", folder)
                if remaining:
                    try:
                        for subfolder in folder.Folders:
                            sub_name = (getattr(subfolder, "Name", "") or "").lower()
                            if sub_name in remaining:
                                found[sub_name] = subfolder
                                remaining.discard(sub_name)
                                if _can_use_folder_cache():
                                    _core._upsert_folder_cache_entry(f"{store_key}/{sub_name}", subfolder)
                    except Exception as exc:
                        logger.debug("sub-folder access error: %s", exc)
        except Exception as exc:
            logger.warning("Skipping Outlook store during folder resolution: %s", exc)
    return found


def get_mailbox_stats(folder_path: str | None = None, account_email: str | None = None) -> dict:
    if account_email is None and _core._account_overrides:
        raise PermissionError(
            "account_email is required when per-account Outlook policy is configured."
        )
    if folder_path:
        _assert_allowed(folder_path, account_email)
    cfg = get_effective_config(account_email)
    if not folder_path and "*" in cfg.allowlist_folders:
        from mailmcp._message_fetch import list_folders

        discovered = list_folders(depth=4, account_email=account_email)
        folder_stats = [
            {
                "folder": str(row.get("name") or ""),
                "unread": int(row.get("unread_count") or 0),
                "total": int(row.get("item_count") or 0),
            }
            for row in discovered
            if str(row.get("name") or "").strip()
        ]
        return {
            "folders": folder_stats,
            "total_unread": sum(row["unread"] for row in folder_stats),
            "total_items": sum(row["total"] for row in folder_stats),
            "errors": [],
        }
    names = [folder_path] if folder_path else cfg.allowlist_folders
    folder_stats: list[dict] = []
    errors: list[dict] = []
    total_unread = 0
    total_items = 0
    bulk_map = _batch_resolve_allowlist_folders(names) if len(names) > 1 and account_email is None else None
    for name in names:
        try:
            if bulk_map is not None:
                folder = bulk_map.get(name.lower())
                if folder is None:
                    raise ValueError(f"Folder not found: '{name}'")
            elif account_email is not None:
                folder = _folders._folder_by_name_for_account(name, account_email=account_email)
            else:
                folder = _folders._folder_by_name(name)
            unread = getattr(folder, "UnReadItemCount", 0)
            total = folder.Items.Count
            folder_stats.append({"folder": name, "unread": unread, "total": total})
            total_unread += unread
            total_items += total
        except Exception as exc:
            logger.warning("get_mailbox_stats: skipping folder %s: %s", name, exc)
            errors.append({"folder": name, "error": str(exc)[:200]})
    return {"folders": folder_stats, "total_unread": total_unread, "total_items": total_items, "errors": errors}
