"""Outlook task and meeting-request helpers."""
from __future__ import annotations

import logging
import re
from datetime import datetime as _dt, timedelta as _td

from mailmcp._core import _mapi, _get_outlook, get_effective_config, _assert_allowed, _assert_write_enabled, _redact, _assert_domains_allowed, _resolve_smtp_from_entry
from mailmcp._formatters import _resolve_sender_email
from mailmcp._folders import _folder_by_name, _folder_by_name_for_account, _get_tasks_folder, _OL_TASKS, _OL_TASK_CLASS, _OL_TASK_STATUS, _OL_PRIORITY_MAP, _OL_PRIORITY_REVERSE_MAP, _OL_MEETING_RESPONSE, _assert_object_belongs_to_account

logger = logging.getLogger(__name__)
_OL_MAIL_CLASS = 43


def list_tasks(include_completed: bool = False, due_before: str | None = None, top: int = 20, folder_path: str | None = None, account_email: str | None = None) -> dict:
    folder_to_check = folder_path or "Tasks"
    _assert_allowed(folder_to_check, account_email)
    cfg = get_effective_config(account_email)
    top = min(top, cfg.max_items)
    due_dt_obj = None
    if due_before is not None:
        try:
            due_dt_obj = _dt.fromisoformat(due_before)
        except ValueError as exc:
            raise ValueError(f"Invalid 'due_before' format: {due_before!r}. Use ISO 8601 (YYYY-MM-DD).") from exc
    mapi = _mapi()
    if folder_path:
        folder = _folder_by_name_for_account(folder_path, account_email=account_email) if account_email is not None else _folder_by_name(folder_path)
    else:
        if account_email is not None:
            from mailmcp import _folders
            folder = _folders._get_tasks_folder(account_email=account_email)
        else:
            folder = mapi.GetDefaultFolder(_OL_TASKS)
    items = folder.Items
    items.Sort("[DueDate]")
    filters = []
    if not include_completed:
        filters.append("[Complete] = False")
    if due_dt_obj is not None:
        filters.append(f"[DueDate] <= '{due_dt_obj.strftime('%m/%d/%Y %I:%M %p')}'")
    if filters:
        try:
            items = items.Restrict(" AND ".join(f"({f})" for f in filters))
        except Exception as exc:
            raise RuntimeError(
                "Outlook rejected the requested task filter; no unfiltered tasks were returned."
            ) from exc
    result = []
    for task in items:
        if len(result) >= top:
            break
        try:
            status_int = int(getattr(task, "Status", 0))
            importance = int(getattr(task, "Importance", 1))
            due_val = getattr(task, "DueDate", None)
            due_iso = str(due_val)[:10] if due_val and str(due_val) != "4501-01-01 00:00:00" else None
            body_text = str(getattr(task, "Body", "") or "")
            result.append({"entry_id": task.EntryID, "subject": _redact(str(task.Subject or ""), account_email), "status": status_int, "status_label": _OL_TASK_STATUS.get(status_int, "Unknown"), "priority": _OL_PRIORITY_REVERSE_MAP.get(importance, "normal"), "due_date": due_iso, "complete": bool(getattr(task, "Complete", False)), "categories": str(getattr(task, "Categories", "") or ""), "body_preview": _redact(body_text[:200], account_email)})
        except Exception as exc:
            logger.warning("list_tasks: skipping item (%s)", exc)
    return {"tasks": result, "count": len(result)}


def create_task(subject: str, body: str | None = None, due_date: str | None = None, priority: str = "normal", confirm: bool = False, account_email: str | None = None) -> dict:
    _assert_write_enabled(account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True is required to create a task. Set confirm=True to proceed.")
    _assert_allowed("Tasks", account_email)
    if not subject or not subject.strip():
        raise ValueError("subject is required and cannot be blank")
    priority_lower = priority.lower()
    if priority_lower not in _OL_PRIORITY_MAP:
        raise ValueError(f"Invalid priority: {priority!r}. Must be one of: low, normal, high.")
    due_dt = None
    if due_date is not None:
        try:
            due_dt = _dt.fromisoformat(due_date)
        except ValueError as exc:
            raise ValueError(f"Invalid 'due_date' format: {due_date!r}. Use ISO 8601 (YYYY-MM-DD).") from exc
    try:
        if account_email is not None:
            task = _get_tasks_folder(account_email=account_email).Items.Add()
        else:
            task = _get_outlook().CreateItem(3)
        task.Subject = subject
        if body:
            task.Body = body
        if due_dt is not None:
            task.DueDate = due_dt
        task.Importance = _OL_PRIORITY_MAP[priority_lower]
        task.Status = 0
        task.Save()
        raw_due = getattr(task, "DueDate", None)
        due_iso = str(raw_due)[:10] if raw_due and str(raw_due) != "4501-01-01 00:00:00" else None
        return {"status": "created", "entry_id": task.EntryID, "subject": _redact(str(task.Subject or ""), account_email), "due_date": due_iso, "priority": priority_lower}
    except (ValueError, TypeError):
        raise
    except Exception as exc:
        raise RuntimeError(f"Cannot create task: {exc}") from exc


def complete_task(entry_id: str, confirm: bool = False, account_email: str | None = None) -> dict:
    _assert_write_enabled(account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True is required to complete a task. Set confirm=True to proceed.")
    try:
        item = _mapi().GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Task not found (entry_id={entry_id!r})") from exc
    from mailmcp import _folders
    _folders._assert_object_belongs_to_account(item, account_email, object_label="task")
    if getattr(item, "Class", None) != _OL_TASK_CLASS:
        raise ValueError(f"Item entry_id={entry_id!r} is not a TaskItem.")
    parent_name = item.Parent.Name
    if account_email is not None:
        default_tasks = _folders._get_tasks_folder(account_email=account_email)
        parent_entry_id = getattr(item.Parent, "EntryID", None)
        default_entry_id = getattr(default_tasks, "EntryID", None)
        if default_entry_id is not None and parent_entry_id == default_entry_id:
            _assert_allowed("Tasks", account_email)
        else:
            _assert_allowed(parent_name, account_email)
    elif parent_name in ("Tasks", "To-Do"):
        _assert_allowed("Tasks", account_email)
    else:
        _assert_allowed(parent_name, account_email)
    item.Status = 2
    item.Complete = True
    item.PercentComplete = 100
    item.DateCompleted = _dt.now()
    item.Save()
    return {"ok": True, "entry_id": entry_id, "status": "completed"}


def list_meeting_requests(top: int = 20, folder_name: str | None = None, folder_path: str | None = None, account_email: str | None = None) -> dict:
    cfg = get_effective_config(account_email)
    top = min(top, cfg.max_items)
    effective_folder = folder_path or folder_name
    folder_to_check = effective_folder if effective_folder is not None else "Inbox"
    _assert_allowed(folder_to_check, account_email)
    mapi = _mapi()
    if effective_folder is not None:
        folder = _folder_by_name_for_account(effective_folder, account_email=account_email) if account_email is not None else _folder_by_name(effective_folder)
    else:
        if account_email is not None:
            from mailmcp import _folders
            folder = _folders._get_default_folder_for_account(6, account_email=account_email, fallback_name="Inbox")
        else:
            folder = mapi.GetDefaultFolder(6)
    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    try:
        items = items.Restrict("[MessageClass] = 'IPM.Schedule.Meeting.Request'")
    except Exception as exc:
        raise RuntimeError(
            "Outlook rejected the meeting-request filter; no unfiltered messages were returned."
        ) from exc
    result = []
    for item in items:
        if len(result) >= top:
            break
        try:
            result.append({"entry_id": item.EntryID, "subject": _redact(str(getattr(item, "Subject", "") or ""), account_email=account_email), "sender_name": _redact(str(getattr(item, "SenderName", "") or ""), account_email=account_email), "sender_email": _redact(_resolve_sender_email(item), account_email=account_email), "received_time": str(getattr(item, "ReceivedTime", "") or ""), "start": str(getattr(item, "Start", None)) if getattr(item, "Start", None) else None, "end": str(getattr(item, "End", None)) if getattr(item, "End", None) else None, "location": _redact(str(getattr(item, "Location", "") or ""), account_email=account_email), "response_requested": bool(getattr(item, "ResponseRequested", False)), "categories": str(getattr(item, "Categories", "") or "")})
        except Exception as exc:
            logger.warning("list_meeting_requests: skipping item (%s)", exc)
    return {"requests": result, "count": len(result)}


def respond_to_meeting(entry_id: str, response: str, confirm: bool = False, account_email: str | None = None) -> dict:
    _assert_write_enabled(account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True is required to respond to a meeting request. Set confirm=True to proceed.")
    response_lower = response.lower()
    if response_lower not in _OL_MEETING_RESPONSE:
        raise ValueError(f"Invalid response: {response!r}. Must be one of: accept, decline, tentative.")
    try:
        item = _mapi().GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Meeting request not found (entry_id={entry_id!r})") from exc
    _assert_object_belongs_to_account(item, account_email, object_label="meeting request")
    if getattr(item, "MessageClass", "") != "IPM.Schedule.Meeting.Request":
        raise ValueError("Item is not a meeting request.")
    _assert_allowed(item.Parent.Name, account_email)
    organizer_smtp = _resolve_smtp_from_entry(item.Sender)
    _assert_domains_allowed([organizer_smtp], account_email=account_email)
    try:
        response_item = item.Respond(_OL_MEETING_RESPONSE[response_lower], True, False)
        cfg = get_effective_config(account_email)
        sent = False
        if cfg.enable_send and response_item is not None:
            response_item.Send()
            sent = True
        elif response_item is not None:
            response_item.Save()
    except Exception as exc:
        raise RuntimeError(f"Cannot respond to meeting: {exc}") from exc
    return {"ok": True, "entry_id": entry_id, "response": response_lower, "sent": sent}


_ACTION_PATTERNS = [
    re.compile(r'^\s*-\s*\[\s*\]\s*(.+)', re.MULTILINE),
    re.compile(r'^\s*(?:TODO|Action|ACTION ITEM|TASK)\s*:\s*(.+)', re.MULTILINE | re.IGNORECASE),
    re.compile(r'^\s*[→•]\s*(.+)', re.MULTILINE),
    re.compile(r'^\s*\d+\.\s+(\S+(?:[ \t]+\S+){2,}[^\n]*)$', re.MULTILINE),
]
_DATE_ISO_RE = re.compile(r'(\d{4}-\d{2}-\d{2})')
_DATE_WEEKDAY_RE = re.compile(r'\bby\s+(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b', re.IGNORECASE)
_DATE_EOW_RE = re.compile(r'\b(?:due|by)\s+(?:end\s+of\s+)?(?:this\s+)?week\b', re.IGNORECASE)
_DATE_EOD_RE = re.compile(r'\b(?:due|by)\s+(?:end\s+of\s+)?(?:this\s+)?day\b|\b(?:by\s+)?EOD\b', re.IGNORECASE)
_PRIORITY_HIGH_RE = re.compile(r'\b(?:urgent|ASAP|high\s+priority|critical|immediately)\b', re.IGNORECASE)
_PRIORITY_LOW_RE = re.compile(r'\b(?:low\s+priority|whenever|no\s+rush)\b', re.IGNORECASE)
_WEEKDAY_NAMES = {'monday': 0, 'tuesday': 1, 'wednesday': 2, 'thursday': 3, 'friday': 4, 'saturday': 5, 'sunday': 6}


def _next_weekday(weekday_name: str) -> str:
    target = _WEEKDAY_NAMES[weekday_name.lower()]
    today = _dt.now()
    days_ahead = (target - today.weekday()) % 7 or 7
    return (today + _td(days=days_ahead)).strftime('%Y-%m-%d')


def _end_of_week() -> str:
    today = _dt.now()
    return (today + _td(days=(4 - today.weekday()) % 7)).strftime('%Y-%m-%d')


def _detect_due_date(text: str) -> str | None:
    match = _DATE_ISO_RE.search(text)
    if match:
        return match.group(1)
    if _DATE_EOD_RE.search(text):
        return _dt.now().strftime('%Y-%m-%d')
    if _DATE_EOW_RE.search(text):
        return _end_of_week()
    match = _DATE_WEEKDAY_RE.search(text)
    return _next_weekday(match.group(1)) if match else None


def _detect_priority(text: str) -> str:
    if _PRIORITY_HIGH_RE.search(text):
        return 'high'
    if _PRIORITY_LOW_RE.search(text):
        return 'low'
    return 'normal'


def _extract_task_candidates(body: str) -> list[dict]:
    seen: set[str] = set()
    candidates = []
    for pattern in _ACTION_PATTERNS:
        for match in pattern.finditer(body):
            title = match.group(1).strip()
            if title and title.lower() not in seen:
                seen.add(title.lower())
                candidates.append({'title': title, 'due_date': _detect_due_date(title), 'priority': _detect_priority(title)})
    return candidates


def extract_tasks_from_message(entry_id: str, auto_create: bool = False, confirm: bool = False, account_email: str | None = None) -> dict:
    if auto_create:
        _assert_write_enabled(account_email=account_email)
    if auto_create and not confirm:
        raise ValueError("confirm=True is required when auto_create=True. Set confirm=True to proceed.")
    try:
        item = _mapi().GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    _assert_object_belongs_to_account(item, account_email, object_label="message")
    _assert_allowed(getattr(item.Parent, "Name", ""), account_email)
    if getattr(item, 'Class', None) != _OL_MAIL_CLASS:
        raise ValueError(f"Item entry_id={entry_id!r} is not a MailItem.")
    cfg = get_effective_config(account_email)
    body = str(getattr(item, 'Body', '') or '')[:cfg.max_body_chars]
    subject = str(getattr(item, 'Subject', '') or '')
    candidates = _extract_task_candidates(body)
    created = []
    if auto_create and candidates:
        candidates = candidates[:cfg.max_items]
        for candidate in candidates:
            created.append(create_task(subject=candidate['title'], due_date=candidate['due_date'], priority=candidate['priority'], confirm=True, account_email=account_email))
    visible_candidates = [{**row, 'title': _redact(row['title'], account_email)} for row in candidates]
    visible_created = [
        {**row, 'subject': _redact(str(row.get('subject') or ''), account_email)}
        for row in created
    ]
    return {
        'candidates': visible_candidates, 'created': visible_created,
        'source_entry_id': entry_id, 'source_subject': _redact(subject, account_email),
    }
