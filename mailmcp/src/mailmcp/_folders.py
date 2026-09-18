"""mailmcp._folders — Folder helpers, constants, attachment path validation."""
from __future__ import annotations

import logging
import threading
from pathlib import Path
from typing import Any

from mailmcp import _core
from mailmcp._core import get_effective_config, _redact
from mailmcp._formatters import _fmt_date

logger = logging.getLogger(__name__)

_OL_INBOX = 6
_OL_DELETED_ITEMS = 3
_OL_CALENDAR = 9
_OL_CONTACTS = 10
_OL_DRAFTS = 16
_OL_APPOINTMENT_CLASS = 26

_OL_FLAG_MARKED = 2
_OL_FLAG_COMPLETE = 1
_OL_NO_FLAG = 0

_FLAG_STATUS_MAP: dict[str, int] = {
    "flagged": _OL_FLAG_MARKED,
    "completed": _OL_FLAG_COMPLETE,
    "cleared": _OL_NO_FLAG,
}

_OL_TASKS = 13
_OL_JUNK_EMAIL = 23
_OL_TASK_CLASS = 48

_OL_TASK_STATUS: dict[int, str] = {
    0: "Not Started",
    1: "In Progress",
    2: "Complete",
    3: "Waiting",
    4: "Deferred",
}

_OL_PRIORITY_MAP: dict[str, int] = {"low": 0, "normal": 1, "high": 2}
_OL_PRIORITY_REVERSE_MAP: dict[int, str] = {v: k for k, v in _OL_PRIORITY_MAP.items()}

_OL_MEETING_RESPONSE: dict[str, int] = {
    "accept": 3,
    "tentative": 2,
    "decline": 4,
}

_ALLOWED_ATTACHMENT_ROOTS: tuple[Path, ...] = (
    Path.home().resolve(),
    Path("C:/Temp").resolve(),
)

_DANGEROUS_SAVE_EXTENSIONS: frozenset[str] = frozenset({
    ".exe", ".ps1", ".bat", ".vbs", ".cmd", ".scr",
    ".js", ".jar", ".hta", ".msi", ".dll", ".com", ".pif",
})

_ALLOWED_ATTACHMENT_EXTENSIONS: frozenset[str] = frozenset({
    ".pdf", ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt",
    ".txt", ".csv", ".rtf", ".msg", ".eml",
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff",
    ".zip", ".7z", ".ics",
})
_ALLOWED_SAVE_EXTENSIONS: frozenset[str] = _ALLOWED_ATTACHMENT_EXTENSIONS


def _validate_attachment_path(path: str) -> Path:
    p = Path(path).resolve()
    if not p.exists():
        raise ValueError(f"Attachment path does not exist: {path!r}")
    if not p.is_file():
        raise ValueError(f"Attachment path is not a regular file: {path!r}")
    if not any(p == root or p.is_relative_to(root) for root in _ALLOWED_ATTACHMENT_ROOTS):
        raise ValueError(
            f"Attachment path {path!r} is outside allowed roots. Move the file to a configured allowed location first."
        )
    if p.suffix.lower() not in _ALLOWED_ATTACHMENT_EXTENSIONS:
        raise ValueError(
            f"Attachment extension {p.suffix!r} is not in the allowlist. Allowed: {sorted(_ALLOWED_ATTACHMENT_EXTENSIONS)}"
        )
    return p


def _can_use_folder_cache() -> bool:
    return threading.current_thread() is threading.main_thread()


def _normalized_store_key(store: Any) -> str:
    display_name = getattr(store, "DisplayName", None)
    if not isinstance(display_name, str) or not display_name.strip():
        return "unknown"
    return display_name.strip().lower()


def _normalized_account_email(account_email: Any) -> str | None:
    if not isinstance(account_email, str):
        return None
    normalized = account_email.strip().lower()
    return normalized or None


def _store_id(store: Any) -> str:
    raw = getattr(store, "StoreID", None)
    if not isinstance(raw, str) or not raw.strip():
        raw = getattr(store, "EntryID", None)
    if not isinstance(raw, str):
        raw = ""
    return str(raw).strip().lower()


def _account_smtp_by_display(mapi: Any) -> dict[str, str]:
    smtp_by_display: dict[str, str] = {}
    try:
        for acct in mapi.Accounts:
            try:
                smtp_raw = getattr(acct, "SmtpAddress", None)
                display_raw = getattr(acct, "DisplayName", None)
                smtp = smtp_raw.strip().lower() if isinstance(smtp_raw, str) else ""
                display = display_raw.strip().lower() if isinstance(display_raw, str) else ""
                if smtp and display:
                    smtp_by_display[display] = smtp
            except Exception:
                pass
    except Exception:
        pass
    return smtp_by_display


def _build_account_store_index(mapi: Any) -> dict[str, Any]:
    smtp_by_display = _account_smtp_by_display(mapi)
    store_ids_by_smtp: dict[str, set[str]] = {}
    store_keys_by_smtp: dict[str, set[str]] = {}
    known_accounts: set[str] = set()
    try:
        for acct in mapi.Accounts:
            try:
                smtp = _normalized_account_email(getattr(acct, "SmtpAddress", None))
                display = (getattr(acct, "DisplayName", "") or "").strip().lower()
                if not smtp:
                    continue
                known_accounts.add(smtp)
                if display:
                    known_accounts.add(display)
                    store_keys_by_smtp.setdefault(smtp, set()).add(display)
                delivery_store = getattr(acct, "DeliveryStore", None)
                if delivery_store is not None:
                    delivery_store_id = _store_id(delivery_store)
                    if delivery_store_id:
                        store_ids_by_smtp.setdefault(smtp, set()).add(delivery_store_id)
                    delivery_store_key = _normalized_store_key(delivery_store)
                    if delivery_store_key:
                        store_keys_by_smtp.setdefault(smtp, set()).add(delivery_store_key)
            except Exception:
                pass
    except Exception:
        pass
    return {
        "smtp_by_display": smtp_by_display,
        "store_ids_by_smtp": store_ids_by_smtp,
        "store_keys_by_smtp": store_keys_by_smtp,
        "known_accounts": known_accounts,
    }


def _store_matches_account(store: Any, account_email: str, account_index: dict[str, Any]) -> bool:
    normalized_account = _normalized_account_email(account_email)
    if normalized_account is None:
        return True
    # Display labels are not account ownership evidence. A verifiable,
    # uniquely owned DeliveryStore ID is required for explicitly scoped
    # operations so a shared store cannot inherit the weaker of two policies.
    store_id = _store_id(store)
    if not store_id:
        return False
    owners = {
        smtp
        for smtp, store_ids in account_index["store_ids_by_smtp"].items()
        if store_id in store_ids
    }
    return owners == {normalized_account}


def _find_store_for_account(account_email: str, mapi: Any | None = None) -> Any:
    normalized_account = _normalized_account_email(account_email)
    if normalized_account is None:
        raise ValueError("account_email must be a non-empty email address")
    if mapi is None:
        mapi = _core._mapi()
    account_index = _build_account_store_index(mapi)
    for store in mapi.Stores:
        if _store_matches_account(store, normalized_account, account_index):
            return store
    if normalized_account not in account_index["known_accounts"]:
        raise ValueError(f"Account not found: '{account_email}'")
    raise PermissionError(f"Cannot resolve Outlook store for account '{account_email}'.")


def _resolve_store_for_object(obj: Any) -> Any | None:
    current = obj
    seen: set[int] = set()
    for _depth in range(10):
        if current is None:
            return None
        identity = id(current)
        if identity in seen:
            return None
        seen.add(identity)
        try:
            display_name = getattr(current, "DisplayName", None)
            if callable(getattr(current, "GetRootFolder", None)) and (
                (isinstance(display_name, str) and display_name.strip()) or _store_id(current)
            ):
                return current
        except Exception:
            pass
        try:
            store = getattr(current, "Store", None)
        except Exception:
            store = None
        if store is not None and (_store_id(store) or _normalized_store_key(store) != "unknown"):
            return store
        try:
            current = getattr(current, "Parent", None)
        except Exception:
            return None
    return None


def _assert_object_belongs_to_account(obj: Any, account_email: str | None, object_label: str = "item") -> Any | None:
    normalized_account = _normalized_account_email(account_email)
    if normalized_account is None:
        return None
    mapi = _core._mapi()
    store = _resolve_store_for_object(obj)
    if store is None:
        raise PermissionError(f"Cannot verify that this {object_label} belongs to account '{account_email}'.")
    account_index = _build_account_store_index(mapi)
    if normalized_account not in account_index["known_accounts"]:
        raise ValueError(f"Account not found: '{account_email}'")
    if not _store_matches_account(store, normalized_account, account_index):
        raise PermissionError(f"The requested {object_label} does not belong to account '{account_email}'.")
    return store


def _find_folder_in_store(store: Any, normalized_folder_name: str) -> Any | None:
    root = store.GetRootFolder()
    for folder in root.Folders:
        if (getattr(folder, "Name", "") or "").lower() == normalized_folder_name:
            return folder
        try:
            for subfolder in folder.Folders:
                if (getattr(subfolder, "Name", "") or "").lower() == normalized_folder_name:
                    return subfolder
        except Exception:
            pass
    return None


def _folder_by_name_for_account(folder_name: str, account_email: str | None = None) -> Any:
    normalized_account = _normalized_account_email(account_email)
    if normalized_account is None:
        return _folder_by_name(folder_name)
    normalized_folder_name = folder_name.lower()
    mapi = _core._mapi()
    account_index = _build_account_store_index(mapi)
    cached_snapshot = _core._get_folder_cache_snapshot() if _can_use_folder_cache() else None
    matched_store = False
    for store in mapi.Stores:
        if not _store_matches_account(store, normalized_account, account_index):
            continue
        matched_store = True
        store_key = f"id:{_store_id(store)}"
        if cached_snapshot is not None:
            cached_folder = cached_snapshot.get(f"{store_key}/{normalized_folder_name}")
            if cached_folder is not None:
                return cached_folder
        try:
            folder = _find_folder_in_store(store, normalized_folder_name)
        except Exception:
            continue
        if folder is not None:
            if _can_use_folder_cache():
                _core._upsert_folder_cache_entry(f"{store_key}/{normalized_folder_name}", folder)
            return folder
    if not matched_store:
        raise ValueError(f"Account not found: '{account_email}'")
    raise ValueError(f"Folder not found for account '{account_email}': '{folder_name}'")


def _get_default_folder_for_account(default_folder_id: int, account_email: str | None = None, fallback_name: str | None = None) -> Any:
    normalized_account = _normalized_account_email(account_email)
    if normalized_account is None:
        return _core._mapi().GetDefaultFolder(default_folder_id)
    store = _find_store_for_account(normalized_account)
    try:
        folder = store.GetDefaultFolder(default_folder_id)
    except Exception:
        folder = None
    if folder is not None:
        return folder
    if fallback_name:
        folder = _find_folder_in_store(store, fallback_name.lower())
        if folder is not None:
            return folder
    if fallback_name:
        raise ValueError(f"Folder not found for account '{account_email}': '{fallback_name}'")
    raise ValueError(f"Default folder {default_folder_id} not found for account '{account_email}'.")


def _folder_by_name(folder_name: str) -> Any:
    normalized = folder_name.lower()
    if normalized == "inbox":
        if _can_use_folder_cache():
            snapshot = _core._get_folder_cache_snapshot()
            if snapshot is not None:
                cached_primary = snapshot.get("primary/inbox")
                if cached_primary is not None:
                    return cached_primary
        mapi = _core._mapi()
        try:
            inbox = mapi.GetDefaultFolder(_OL_INBOX)
        except Exception as exc:
            logger.warning("GetDefaultFolder(%d) failed, falling back to store walk: %s", _OL_INBOX, exc)
        else:
            if _can_use_folder_cache():
                _core._upsert_folder_cache_entry("primary/inbox", inbox)
            return inbox
    else:
        mapi = None
    if _can_use_folder_cache():
        snapshot = _core._get_folder_cache_snapshot()
        if snapshot is not None:
            target = normalized
            for key, cached_f in snapshot.items():
                if key.endswith(f"/{target}"):
                    return cached_f
    if mapi is None:
        mapi = _core._mapi()
    for store in mapi.Stores:
        try:
            store_key = _normalized_store_key(store)
            folder = _find_folder_in_store(store, normalized)
            if folder is not None:
                if _can_use_folder_cache():
                    _core._upsert_folder_cache_entry(f"{store_key}/{normalized}", folder)
                return folder
        except Exception:
            continue
    if normalized == "inbox":
        try:
            inbox = mapi.GetDefaultFolder(_OL_INBOX)
        except Exception as exc:
            raise ValueError(f"Inbox not found in any store (GetDefaultFolder also failed: {exc})") from exc
        if _can_use_folder_cache():
            _core._upsert_folder_cache_entry("primary/inbox", inbox)
        return inbox
    raise ValueError(f"Folder not found: '{folder_name}'")


def _get_calendar_folder(account_email: str | None = None) -> Any:
    return _get_default_folder_for_account(_OL_CALENDAR, account_email=account_email, fallback_name="Calendar")


def _get_deleted_items_folder(account_email: str | None = None) -> Any:
    return _get_default_folder_for_account(_OL_DELETED_ITEMS, account_email=account_email, fallback_name="Deleted Items")


def _get_drafts_folder(account_email: str | None = None) -> Any:
    return _get_default_folder_for_account(_OL_DRAFTS, account_email=account_email, fallback_name="Drafts")


def _get_tasks_folder(account_email: str | None = None) -> Any:
    return _get_default_folder_for_account(_OL_TASKS, account_email=account_email, fallback_name="Tasks")


def _get_junk_folder(account_email: str | None = None) -> Any:
    return _get_default_folder_for_account(_OL_JUNK_EMAIL, account_email=account_email, fallback_name="Junk Email")


def _appointment_to_dict(apt: Any, include_body: bool = False, account_email: str | None = None) -> dict:
    cfg = get_effective_config(account_email)
    result: dict = {
        "entry_id": apt.EntryID,
        "subject": _redact(getattr(apt, "Subject", "") or "", account_email=account_email),
        "start": _fmt_date(getattr(apt, "Start", None)),
        "end": _fmt_date(getattr(apt, "End", None)),
        "location": _redact(getattr(apt, "Location", "") or "", account_email=account_email),
        "organizer": _redact(getattr(apt, "Organizer", "") or "", account_email=account_email),
        "all_day_event": bool(getattr(apt, "AllDayEvent", False)),
        "meeting_status": getattr(apt, "MeetingStatus", 0),
        "is_recurring": bool(getattr(apt, "IsRecurring", False)),
        "required_attendees": _redact(getattr(apt, "RequiredAttendees", "") or "", account_email=account_email),
        "optional_attendees": _redact(getattr(apt, "OptionalAttendees", "") or "", account_email=account_email),
        "categories": _redact(getattr(apt, "Categories", "") or "", account_email=account_email),
        "sensitivity": getattr(apt, "Sensitivity", 0),
        "importance": getattr(apt, "Importance", 1),
    }
    if include_body:
        body = getattr(apt, "Body", "") or ""
        truncated = len(body) > cfg.max_body_chars
        result["body"] = _redact(body[: cfg.max_body_chars], account_email=account_email)
        result["body_truncated"] = truncated
    return result


def create_folder(parent_folder: str, name: str, account_email: str | None = None) -> dict:
    _core._assert_write_enabled(account_email)
    _core._assert_allowed(parent_folder, account_email)
    name = name.strip()
    if not name:
        raise ValueError("Folder name must not be empty.")
    for ch in ("\\", "/", ":"):
        if ch in name:
            raise ValueError(f"Folder name must not contain '{ch}'. Received: {name!r}")
    if "\x00" in name:
        raise ValueError("Folder name must not contain null bytes")
    if len(name) > 255:
        raise ValueError("Folder name must not exceed 255 characters")
    if name.startswith("."):
        raise ValueError("Folder name must not start with a dot")
    parent_com = _folder_by_name_for_account(parent_folder, account_email)
    try:
        new_folder = parent_com.Folders.Add(name)
    except Exception as exc:
        exc_str = str(exc).lower()
        if "already exists" in exc_str or "operation failed" in exc_str:
            raise ValueError(f"A folder named '{name}' already exists under '{parent_folder}'.") from exc
        raise RuntimeError(f"Cannot create folder '{name}' under '{parent_folder}': {exc}") from exc
    _core._invalidate_folder_cache()
    return {
        "ok": True,
        "name": new_folder.Name,
        "parent": parent_folder,
        "entry_id": getattr(new_folder, "EntryID", None),
    }
