"""mailmcp._categories — Category management COM helpers (Phase 2.0)."""
from __future__ import annotations

import logging

from mailmcp import _folders
from mailmcp._core import _mapi, get_effective_config, _assert_allowed, _assert_write_enabled
from mailmcp._folders import _folder_by_name, _folder_by_name_for_account
from mailmcp._formatters import _msg_header, _sql_escape

logger = logging.getLogger(__name__)


def list_categories() -> dict:
    """List all Outlook master categories defined in the user's profile."""
    try:
        mapi = _mapi()
        result = []
        for cat in mapi.Categories:
            result.append({
                "name":         str(cat.Name),
                "color":        int(getattr(cat, "Color", 0)),
                "shortcut_key": int(getattr(cat, "ShortcutKey", 0)),
            })
        return {"categories": result, "count": len(result)}
    except Exception as exc:
        logger.error("list_categories failed: %s", exc)
        raise RuntimeError(f"Cannot read categories: {exc}") from exc


def set_message_category(
    entry_id: str,
    categories: list[str],
    confirm: bool = False,
    account_email: str | None = None,
) -> dict:
    """Assign categories to any Outlook item (mail, task, appointment, etc.)."""
    if not confirm:
        raise ValueError(
            "confirm=True is required to set message categories. "
            "Set confirm=True to proceed."
        )
    _assert_write_enabled(account_email=account_email)

    try:
        mapi = _mapi()
        item = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        hresult = getattr(exc, 'hresult', None)
        if hresult is not None:
            raise ValueError(f"Item not found (entry_id={entry_id!r}) — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Item not found (entry_id={entry_id!r})") from exc

    _folders._assert_object_belongs_to_account(item, account_email, object_label="item")

    try:
        _parent_name = item.Parent.Name
    except Exception as _exc:
        _hresult = getattr(_exc, "hresult", None)
        if _hresult is not None:
            raise ValueError(
                f"Cannot determine source folder for entry_id={entry_id!r} — COM HRESULT {_hresult:#010x}"
            ) from _exc
        raise ValueError(
            f"Cannot determine source folder for entry_id={entry_id!r}"
        ) from _exc
    _assert_allowed(_parent_name, account_email)  # NB-003

    try:
        item.Categories = ", ".join(categories)
        item.Save()
        saved = str(getattr(item, "Categories", "") or "")
        saved_list = [c.strip() for c in saved.split(",") if c.strip()]
        return {"ok": True, "entry_id": entry_id, "categories": saved_list}
    except (ValueError, TypeError):
        raise
    except Exception as exc:
        logger.error("set_message_category failed: %s", exc)
        raise RuntimeError(f"Cannot set categories: {exc}") from exc


def get_messages_by_category(
    category: str,
    folder_name: str | None = None,
    top: int = 20,
    account_email: str | None = None,
) -> dict:
    """Retrieve messages tagged with a specific category from a folder."""
    cfg = get_effective_config(account_email)
    top = min(top, cfg.max_items)

    _folder_to_check = folder_name if folder_name is not None else "Inbox"
    _assert_allowed(_folder_to_check, account_email)
    mapi = _mapi()
    if folder_name is not None:
        if account_email is not None:
            folder = _folder_by_name_for_account(folder_name, account_email=account_email)
        else:
            folder = _folder_by_name(folder_name)
    else:
        if account_email is not None:
            folder = _folders._get_default_folder_for_account(6, account_email=account_email, fallback_name="Inbox")
        else:
            folder = mapi.GetDefaultFolder(6)  # olFolderInbox

    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception as exc:
        logger.warning("get_messages_by_category: Sort failed (%s)", exc)

    safe_cat = _sql_escape(category)
    filter_str = f"[Categories] LIKE '%{safe_cat}%'"
    try:
        restricted = items.Restrict(filter_str)
        restricted.Sort("[ReceivedTime]", True)
        items = restricted
    except Exception as exc:
        raise RuntimeError(
            "Outlook could not apply the requested category filter; no unfiltered messages were returned."
        ) from exc

    result = []
    count = 0
    for msg in items:
        if count >= top:
            break
        try:
            result.append(_msg_header(msg, account_email=account_email))
            count += 1
        except Exception as exc:
            logger.warning("get_messages_by_category: skipping item (%s)", exc)
            continue

    return {"messages": result, "count": len(result)}
