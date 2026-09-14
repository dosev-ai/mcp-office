"""Account, folder, health and single-message read operations."""
from __future__ import annotations

import importlib.metadata
import logging

from mailmcp import _core, _folders
from mailmcp._core import get_config, get_effective_config, _assert_allowed, _redact
from mailmcp._formatters import _msg_header

logger = logging.getLogger(__name__)


def health() -> dict:
    try:
        app = _core._get_outlook()
        return {
            "status": "ok",
            "outlook_name": getattr(app, "Name", "Microsoft Outlook"),
            "outlook_version": getattr(app, "Version", "unknown"),
            "send_enabled": get_config().enable_send,
            "server_version": importlib.metadata.version("mcp-office"),
        }
    except Exception as exc:
        return {"status": "error", "message": str(exc)}


def list_accounts() -> list[dict]:
    mapi = _core._mapi()
    smtp_by_display: dict[str, str] = {}
    try:
        for acct in mapi.Accounts:
            try:
                smtp = (getattr(acct, "SmtpAddress", "") or "").strip().lower()
                display = (getattr(acct, "DisplayName", "") or "").strip()
                if smtp and display:
                    smtp_by_display[display.lower()] = smtp
            except Exception:
                pass
    except Exception:
        pass
    result = []
    for store in mapi.Stores:
        display_name = (store.DisplayName or "").strip()
        smtp_addr = smtp_by_display.get(display_name.lower(), "")
        effective_cfg = _core.get_effective_config(smtp_addr if smtp_addr else None)
        result.append({
            "display_name": _redact(display_name),
            "exchange_store_type": getattr(store, "ExchangeStoreType", None),
            "config_profile": {
                "allowlist_folders": effective_cfg.allowlist_folders,
                "enable_write": effective_cfg.enable_write,
                "enable_send": effective_cfg.enable_send,
                "enable_delete": effective_cfg.enable_delete,
                "max_items": effective_cfg.max_items,
                "is_override": bool(smtp_addr) and smtp_addr in _core._account_overrides,
            },
        })
    return result


def list_folders(store_name: str | None = None, depth: int = 1) -> list[dict]:
    cached_snapshot = _core._get_folder_cache_snapshot() if (store_name is None and depth == 2) else None
    if cached_snapshot is not None and _core._folder_cache_is_full:
        result = []
        for _key, folder in cached_snapshot.items():
            try:
                result.append({
                    "name": folder.Name,
                    "unread_count": getattr(folder, "UnReadItemCount", 0),
                    "item_count": getattr(folder, "Items", None) and folder.Items.Count or 0,
                })
            except Exception as exc:
                logger.debug("Skipping stale cached folder object: %s", exc)
        return result
    mapi = _core._mapi()
    result = []
    new_cache: dict = {}

    def _walk(folder, current_depth: int, store_key: str):
        if current_depth <= 0:
            return
        try:
            for child in folder.Folders:
                result.append({
                    "name": child.Name,
                    "unread_count": getattr(child, "UnReadItemCount", 0),
                    "item_count": getattr(child, "Items", None) and child.Items.Count or 0,
                })
                if child.Name:
                    new_cache[f"{store_key}/{child.Name.lower()}"] = child
                _walk(child, current_depth - 1, store_key)
        except Exception:
            pass

    for store in mapi.Stores:
        if store_name and store.DisplayName.lower() != store_name.lower():
            continue
        try:
            store_key = getattr(store, "DisplayName", "Unknown").lower()
            _walk(store.GetRootFolder(), min(depth, 4), store_key)
        except Exception:
            pass
    if store_name is None and depth == 2 and new_cache:
        _core._set_folder_cache(new_cache)
    return result


def get_message(entry_id: str, include_body: bool = False, account_email: str | None = None) -> dict:
    mapi = _core._mapi()
    try:
        msg = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            raise ValueError(f"Message not found (entry_id={entry_id!r}) — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    _folders._assert_object_belongs_to_account(msg, account_email, object_label="message")
    try:
        parent_name = msg.Parent.Name
    except Exception as exc:
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            raise ValueError(f"Cannot determine source folder for entry_id={entry_id!r} — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Cannot determine source folder for entry_id={entry_id!r}") from exc
    _assert_allowed(parent_name, account_email)
    result = _msg_header(msg, account_email=account_email)
    if include_body:
        body = getattr(msg, "Body", "") or ""
        cfg = get_effective_config(account_email)
        result["body"] = _redact(body[: cfg.max_body_chars], account_email=account_email)
        result["body_truncated"] = len(body) > cfg.max_body_chars
    return result
