"""Account, folder, health and single-message read operations."""
from __future__ import annotations

import importlib.metadata
import logging

from mailmcp import _core, _folders
from mailmcp._core import get_config, get_effective_config, _assert_allowed, _redact
from mailmcp._formatters import _msg_header

logger = logging.getLogger(__name__)


def _server_version() -> str:
    try:
        return importlib.metadata.version("mcp-office")
    except importlib.metadata.PackageNotFoundError:
        return "unknown"


def health() -> dict:
    try:
        app = _core._get_outlook()
        return {
            "status": "ok",
            "outlook_name": getattr(app, "Name", "Microsoft Outlook"),
            "outlook_version": getattr(app, "Version", "unknown"),
            "send_enabled": get_config().enable_send,
            "server_version": _server_version(),
        }
    except Exception as exc:
        return {"status": "error", "message": str(exc)}


def list_accounts() -> list[dict]:
    mapi = _core._mapi()
    account_index = _folders._build_account_store_index(mapi)
    owners_by_store_id: dict[str, list[str]] = {}
    for smtp, store_ids in account_index["store_ids_by_smtp"].items():
        for store_id in store_ids:
            owners_by_store_id.setdefault(store_id, []).append(smtp)
    result = []
    for store in mapi.Stores:
        display_name = (store.DisplayName or "").strip()
        store_id = _folders._store_id(store)
        owners = sorted(set(owners_by_store_id.get(store_id, [])))
        if len(owners) > 1:
            raise PermissionError(
                "Cannot uniquely resolve the Outlook account owner for a shared DeliveryStore."
            )
        smtp_addr = owners[0] if owners else ""
        effective_cfg = _core.get_effective_config(smtp_addr if smtp_addr else None)
        result.append({
            "display_name": _redact(display_name, smtp_addr if smtp_addr else None),
            "exchange_store_type": getattr(store, "ExchangeStoreType", None),
            "config_profile": {
                "allowlist_folders": [
                    _redact(folder_name, smtp_addr if smtp_addr else None)
                    for folder_name in effective_cfg.allowlist_folders
                ],
                "enable_write": effective_cfg.enable_write,
                "enable_send": effective_cfg.enable_send,
                "enable_delete": effective_cfg.enable_delete,
                "max_items": effective_cfg.max_items,
                "is_override": bool(smtp_addr) and smtp_addr in _core._account_overrides,
            },
        })
    return result


def _list_folders_scoped(
    store_name: str | None = None,
    depth: int = 1,
    account_email: str | None = None,
    *,
    redact_names: bool,
) -> list[dict]:
    if account_email is None and _core._account_overrides:
        raise PermissionError(
            "account_email is required for folder discovery when per-account Outlook policy is configured."
        )
    cfg = get_effective_config(account_email)
    allowed = {name.lower() for name in cfg.allowlist_folders}
    mapi = _core._mapi()
    if account_email is not None:
        stores = [_folders._find_store_for_account(account_email, mapi)]
    else:
        try:
            default_inbox = mapi.GetDefaultFolder(_folders._OL_INBOX)
            default_store = _folders._resolve_store_for_object(default_inbox)
        except Exception as exc:
            raise PermissionError("Cannot resolve the default Outlook store for folder discovery.") from exc
        if default_store is None:
            raise PermissionError("Cannot resolve the default Outlook store for folder discovery.")
        stores = [default_store]

    selected_store = stores[0]
    selected_display = str(getattr(selected_store, "DisplayName", "") or "").strip()
    if store_name and selected_display.lower() != store_name.strip().lower():
        raise PermissionError("store_name is outside the selected Outlook account scope.")

    result: list[dict] = []

    def _walk(folder, current_depth: int) -> None:
        if current_depth <= 0:
            return
        try:
            children = folder.Folders
        except Exception:
            return
        try:
            for child in children:
                child_name = str(getattr(child, "Name", "") or "").strip()
                if child_name and ("*" in allowed or child_name.lower() in allowed):
                    result.append({
                        "name": _redact(child_name, account_email) if redact_names else child_name,
                        "unread_count": getattr(child, "UnReadItemCount", 0),
                        "item_count": getattr(child, "Items", None) and child.Items.Count or 0,
                    })
                _walk(child, current_depth - 1)
        except Exception as exc:
            logger.debug("Skipping inaccessible folder during scoped discovery: %s", exc)

    try:
        root = selected_store.GetRootFolder()
    except Exception as exc:
        raise PermissionError("Cannot open the selected Outlook account root folder.") from exc
    _walk(root, min(depth, 4))
    return result


def list_folders(
    store_name: str | None = None,
    depth: int = 1,
    account_email: str | None = None,
) -> list[dict]:
    return _list_folders_scoped(
        store_name=store_name,
        depth=depth,
        account_email=account_email,
        redact_names=True,
    )


def _list_folders_raw(
    store_name: str | None = None,
    depth: int = 1,
    account_email: str | None = None,
) -> list[dict]:
    """Internal-only folder discovery retaining raw names for COM resolution."""
    return _list_folders_scoped(
        store_name=store_name,
        depth=depth,
        account_email=account_email,
        redact_names=False,
    )


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
