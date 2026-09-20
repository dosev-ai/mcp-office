"""Attachment save operations."""
from __future__ import annotations

from pathlib import Path
import threading

from mailmcp import _core, _folders
from mailmcp._core import get_effective_config, _assert_allowed

_SAVE_LOCK = threading.Lock()


def _effective_account_for_item(mapi, msg, account_email: str | None) -> str | None:
    if account_email is not None:
        _folders._assert_object_belongs_to_account(msg, account_email, object_label="message")
        return account_email.strip().lower()
    if not _core._account_overrides:
        return None
    store = _folders._resolve_store_for_object(msg)
    if store is None:
        raise PermissionError("Cannot determine the source message's Outlook account.")
    store_id = _folders._store_id(store)
    if not store_id:
        raise PermissionError("Cannot determine the source message's Outlook store identity.")
    index = _folders._build_account_store_index(mapi)
    owners = [
        smtp for smtp, store_ids in index["store_ids_by_smtp"].items()
        if store_id in store_ids
    ]
    if len(owners) != 1:
        raise PermissionError("Cannot resolve a unique Outlook account for the source message.")
    return owners[0]


def save_attachments(
    entry_id: str,
    save_dir: str,
    account_email: str | None = None,
) -> list[dict]:
    dest = Path(save_dir).resolve()
    if not any(dest == root or dest.is_relative_to(root) for root in _folders._ALLOWED_ATTACHMENT_ROOTS):
        raise ValueError("save_dir is outside the allowed roots. Move to a configured allowed location first.")
    mapi = _core._mapi()
    try:
        msg = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            raise ValueError(f"Message not found (entry_id={entry_id!r}) — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    effective_account = _effective_account_for_item(mapi, msg, account_email)
    try:
        parent_name = msg.Parent.Name
    except Exception as exc:
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            raise ValueError(f"Cannot determine source folder for entry_id={entry_id!r} — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Cannot determine source folder for entry_id={entry_id!r}") from exc
    _assert_allowed(parent_name, effective_account)
    cfg = get_effective_config(effective_account)
    max_bytes = cfg.attachment_max_mb * 1024 * 1024
    dest.mkdir(parents=True, exist_ok=True)
    saved = []
    seen_names: set[str] = set()
    for att in msg.Attachments:
        size = getattr(att, "Size", 0)
        raw_name = att.FileName or f"attachment_{att.Index}"
        raw_basename = Path(raw_name).name
        if not raw_basename or raw_basename == ".." or "\x00" in raw_basename:
            raise ValueError(f"Attachment filename is unsafe: {raw_name!r}")
        suffix = Path(raw_basename).suffix
        stem = Path(raw_basename).stem
        name = f"{_core._redact(stem, effective_account)}{suffix}"
        if not name or name == ".." or "\x00" in name:
            raise ValueError("Attachment filename is unsafe after privacy redaction.")
        if suffix.lower() in _folders._DANGEROUS_SAVE_EXTENSIONS:
            saved.append({"name": name, "skipped": True, "reason": "file extension blocked by policy"})
            continue
        if suffix.lower() not in _folders._ALLOWED_SAVE_EXTENSIONS:
            saved.append({"name": name, "skipped": True, "reason": "file extension not in allowlist"})
            continue
        if size > max_bytes:
            saved.append({"name": name, "skipped": True, "reason": f"size {size} > cap {max_bytes}"})
            continue
        stem = Path(name).stem
        suffix = Path(name).suffix
        with _SAVE_LOCK:
            candidate = name
            counter = 0
            while candidate.lower() in seen_names or (dest / candidate).exists():
                counter += 1
                candidate = f"{stem}_{counter}{suffix}"
            name = candidate
            seen_names.add(name.lower())
            out_path = dest / name
            if not out_path.resolve().is_relative_to(dest):
                raise ValueError(f"Resolved attachment path escapes save_dir: {out_path!r}")
            att.SaveAsFile(str(out_path))
            size_bytes = out_path.stat().st_size
        saved.append({"name": name, "path": str(out_path), "size_bytes": size_bytes})
    return saved
