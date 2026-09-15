"""Attachment save operations."""
from __future__ import annotations

from pathlib import Path

from mailmcp import _core, _folders
from mailmcp._core import get_config, _assert_allowed


def save_attachments(entry_id: str, save_dir: str) -> list[dict]:
    dest = Path(save_dir).resolve()
    if not any(dest == root or dest.is_relative_to(root) for root in _folders._ALLOWED_ATTACHMENT_ROOTS):
        raise ValueError("save_dir is outside the allowed roots. Move to a configured allowed location first.")
    dest.mkdir(parents=True, exist_ok=True)
    mapi = _core._mapi()
    try:
        msg = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            raise ValueError(f"Message not found (entry_id={entry_id!r}) — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    try:
        parent_name = msg.Parent.Name
    except Exception as exc:
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            raise ValueError(f"Cannot determine source folder for entry_id={entry_id!r} — COM HRESULT {hresult:#010x}") from exc
        raise ValueError(f"Cannot determine source folder for entry_id={entry_id!r}") from exc
    _assert_allowed(parent_name)
    cfg = get_config()
    max_bytes = cfg.attachment_max_mb * 1024 * 1024
    saved = []
    seen_names: set[str] = set()
    for att in msg.Attachments:
        size = getattr(att, "Size", 0)
        raw_name = att.FileName or f"attachment_{att.Index}"
        name = Path(raw_name).name
        if not name or name == ".." or "\x00" in name:
            raise ValueError(f"Attachment filename is unsafe: {raw_name!r}")
        if Path(name).suffix.lower() in _folders._DANGEROUS_SAVE_EXTENSIONS:
            saved.append({"name": name, "skipped": True, "reason": "file extension blocked by policy"})
            continue
        if Path(name).suffix.lower() not in _folders._ALLOWED_SAVE_EXTENSIONS:
            saved.append({"name": name, "skipped": True, "reason": "file extension not in allowlist"})
            continue
        stem = Path(name).stem
        suffix = Path(name).suffix
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
        if size > max_bytes:
            saved.append({"name": name, "skipped": True, "reason": f"size {size} > cap {max_bytes}"})
            continue
        att.SaveAsFile(str(out_path))
        saved.append({"name": name, "path": str(out_path), "size_bytes": out_path.stat().st_size})
    return saved
