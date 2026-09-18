"""Governed routine: search Outlook messages. Read-only."""
from __future__ import annotations

import os


def search_messages_v1(
    folder_name: str = "Inbox",
    subject: str | None = None,
    sender: str | None = None,
    body_contains: str | None = None,
    max_results: int = 20,
) -> dict:
    raw_allowlist = os.environ.get("OUTLOOK_ALLOWLIST_FOLDERS", "").strip()
    if not raw_allowlist:
        return {"error": "OUTLOOK_ALLOWLIST_FOLDERS not configured", "error_kind": "config_error"}
    allowlist = [f.strip() for f in raw_allowlist.split(",") if f.strip()]
    if folder_name.lower() not in [f.lower() for f in allowlist] and "*" not in allowlist:
        return {"error": "folder not in allowlist", "error_kind": "allowlist_violation"}
    try:
        from mailmcp._messages import search_messages
        results = search_messages(folder_name=folder_name, subject=subject, sender=sender, body_contains=body_contains, top=max_results)
        messages = results.get("messages", []) if isinstance(results, dict) else list(results)
        return {
            "messages": messages,
            "message_count": len(messages),
            "folder_name": folder_name,
            "evidence": {"routine": "outlook.search_messages_v1", "max_results": max_results},
        }
    except Exception as exc:
        return {"error": str(exc), "error_kind": "execution_error"}
