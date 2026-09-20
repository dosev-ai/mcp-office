"""Governed routine: compose an Outlook draft. Never sends."""
from __future__ import annotations

import os


def compose_draft_v1(
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    confirm: bool = False,
    source_entry_id: str | None = None,
) -> dict:
    if os.environ.get("OUTLOOK_ENABLE_WRITE", "").strip().lower() != "true":
        return {"error": "OUTLOOK_ENABLE_WRITE not set to true", "error_kind": "gate_error"}
    if not confirm:
        return {"error": "confirm=True is required to create a draft. This is a safety gate.", "error_kind": "gate_error"}
    try:
        from mailmcp._mail_compose import compose_mail
        result = compose_mail(to=to, subject=subject, body=body, cc=cc, confirm=True, source_entry_id=source_entry_id)
        envelope: dict = {
            "status": "draft_saved",
            "entry_id": result.get("entry_id", ""),
            "subject": subject,
            "to": to,
            "evidence": {"routine": "outlook.compose_draft_v1"},
        }
        if source_entry_id:
            envelope["source_entry_id"] = source_entry_id
        return envelope
    except Exception as exc:
        return {"error": str(exc), "error_kind": "execution_error"}
