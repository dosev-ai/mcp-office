"""Compare effective policy before atomically applying a configuration reload."""
from __future__ import annotations

from typing import Any


def _assert_no_scope_expansion(startup: Any, candidate: Any) -> None:
    old_folders = {value.strip().lower() for value in startup.allowlist_folders}
    new_folders = {value.strip().lower() for value in candidate.allowlist_folders}
    if "*" not in old_folders and not new_folders.issubset(old_folders):
        raise PermissionError("reload_config: cannot expand folder access at runtime; restart required.")
    old_domains = {value.strip().lower() for value in startup.allowlist_domains}
    new_domains = {value.strip().lower() for value in candidate.allowlist_domains}
    # An empty domain allowlist means unrestricted, not deny-all.
    if old_domains and (not new_domains or not new_domains.issubset(old_domains)):
        raise PermissionError("reload_config: cannot expand recipient domains at runtime; restart required.")
    redact_rank = {"none": 0, "emails": 1, "emails+domains": 2}
    if startup.redact_mode not in redact_rank or candidate.redact_mode not in redact_rank:
        raise PermissionError("reload_config: invalid redaction policy; restart with a valid mode.")
    if redact_rank[candidate.redact_mode] < redact_rank[startup.redact_mode]:
        raise PermissionError("reload_config: cannot weaken redaction at runtime; restart required.")
    for field in ("max_items", "max_body_chars", "attachment_max_mb"):
        if getattr(candidate, field) > getattr(startup, field):
            raise PermissionError(f"reload_config: cannot expand {field} at runtime; restart required.")
