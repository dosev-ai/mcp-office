"""Recipient and contact search operations."""
from __future__ import annotations

from mailmcp import _core, _folders
from mailmcp._core import get_effective_config, _assert_allowed, _redact
from mailmcp._formatters import _sql_escape


def search_recipients(
    name: str | None = None,
    query: str | None = None,
    account_email: str | None = None,
) -> list[dict]:
    effective_name = name if name is not None else query
    if not isinstance(effective_name, str) or not effective_name.strip():
        raise ValueError("A non-blank 'name' or 'query' must be provided.")
    effective_name = effective_name.strip()
    cfg = get_effective_config(account_email)
    limit = cfg.max_items
    if limit <= 0:
        return []

    mapi = _core._mapi()
    results = []
    try:
        recip = mapi.CreateRecipient(effective_name)
        recip.Resolve()
        if recip.Resolved and len(results) < limit:
            entry = recip.AddressEntry
            raw_address = getattr(entry, "Address", None) or ""
            if raw_address.startswith(("/o=", "/O=")):
                try:
                    smtp = entry.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                    if smtp:
                        raw_address = smtp
                except Exception:
                    pass
            results.append({"name": entry.Name, "email": raw_address or None, "type": getattr(entry, "Type", None), "source": "gal"})
    except Exception:
        pass
    try:
        _assert_allowed("Contacts", account_email)
        contacts = (
            _folders._folder_by_name_for_account("Contacts", account_email=account_email)
            if account_email is not None
            else _folders._folder_by_name("Contacts")
        )
        items = contacts.Items.Restrict(f"@SQL=\"urn:schemas:contacts:cn\" LIKE '%{_sql_escape(effective_name)}%'")
        for contact in items:
            if len(results) >= limit:
                break
            results.append({
                "name": getattr(contact, "FullName", None) or getattr(contact, "Subject", None),
                "email": getattr(contact, "Email1Address", None),
                "source": "contacts",
            })
    except Exception:
        pass
    if cfg.redact_mode != "none":
        results = [
            {k: _redact(str(v), account_email) if isinstance(v, str) else v for k, v in row.items()}
            for row in results
        ]
    return results[:limit]
