"""Recipient and contact search operations."""
from __future__ import annotations

from mailmcp import _core, _folders
from mailmcp._core import get_config, _assert_allowed, _redact
from mailmcp._formatters import _sql_escape


def search_recipients(name: str | None = None, query: str | None = None) -> list[dict]:
    effective_name = name if name is not None else query
    if effective_name is None:
        raise ValueError("At least one of 'name' or 'query' must be provided.")
    mapi = _core._mapi()
    results = []
    try:
        recip = mapi.CreateRecipient(effective_name)
        recip.Resolve()
        if recip.Resolved:
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
        _assert_allowed("Contacts")
        contacts = _folders._folder_by_name("Contacts")
        items = contacts.Items.Restrict(f"@SQL=\"urn:schemas:contacts:cn\" LIKE '%{_sql_escape(effective_name)}%'")
        for contact in items:
            results.append({
                "name": getattr(contact, "FullName", None) or getattr(contact, "Subject", None),
                "email": getattr(contact, "Email1Address", None),
                "source": "contacts",
            })
    except Exception:
        pass
    if get_config().redact_mode != "none":
        results = [{k: _redact(str(v)) if isinstance(v, str) else v for k, v in row.items()} for row in results]
    return results
