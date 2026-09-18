"""mailmcp._contacts_ops — MCP tool handler wrappers for Contacts and folder creation.

Ops layer only — no COM imports, no MCP SDK imports.
All COM calls are delegated to outlook_com (which re-exports from _contacts_com
and _folders).
"""
from __future__ import annotations

import logging

from mailmcp import outlook_com as ol

logger = logging.getLogger(__name__)


def outlook_list_contacts(
    top: int = 50,
    filter_company: str | None = None,
    account_email: str | None = None,
) -> dict:
    """List contacts from the Contacts folder.

    Returns up to *top* contacts sorted by last name. Optionally filter by
    company name (DASL; injection-safe). Requires 'Contacts' in
    OUTLOOK_ALLOWLIST_FOLDERS.

    Returns:
        {"contacts": [<contact_dict>, ...], "count": int}
    """
    contacts = ol.list_contacts(
        top=top,
        filter_company=filter_company,
        account_email=account_email,
    )
    return {"contacts": contacts, "count": len(contacts)}


def outlook_get_contact(
    entry_id: str,
    account_email: str | None = None,
) -> dict:
    """Get a single contact by EntryID.

    Raises ValueError if the EntryID is not found or does not point to a
    Contact item (olContact class == 40).

    Returns:
        Flat dict with keys: entry_id, full_name, first_name, last_name,
        email, phone_business, phone_mobile, company, job_title, city, country.
    """
    return ol.get_contact(entry_id=entry_id, account_email=account_email)


def outlook_search_contacts(
    query: str,
    top: int = 25,
    account_email: str | None = None,
) -> dict:
    """Search contacts by partial match across name, email, and company.

    Returns up to *top* contacts whose display name, email address, or
    company name contains *query* (case-insensitive DASL LIKE).  Requires
    'Contacts' in OUTLOOK_ALLOWLIST_FOLDERS.

    Returns:
        {"contacts": [<contact_dict>, ...], "count": int, "query": str}
    """
    contacts = ol.search_contacts(
        query=query,
        top=top,
        account_email=account_email,
    )
    return {"contacts": contacts, "count": len(contacts), "query": query}


def outlook_create_folder(
    parent_folder: str,
    name: str,
    confirm: bool = False,
    account_email: str | None = None,
) -> dict:
    """Create a subfolder under parent_folder.

    Gate order (B-003):
      1. confirm=True required (ops layer) — fires before any COM call
      2. OUTLOOK_ENABLE_WRITE=true required
      3. parent_folder must be in OUTLOOK_ALLOWLIST_FOLDERS
      4. name validation (non-empty; no \\ / :)
      5. COM call

    Raises:
        ValueError: confirm is False, name is invalid, or folder already exists.
        PermissionError: write is disabled or parent_folder not in allowlist.

    Returns:
        {"ok": True, "name": str, "parent": str, "entry_id": str | None}
    """
    if not confirm:
        raise ValueError(
            "outlook_create_folder requires confirm=True to prevent accidental folder creation."
        )
    return ol.create_folder(
        parent_folder=parent_folder,
        name=name,
        account_email=account_email,
    )
