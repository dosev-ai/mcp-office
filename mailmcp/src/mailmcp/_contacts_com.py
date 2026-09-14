"""mailmcp._contacts_com — Outlook Contacts COM helpers.

COM layer only — no MCP SDK imports.
All Outlook COM calls are confined to this module.
"""
from __future__ import annotations

import logging
from typing import Any

from mailmcp import _core, _folders
from mailmcp._core import get_effective_config
from mailmcp._formatters import _sql_escape

logger = logging.getLogger(__name__)

# olContact item.Class value (SPEC-GAP-3)
_OL_CONTACT_CLASS = 40


def _contact_to_dict(c: Any) -> dict:
    """Map an Outlook ContactItem COM object to a plain dict.

    Uses exact COM property names (SPEC-GAP-1):
      BusinessTelephoneNumber — not "BusinessTelephone"
      MobileTelephoneNumber   — not "MobilePhone"
    """
    return {
        "entry_id": getattr(c, "EntryID", None),
        "full_name": getattr(c, "FullName", "") or "",
        "first_name": getattr(c, "FirstName", "") or "",
        "last_name": getattr(c, "LastName", "") or "",
        "email": getattr(c, "Email1Address", None),
        "phone_business": getattr(c, "BusinessTelephoneNumber", "") or "",
        "phone_mobile": getattr(c, "MobileTelephoneNumber", "") or "",
        "company": getattr(c, "CompanyName", "") or "",
        "job_title": getattr(c, "JobTitle", "") or "",
        "city": getattr(c, "BusinessAddressCity", "") or "",
        "country": getattr(c, "BusinessAddressCountry", "") or "",
    }


def list_contacts(
    top: int = 50,
    filter_company: str | None = None,
    account_email: str | None = None,
) -> list[dict]:
    """Return up to *top* contacts from the Contacts folder, sorted by last name.

    Requires 'Contacts' in OUTLOOK_ALLOWLIST_FOLDERS.

    If *filter_company* is provided, applies a DASL Restrict filter on the
    company field. The value is passed through _sql_escape() to prevent
    DASL injection (B-001).
    """
    _core._assert_allowed("Contacts", account_email)
    cfg = get_effective_config(account_email)
    cap = min(top, cfg.max_items)

    folder = _folders._folder_by_name_for_account("Contacts", account_email)
    items = folder.Items
    items.Sort("[LastName]")

    if filter_company is not None:
        escaped = _sql_escape(filter_company)  # B-001: injection-safe
        items = items.Restrict(
            f"@SQL=\"urn:schemas:contacts:organizationname\" LIKE '%{escaped}%'"
        )

    results: list[dict] = []
    for c in items:
        if getattr(c, "Class", None) != _OL_CONTACT_CLASS:
            continue
        if len(results) >= cap:
            break
        try:
            results.append(_contact_to_dict(c))
        except Exception:
            logger.warning("Skipping malformed contact item with EntryID=%r", getattr(c, "EntryID", None))
    return results


def search_contacts(
    query: str,
    top: int = 25,
    account_email: str | None = None,
) -> list[dict]:
    """Search contacts by partial match across name, email, and company.

    Builds a DASL OR filter across urn:schemas:contacts:cn,
    urn:schemas:contacts:email1emailaddress, and
    urn:schemas:contacts:organizationname.  All user input is passed
    through _sql_escape() to prevent DASL injection.

    Args:
        query: Non-empty search string.
        top: Maximum results to return (1-100).
        account_email: Optional account to scope the lookup.

    Returns:
        List of contact dicts (via _contact_to_dict).

    Raises:
        ValueError: query is empty or top is out of range.
        PermissionError: 'Contacts' not in allowlist.
    """
    if not isinstance(query, str) or not query.strip():
        raise ValueError("query must be a non-empty string")
    if not isinstance(top, int) or top < 1 or top > 100:
        raise ValueError("top must be an integer between 1 and 100")

    _core._assert_allowed("Contacts", account_email)
    cfg = get_effective_config(account_email)
    cap = min(top, cfg.max_items)

    folder = _folders._folder_by_name_for_account("Contacts", account_email)
    items = folder.Items
    items.Sort("[LastName]")

    escaped = _sql_escape(query.strip())
    dasl_filter = (
        '@SQL='
        f'"urn:schemas:contacts:cn" LIKE \'%{escaped}%\''
        f' OR "urn:schemas:contacts:email1emailaddress" LIKE \'%{escaped}%\''
        f' OR "urn:schemas:contacts:organizationname" LIKE \'%{escaped}%\''
    )
    restricted = items.Restrict(dasl_filter)

    results: list[dict] = []
    for c in restricted:
        if getattr(c, "Class", None) != _OL_CONTACT_CLASS:
            continue
        if len(results) >= cap:
            break
        try:
            results.append(_contact_to_dict(c))
        except Exception:
            logger.warning(
                "Skipping malformed contact item with EntryID=%r",
                getattr(c, "EntryID", None),
            )
    return results


def get_contact(entry_id: str, account_email: str | None = None) -> dict:
    """Return a single contact by EntryID.

    Security (B-002):
    - Wraps GetItemFromID() in try/except; re-raises as ValueError so raw
      pywintypes.com_error never reaches the MCP JSON layer.
    - Validates item.Class == _OL_CONTACT_CLASS (40) BEFORE accessing any
      contact-specific attribute to prevent cross-item type confusion.
    """
    _core._assert_allowed("Contacts", account_email)  # H1-003: allowlist gate FIRST
    mapi = _core._mapi()
    try:
        item = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Contact not found: {entry_id!r}") from exc

    # B-002: check Class BEFORE any attribute access
    if getattr(item, "Class", None) != _OL_CONTACT_CLASS:
        raise ValueError(
            f"Item {entry_id!r} is not a Contact "
            f"(Class={getattr(item, 'Class', None)}, expected {_OL_CONTACT_CLASS})."
        )
    return _contact_to_dict(item)
