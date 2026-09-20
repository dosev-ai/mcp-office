"""Adjacent-path negatives for the same account and recipient policy boundary."""
from __future__ import annotations

import pytest

from mailmcp import _core, _folders, _mail_calendar


def test_store_id_cannot_be_overridden_by_display_name(mailbox):
    mailbox.stores[1].DisplayName = "owner-a@example.com"
    index = _folders._build_account_store_index(mailbox.mapi)
    assert not _folders._store_matches_account(mailbox.stores[1], "owner-a@example.com", index)
    assert _folders._store_matches_account(mailbox.stores[0], "owner-a@example.com", index)


def test_account_resolution_skips_spoofed_first_store(mailbox):
    mailbox.stores[1].DisplayName = "owner-a@example.com"
    mailbox.mapi.Stores = list(reversed(mailbox.stores))
    assert _folders._find_store_for_account("owner-a@example.com") is mailbox.stores[0]


def test_named_account_without_delivery_store_identity_fails_closed(mailbox):
    mailbox.mapi.Accounts[0].DeliveryStore = None
    with pytest.raises(PermissionError):
        _folders._find_store_for_account("owner-a@example.com")


def test_account_cache_cannot_reuse_other_store_same_display(mailbox):
    _core._set_folder_cache({"owner-a@example.com/contacts": mailbox.folders[("b", "Contacts")]})
    assert _folders._folder_by_name_for_account("Contacts", "owner-a@example.com") is mailbox.folders[("a", "Contacts")]


def test_disjoint_account_domain_policy_denies_instead_of_expanding():
    global_cfg = _core.OutlookConfig(allowlist_domains=["example.com"])
    override = _core.OutlookAccountOverride(email="owner@example.com", allowlist_domains=["other.test"])
    with pytest.raises(PermissionError, match="domain"):
        _core._merge_config(global_cfg, override)


def test_domain_intersection_preserves_shared_subset():
    global_cfg = _core.OutlookConfig(allowlist_domains=["example.com", "other.test"])
    override = _core.OutlookAccountOverride(email="owner@example.com", allowlist_domains=["example.com"])
    assert _core._merge_config(global_cfg, override).allowlist_domains == ["example.com"]


def test_meeting_creation_rejects_inverted_range_before_com(mailbox):
    with pytest.raises(ValueError, match="after"):
        _mail_calendar.create_meeting_draft("Synthetic", "2026-01-01T11:00:00", "2026-01-01T10:00:00", ["person@example.com"], confirm=True)
    _core._get_outlook.assert_not_called()


def test_freebusy_nonstring_address_rejected_before_com(mailbox):
    with pytest.raises(ValueError, match="email"):
        _mail_calendar.check_freebusy([None], "2026-01-01T09:00:00", "2026-01-01T10:00:00")
    mailbox.get_mapi.assert_not_called()
