from __future__ import annotations

from dataclasses import replace
from unittest.mock import Mock

import pytest

from mailmcp import _contacts_com, _core, _tasks


@pytest.mark.parametrize("mode", ["none", "emails", "emails+domains"])
def test_contact_redaction_uses_effective_account_policy(mailbox, mode, monkeypatch):
    _core.set_config(replace(_core.get_config(), redact_mode="none"))
    monkeypatch.setattr(_core, "_account_overrides", {
        "owner-a@example.com": _core.OutlookAccountOverride(email="owner-a@example.com", redact_mode=mode),
    })
    mailbox.item.Class = 40
    mailbox.item.Parent = mailbox.folders[("a", "Contacts")]
    mailbox.item.Email1Address = "person@example.com"
    mailbox.item.FullName = "Contact person@example.com"
    mailbox.item.CompanyName = "example.com"
    result = _contacts_com.get_contact("synthetic-item", account_email="owner-a@example.com")
    assert result["entry_id"] == "synthetic-item"
    if mode == "none":
        assert result["email"] == "person@example.com"
    else:
        assert result["email"] == "[email]"
        assert "person@example.com" not in result["full_name"]
    if mode == "emails+domains":
        assert result["company"] == "[domain]"


@pytest.mark.parametrize("change", ["remove-account", "remove-folders", "expand-folders", "remove-domains", "expand-domains"])
def test_reload_cannot_expand_account_restrictions(monkeypatch, change):
    global_config = _core.OutlookConfig(allowlist_folders=["Inbox", "Contacts"], allowlist_domains=["example.com", "other.test"])
    original = _core.OutlookAccountOverride(email="owner-a@example.com", allowlist_folders=["Inbox"], allowlist_domains=["example.com"])
    monkeypatch.setattr(_core, "_config", global_config)
    monkeypatch.setattr(_core, "_startup_config", global_config)
    monkeypatch.setattr(_core, "_account_overrides", {original.email: original})
    monkeypatch.setattr(_core, "_startup_account_overrides", {original.email: original})
    monkeypatch.setenv("OUTLOOK_ALLOWLIST_FOLDERS", "Inbox,Contacts")
    monkeypatch.setenv("OUTLOOK_ALLOWLIST_DOMAINS", "example.com,other.test")
    if change != "remove-account":
        monkeypatch.setenv("OUTLOOK_ACCOUNT_1_EMAIL", original.email)
        if change != "remove-folders":
            monkeypatch.setenv("OUTLOOK_ACCOUNT_1_ALLOWLIST_FOLDERS", "Inbox,Contacts" if change == "expand-folders" else "Inbox")
        if change != "remove-domains":
            monkeypatch.setenv("OUTLOOK_ACCOUNT_1_ALLOWLIST_DOMAINS", "example.com,other.test" if change == "expand-domains" else "example.com")
    with pytest.raises(PermissionError):
        _core.reload_config()
    assert _core._config is global_config
    assert _core._account_overrides[original.email] is original


def test_reload_allows_restriction_tightening(monkeypatch):
    config = _core.OutlookConfig(allowlist_folders=["Inbox", "Contacts"])
    monkeypatch.setattr(_core, "_config", config)
    monkeypatch.setattr(_core, "_startup_config", config)
    monkeypatch.setenv("OUTLOOK_ALLOWLIST_FOLDERS", "Inbox")
    assert _core.reload_config()["allowlist_folders"] == ["Inbox"]


@pytest.mark.parametrize("auto_create", [False, True])
def test_task_extraction_redacts_response_but_preserves_created_content(mailbox, monkeypatch, auto_create):
    _core.set_config(replace(_core.get_config(), redact_mode="emails"))
    mailbox.item.Parent = mailbox.folders[("a", "Inbox")]
    mailbox.item.Body = "TODO: Reply to person@example.com"
    mailbox.item.Subject = "Message from person@example.com"
    monkeypatch.setattr(_tasks, "_mapi", mailbox.get_mapi)
    creator = Mock(return_value={"entry_id": "synthetic-task", "subject": "Reply to person@example.com"})
    monkeypatch.setattr(_tasks, "create_task", creator)
    result = _tasks.extract_tasks_from_message("synthetic-item", auto_create=auto_create, confirm=auto_create)
    assert "person@example.com" not in result["source_subject"]
    assert "person@example.com" not in result["candidates"][0]["title"]
    if auto_create:
        assert creator.call_args.kwargs["subject"] == "Reply to person@example.com"
        assert "person@example.com" not in result["created"][0]["subject"]


@pytest.mark.parametrize("field,value", [("redact_mode", "none"), ("max_items", 50), ("max_body_chars", 4000)])
def test_removing_account_profile_cannot_widen_privacy_or_read_limits(monkeypatch, field, value):
    config = _core.OutlookConfig()
    override = _core.OutlookAccountOverride(email="owner-a@example.com", redact_mode="emails", max_items=10, max_body_chars=100)
    monkeypatch.setattr(_core, "_config", config)
    monkeypatch.setattr(_core, "_startup_config", config)
    # Vary one restriction so each case proves its own guard.
    override = replace(override, redact_mode="none", max_items=50, max_body_chars=4000)
    override = replace(override, **{field: "emails" if field == "redact_mode" else value // 10})
    monkeypatch.setattr(_core, "_account_overrides", {override.email: override})
    monkeypatch.setattr(_core, "_startup_account_overrides", {override.email: override})
    with pytest.raises(PermissionError):
        _core.reload_config()
    assert _core._account_overrides[override.email] is override
