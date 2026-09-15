"""Send, message mutation, thread, and forwarding-rule operations."""
from __future__ import annotations

import logging

from mailmcp import _core
from mailmcp import _folders
from mailmcp._core import get_config, get_effective_config, _assert_allowed, _EMAIL_VALIDATE_RE, _redact, _assert_domains_allowed, _resolve_smtp_from_entry
from mailmcp._formatters import _fmt_date, _resolve_sender_email
from mailmcp._folders import _FLAG_STATUS_MAP
from mailmcp._mail_calendar import _FB_STATUS as _FB_STATUS, check_freebusy as check_freebusy, create_meeting_draft as create_meeting_draft

from mailmcp._item_guards import _assert_draft_item, _assert_mail_item

logger = logging.getLogger(__name__)
_MAX_THREAD_SCAN_ITEMS = 500


def _jet_escape(value: str) -> str:
    if "\x00" in value:
        raise ValueError("Search/filter string contains null bytes which are not permitted.")
    return value.replace("'", "''")


def _effective_account_for_item(item, account_email: str | None) -> str | None:
    if account_email is not None:
        _folders._assert_object_belongs_to_account(item, account_email, object_label="message")
        return account_email.strip().lower()
    if not _core._account_overrides:
        return None
    mapi = _core._mapi()
    store = _folders._resolve_store_for_object(item)
    if store is None:
        raise PermissionError("Cannot determine the message's Outlook account.")
    store_id = _folders._store_id(store)
    if not store_id:
        raise PermissionError("Cannot determine the message's Outlook store identity.")
    index = _folders._build_account_store_index(mapi)
    owners = [
        smtp for smtp, store_ids in index["store_ids_by_smtp"].items()
        if store_id in store_ids
    ]
    if len(owners) != 1:
        raise PermissionError("Cannot resolve a unique Outlook account for the message.")
    return owners[0]


def send_mail(entry_id: str, confirm: bool = False, account_email: str | None = None) -> dict:
    cfg = get_effective_config(account_email)
    if not cfg.enable_send:
        raise PermissionError("send_mail is disabled. Set OUTLOOK_ENABLE_SEND=true to enable.")
    if not confirm:
        raise ValueError("confirm=True is required to send. This is a safety gate.")
    mapi = _core._mapi()
    try:
        msg = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    _assert_draft_item(msg, account_email)
    recipients = []
    for i in range(1, msg.Recipients.Count + 1):
        recipients.append(_resolve_smtp_from_entry(msg.Recipients.Item(i).AddressEntry))
    _assert_domains_allowed(recipients, account_email=account_email)
    msg.Send()
    return {"status": "sent", "entry_id": entry_id}


def move_message(entry_id: str, target_folder: str, confirm: bool = False, account_email: str | None = None) -> dict:
    _core._assert_write_enabled(account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True required to move a message.")
    _assert_allowed(target_folder, account_email)
    mapi = _core._mapi()
    try:
        item = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    _assert_mail_item(item)
    _folders._assert_object_belongs_to_account(item, account_email, object_label="message")
    try:
        _assert_allowed(item.Parent.Name, account_email)
    except AttributeError as exc:
        raise ValueError(f"Cannot determine source folder for entry_id={entry_id!r}") from exc
    target = _folders._folder_by_name_for_account(target_folder, account_email=account_email) if account_email is not None else _folders._folder_by_name(target_folder)
    moved = item.Move(target)
    new_entry_id = getattr(moved, "EntryID", entry_id) if moved is not None else entry_id
    return {"ok": True, "entry_id": new_entry_id, "moved_to": target_folder}


def flag_message(entry_id: str, flag_status: str, confirm: bool = False, account_email: str | None = None) -> dict:
    _core._assert_write_enabled(account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True required to flag a message.")
    status = flag_status.lower().strip()
    if status not in _FLAG_STATUS_MAP:
        raise ValueError(f"Invalid flag_status {flag_status!r}.")
    mapi = _core._mapi()
    try:
        item = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    _assert_mail_item(item)
    _folders._assert_object_belongs_to_account(item, account_email, object_label="message")
    _assert_allowed(item.Parent.Name, account_email)
    item.FlagStatus = _FLAG_STATUS_MAP[status]
    item.Save()
    return {"ok": True, "entry_id": entry_id, "flag_status": status}


def mark_read(entry_id: str, read: bool = True, confirm: bool = False, account_email: str | None = None) -> dict:
    _core._assert_write_enabled(account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True required to mark a message as read/unread.")
    mapi = _core._mapi()
    try:
        item = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    _assert_mail_item(item)
    _folders._assert_object_belongs_to_account(item, account_email, object_label="message")
    _assert_allowed(item.Parent.Name, account_email)
    item.UnRead = not read
    item.Save()
    return {"ok": True, "entry_id": entry_id, "read": read}


def get_conversation_thread(
    conversation_id: str | None = None,
    max_items: int = 50,
    entry_id: str | None = None,
    account_email: str | None = None,
) -> dict:
    if conversation_id is None and entry_id is None:
        raise ValueError("At least one of 'conversation_id' or 'entry_id' must be provided.")
    effective_account = account_email.strip().lower() if isinstance(account_email, str) and account_email.strip() else None
    if entry_id is None and effective_account is None and _core._account_overrides:
        raise ValueError("account_email is required for conversation_id-only lookup when per-account policy is configured.")
    if conversation_id is None:
        mapi = _core._mapi()
        try:
            msg = mapi.GetItemFromID(entry_id)
        except Exception as exc:
            raise ValueError(f"Could not resolve entry_id to a conversation (entry_id={entry_id!r})") from exc
        effective_account = _effective_account_for_item(msg, effective_account)
        _assert_allowed(msg.Parent.Name, effective_account)
        conversation_id = getattr(msg, "ConversationID", None)
        if not conversation_id:
            raise ValueError(f"Message with entry_id={entry_id!r} has no ConversationID.")
    cfg = get_effective_config(effective_account)
    limit = min(max_items, cfg.max_items)
    messages: list[dict] = []
    conv_topic = None
    if entry_id is not None:
        try:
            source = _core._mapi().GetItemFromID(entry_id)
            source_account = _effective_account_for_item(source, effective_account)
            if effective_account is None:
                effective_account = source_account
                cfg = get_effective_config(effective_account)
                limit = min(max_items, cfg.max_items)
            conv_topic = getattr(source, "ConversationTopic", None)
        except Exception:
            pass
    folder_resolver = (
        (lambda name: _folders._folder_by_name_for_account(name, account_email=effective_account))
        if effective_account is not None
        else _folders._folder_by_name
    )
    if conv_topic is None:
        scan_count = 0
        for folder_name in cfg.allowlist_folders:
            try:
                items = folder_resolver(folder_name).Items
                for msg in items:
                    scan_count += 1
                    if scan_count > _MAX_THREAD_SCAN_ITEMS:
                        break
                    if getattr(msg, "ConversationID", None) == conversation_id:
                        conv_topic = getattr(msg, "ConversationTopic", None)
                        break
                if conv_topic is not None or scan_count > _MAX_THREAD_SCAN_ITEMS:
                    break
            except Exception:
                pass
    if conv_topic is None:
        return {"ok": True, "conversation_id": conversation_id, "count": 0, "messages": []}
    escaped_topic = _jet_escape(conv_topic)
    for folder_name in cfg.allowlist_folders:
        try:
            restricted = folder_resolver(folder_name).Items.Restrict(f"[ConversationTopic] = '{escaped_topic}'")
            for msg in restricted:
                if getattr(msg, "ConversationID", None) != conversation_id:
                    continue
                try:
                    body = getattr(msg, "Body", "") or ""
                    messages.append({
                        "entry_id": msg.EntryID,
                        "subject": _redact(getattr(msg, "Subject", "") or "", effective_account),
                        "sender_name": _redact(getattr(msg, "SenderName", "") or "", effective_account),
                        "sender_email": _redact(_resolve_sender_email(msg), effective_account),
                        "received_time": _fmt_date(getattr(msg, "ReceivedTime", None)),
                        "body_preview": _redact(body[: cfg.max_body_chars], effective_account),
                        "folder_name": folder_name,
                    })
                except Exception as exc:
                    logger.debug("Skipping message in conversation thread: %s", exc)
        except Exception as exc:
            logger.debug("Skipping folder %s in conversation thread: %s", folder_name, exc)
    messages.sort(key=lambda row: row["received_time"] or "9999")
    messages = messages[:limit]
    return {"ok": True, "conversation_id": conversation_id, "count": len(messages), "messages": messages}


def mark_junk(entry_id: str, confirm: bool = False, account_email: str | None = None) -> dict:
    _core._assert_write_enabled(account_email=account_email)
    if not confirm:
        raise ValueError("confirm=True is required to mark a message as junk.")
    _assert_allowed("Junk Email", account_email)
    mapi = _core._mapi()
    try:
        item = mapi.GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"Message not found (entry_id={entry_id!r})") from exc
    _assert_mail_item(item)
    _folders._assert_object_belongs_to_account(item, account_email, object_label="message")
    _assert_allowed(item.Parent.Name, account_email)
    junk = _folders._get_junk_folder(account_email=account_email)
    moved = item.Move(junk)
    return {"ok": True, "entry_id": getattr(moved, "EntryID", entry_id) if moved is not None else entry_id, "moved_to": "Junk Email"}


def _get_rules_collection(account_email: str | None = None):
    mapi = _core._mapi()
    if account_email is None:
        store = getattr(mapi, "DefaultStore", None)
        if store is None:
            try:
                default_inbox = mapi.GetDefaultFolder(_folders._OL_INBOX)
                store = _folders._resolve_store_for_object(default_inbox)
            except Exception as exc:
                raise RuntimeError(f"Cannot resolve the default Outlook store for rules: {exc}") from exc
        get_rules = getattr(store, "GetRules", None)
        if not callable(get_rules):
            raise RuntimeError("Cannot access Outlook rules: default Store.GetRules() is unavailable.")
        try:
            return get_rules()
        except Exception as exc:
            raise RuntimeError(f"Cannot access Outlook rules: {exc}") from exc
    store = _folders._find_store_for_account(account_email, mapi=mapi)
    get_rules = getattr(store, "GetRules", None)
    if not callable(get_rules):
        raise PermissionError(f"Cannot access Outlook rules for account '{account_email}'.")
    try:
        return get_rules()
    except Exception as exc:
        raise RuntimeError(f"Cannot access Outlook rules for account '{account_email}': {exc}") from exc


def _assert_rules_enabled(account_email: str | None = None) -> None:
    if not get_effective_config(account_email).enable_rules:
        raise PermissionError("Forwarding-rule mutations are disabled. Set OUTLOOK_ENABLE_RULES=true to enable.")


def list_forwarding_rules(account_email: str | None = None) -> dict:
    rules = _get_rules_collection(account_email=account_email)
    result = []
    for i in range(rules.Count):
        try:
            rule = rules.Item(i + 1)
            fwd_addrs = []
            try:
                action = rule.Actions.ForwardTo
                for j in range(action.RecipientCount):
                    fwd_addrs.append(_redact(getattr(action.Item(j + 1), "Address", "") or "", account_email=account_email))
            except Exception:
                pass
            if fwd_addrs:
                result.append({"rule_id": str(i + 1), "rule_name": _redact(str(rule.Name or ""), account_email=account_email), "enabled": bool(rule.Enabled), "forward_to": fwd_addrs})
        except Exception as exc:
            logger.warning("Skipping Outlook rule %d: %s", i + 1, exc)
    return {"rules": result, "count": len(result)}


def create_forwarding_rule(forward_to: str, folder_name: str, subject_filter: str | None = None, confirm: bool = False, account_email: str | None = None) -> dict:
    _assert_rules_enabled(account_email)
    if not confirm:
        raise ValueError("confirm=True is required to create a forwarding rule.")
    if not _EMAIL_VALIDATE_RE.match(forward_to):
        raise ValueError(f"Invalid forward_to email address: {forward_to!r}")
    _assert_domains_allowed([forward_to], account_email=account_email)
    _assert_allowed(folder_name, account_email)
    rules = _get_rules_collection(account_email=account_email)
    rule_name = f"Forward to {forward_to}" + (f" ({subject_filter})" if subject_filter else "")
    try:
        rule = rules.Create(rule_name, 0)
        target = _folders._folder_by_name_for_account(folder_name, account_email=account_email) if account_email is not None else _folders._folder_by_name(folder_name)
        condition = rule.Conditions.FolderCondition
        condition.Folders.Add(target)
        condition.Enabled = True
        if subject_filter:
            subject_condition = rule.Conditions.Subject
            subject_condition.Text = [subject_filter]
            subject_condition.Enabled = True
        action = rule.Actions.ForwardTo
        recipient = action.Recipients.Add(forward_to)
        recipient.Resolve()
        if not getattr(recipient, "Resolved", False):
            raise PermissionError("Forwarding-rule recipient could not be resolved.")
        address_entry = getattr(recipient, "AddressEntry", None)
        if address_entry is None:
            raise PermissionError("Cannot verify the resolved forwarding-rule recipient.")
        resolved_smtp = _resolve_smtp_from_entry(address_entry)
        _assert_domains_allowed([resolved_smtp], account_email=account_email)
        action.Enabled = True
        rules.Save(True)
    except Exception as exc:
        raise RuntimeError(f"Cannot create forwarding rule: {exc}") from exc
    return {"ok": True, "rule_id": str(rules.Count), "rule_name": _redact(rule_name, account_email=account_email), "forward_to": _redact(forward_to, account_email=account_email)}


def delete_forwarding_rule(rule_id: str, confirm: bool = False, account_email: str | None = None) -> dict:
    _assert_rules_enabled(account_email)
    if not confirm:
        raise ValueError("confirm=True is required to delete a forwarding rule.")
    rules = _get_rules_collection(account_email=account_email)
    try:
        idx = int(rule_id)
    except (ValueError, TypeError):
        raise ValueError(f"rule_id must be a positive integer string, got {rule_id!r}")
    if idx < 1 or idx > rules.Count:
        raise ValueError(f"rule_id {rule_id!r} out of range (1–{rules.Count}).")
    try:
        rule_name = str(rules.Item(idx).Name or "")
        rules.Remove(idx)
        rules.Save(True)
    except Exception as exc:
        raise RuntimeError(f"Cannot delete rule {rule_id!r}: {exc}") from exc
    redacted = _redact(rule_name, account_email=account_email)
    return {"ok": True, "rule_id": rule_id, "rule_name": redacted, "message": f"Deleted rule: {redacted}"}
