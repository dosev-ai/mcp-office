"""Outlook forwarding-rule tool handlers."""
from __future__ import annotations

from mailmcp import outlook_com as ol


def handle_forwarding_rule(
    operation: str,
    rule_id: str | None = None,
    forward_to: str | None = None,
    folder_name: str | None = None,
    subject_filter: str | None = None,
    confirm: bool = False,
    account_email: str | None = None,
) -> dict:
    op = operation.lower()
    if op == "list":
        return ol.list_forwarding_rules(account_email=account_email)
    if op == "create":
        if forward_to is None:
            raise ValueError("forward_to is required for operation='create'")
        if folder_name is None:
            raise ValueError("folder_name is required for operation='create'")
        return ol.create_forwarding_rule(
            forward_to=forward_to,
            folder_name=folder_name,
            subject_filter=subject_filter,
            confirm=confirm,
            account_email=account_email,
        )
    if op == "delete":
        if rule_id is None:
            raise ValueError("rule_id is required for operation='delete'")
        return ol.delete_forwarding_rule(rule_id=rule_id, confirm=confirm, account_email=account_email)
    raise ValueError(f"Unknown operation: {operation!r}. Must be one of: list, create, delete")


def outlook_forwarding_rule(
    operation: str,
    rule_id: str | None = None,
    forward_to: str | None = None,
    folder_name: str | None = None,
    subject_filter: str | None = None,
    confirm: bool = False,
    account_email: str | None = None,
) -> dict:
    return handle_forwarding_rule(
        operation=operation,
        rule_id=rule_id,
        forward_to=forward_to,
        folder_name=folder_name,
        subject_filter=subject_filter,
        confirm=confirm,
        account_email=account_email,
    )
