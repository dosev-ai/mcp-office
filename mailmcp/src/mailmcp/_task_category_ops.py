"""Framework-agnostic tool handlers for tasks, meeting requests, and categories."""
from __future__ import annotations

from mailmcp import outlook_com as ol


def outlook_list_tasks(include_completed: bool = False, due_before: str | None = None, top: int = 20, account_email: str | None = None) -> dict:
    return ol.list_tasks(include_completed=include_completed, due_before=due_before, top=top, account_email=account_email)


def outlook_create_task(subject: str, body: str | None = None, due_date: str | None = None, priority: str = "normal", confirm: bool = False, account_email: str | None = None) -> dict:
    return ol.create_task(subject=subject, body=body, due_date=due_date, priority=priority, confirm=confirm, account_email=account_email)


def outlook_complete_task(entry_id: str, confirm: bool = False, account_email: str | None = None) -> dict:
    return ol.complete_task(entry_id=entry_id, confirm=confirm, account_email=account_email)


def outlook_list_meeting_requests(top: int = 20, folder_name: str | None = None, account_email: str | None = None) -> dict:
    return ol.list_meeting_requests(top=top, folder_name=folder_name, account_email=account_email)


def outlook_respond_to_meeting(entry_id: str, response: str, confirm: bool = False, account_email: str | None = None) -> dict:
    return ol.respond_to_meeting(entry_id=entry_id, response=response, confirm=confirm, account_email=account_email)


def outlook_list_categories() -> dict:
    return ol.list_categories()


def outlook_set_message_category(entry_id: str, categories: list[str], confirm: bool = False, account_email: str | None = None) -> dict:
    return ol.set_message_category(entry_id=entry_id, categories=categories, confirm=confirm, account_email=account_email)


def outlook_get_messages_by_category(category: str, folder_name: str | None = None, top: int = 20, account_email: str | None = None) -> dict:
    return ol.get_messages_by_category(category=category, folder_name=folder_name, top=top, account_email=account_email)


def outlook_task(operation: str, entry_id: str | None = None, subject: str | None = None, body: str | None = None, due_date: str | None = None, priority: str = "normal", include_completed: bool = False, due_before: str | None = None, top: int = 20, confirm: bool = False, auto_create: bool = False, account_email: str | None = None) -> dict:
    op = operation.lower()
    if op == "list":
        return outlook_list_tasks(include_completed=include_completed, due_before=due_before, top=top, account_email=account_email)
    if op == "create":
        if subject is None:
            raise ValueError("subject is required for operation='create'")
        return outlook_create_task(subject=subject, body=body, due_date=due_date, priority=priority, confirm=confirm, account_email=account_email)
    if op == "complete":
        if entry_id is None:
            raise ValueError("entry_id is required for operation='complete'")
        return outlook_complete_task(entry_id=entry_id, confirm=confirm, account_email=account_email)
    if op == "extract_from_message":
        if entry_id is None:
            raise ValueError("entry_id is required for operation='extract_from_message'")
        return ol.extract_tasks_from_message(entry_id=entry_id, auto_create=auto_create, confirm=confirm, account_email=account_email)
    raise ValueError(f"Unknown operation: {operation!r}. Must be one of: list, create, complete, extract_from_message")


def outlook_category(operation: str, entry_id: str | None = None, categories: list[str] | None = None, category: str | None = None, folder_name: str | None = None, top: int = 20, confirm: bool = False, account_email: str | None = None) -> dict:
    op = operation.lower()
    if op == "list":
        return outlook_list_categories()
    if op == "set":
        if entry_id is None:
            raise ValueError("entry_id is required for operation='set'")
        if categories is None:
            raise ValueError("categories is required for operation='set' (pass [] to clear)")
        return outlook_set_message_category(entry_id=entry_id, categories=categories, confirm=confirm, account_email=account_email)
    if op == "get_messages":
        if category is None:
            raise ValueError("category is required for operation='get_messages'")
        return outlook_get_messages_by_category(category=category, folder_name=folder_name, top=top, account_email=account_email)
    raise ValueError(f"Unknown operation: {operation!r}. Must be one of: list, set, get_messages")
