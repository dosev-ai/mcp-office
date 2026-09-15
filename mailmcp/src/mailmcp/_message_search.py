"""Direct Outlook COM message-search operations for the public MailMCP surface."""
from __future__ import annotations

import concurrent.futures
import os
from typing import Any, Callable

from mailmcp import outlook_com as ol


def _search_timeout_secs() -> int:
    try:
        raw = int(os.getenv("OUTLOOK_SEARCH_TIMEOUT_SECS", "60"))
    except ValueError:
        raw = 60
    return max(5, min(600, raw))


def _run_com_with_timeout(fn: Callable[[], Any], op_label: str) -> Any:
    timeout_secs = _search_timeout_secs()

    def _worker() -> Any:
        import pythoncom as _pythoncom
        _pythoncom.CoInitialize()
        try:
            return fn()
        finally:
            _pythoncom.CoUninitialize()

    executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
    future = executor.submit(_worker)
    try:
        return future.result(timeout=timeout_secs)
    except concurrent.futures.TimeoutError:
        executor.shutdown(wait=False)
        raise TimeoutError(
            f"{op_label} timed out after {timeout_secs}s. Outlook may be syncing, the search filter may be slow, or the MAPI store may be unresponsive. Retry with a tighter scan_limit / subject filter, or set OUTLOOK_SEARCH_TIMEOUT_SECS to extend the budget."
        )
    except Exception:
        executor.shutdown(wait=False)
        raise
    else:
        executor.shutdown(wait=False)


def outlook_search_messages(
    folder_name: str = "Inbox",
    subject: str | None = None,
    sender: str | None = None,
    body_contains: str | None = None,
    top: int = 20,
    account_email: str | None = None,
    scan_limit: int | None = None,
) -> dict:
    def _call() -> Any:
        return ol.search_messages(
            folder_name=folder_name,
            subject=subject,
            sender=sender,
            body_contains=body_contains,
            top=top,
            account_email=account_email,
            scan_limit=scan_limit,
            include_scan_metadata=True,
        )

    raw = _run_com_with_timeout(_call, op_label=f"outlook_search_messages(folder={folder_name!r})")
    if isinstance(raw, dict):
        messages = raw.get("messages")
        result = list(messages) if isinstance(messages, list) else []
        scan_capped = bool(raw.get("scan_capped", False))
    else:
        result = list(raw)
        scan_capped = False
    return {"messages": result, "count": len(result), "scan_capped": scan_capped}


def outlook_search_recipients(
    name: str | None = None,
    query: str | None = None,
    account_email: str | None = None,
) -> dict:
    result = ol.search_recipients(name=name, query=query, account_email=account_email)
    return {"recipients": result, "count": len(result)}


def outlook_search_all_folders(
    query: str | None = None,
    subject: str | None = None,
    sender: str | None = None,
    body_contains: str | None = None,
    top: int | None = None,
    scan_limit: int | None = None,
    timeout_seconds: int = 30,
    account_email: str | None = None,
) -> dict:
    timeout_seconds = max(1, min(timeout_seconds, 300))
    detailed_kwargs = {
        "query": query,
        "subject": subject,
        "sender": sender,
        "body_contains": body_contains,
        "top": top,
        "timeout_seconds": timeout_seconds,
        "account_email": account_email,
    }
    if scan_limit is not None:
        detailed_kwargs["scan_limit"] = scan_limit
    result = ol.search_all_folders_detailed(**detailed_kwargs)
    return {
        "messages": result["messages"],
        "count": len(result["messages"]),
        "errors": result["errors"],
        "partial_results": result.get("partial_results", False),
        "folders_searched": result.get("folders_searched", 0),
        "folders_total": result.get("folders_total", 0),
        "any_folder_capped": result.get("any_folder_capped", False),
    }
