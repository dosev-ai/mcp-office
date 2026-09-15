"""Read-only message and folder operations."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import timedelta
import logging
import math
import os
import time

from mailmcp import _folders
from mailmcp._core import get_config, get_effective_config, _assert_allowed
from mailmcp._formatters import _msg_header, _parse_date, _format_outlook_date

from mailmcp._message_fetch import health, list_accounts, list_folders, get_message  # noqa: F401
from mailmcp._message_save import save_attachments  # noqa: F401
from mailmcp._message_recipients import search_recipients  # noqa: F401

logger = logging.getLogger(__name__)


def _resolve_top_limit(top: int | None, max_items: int) -> int:
    if top is None:
        return max_items
    if isinstance(top, bool) or not isinstance(top, int) or top <= 0:
        raise ValueError("top must be a positive integer or None.")
    return min(top, max_items)


def _resolve_scan_limit(scan_limit: int | None) -> int:
    try:
        env_cap = int(os.environ.get("OUTLOOK_SEARCH_SCAN_LIMIT", "500"))
    except (ValueError, TypeError):
        env_cap = 500
    if scan_limit is None:
        if env_cap == 0:
            return 0
        return max(50, min(env_cap, 5000))
    if isinstance(scan_limit, bool) or not isinstance(scan_limit, int):
        raise ValueError("scan_limit must be a positive integer or None.")
    if scan_limit <= 0:
        raise ValueError("scan_limit must be a positive integer or None. To disable the scan cap set OUTLOOK_SEARCH_SCAN_LIMIT=0.")
    if env_cap == 0:
        return max(50, scan_limit)
    return max(50, min(scan_limit, env_cap, 5000))


def _search_all_folders_max_workers(folder_count: int) -> int:
    if folder_count <= 1:
        return 1
    try:
        env_workers = int(os.environ.get("OUTLOOK_SEARCH_WORKERS", "4"))
    except (ValueError, TypeError):
        env_workers = 4
    env_workers = max(1, min(env_workers, 16))
    return min(folder_count, env_workers)


def _merge_search_all_folders_results(folder_results: list[list[dict]], allowlist_folders: list[str], limit: int) -> list[dict]:
    merged = [msg for folder_result in folder_results for msg in folder_result]
    folder_order = {folder_name: index for index, folder_name in enumerate(allowlist_folders)}
    merged.sort(key=lambda msg: msg.get("entry_id") or "")
    merged.sort(key=lambda msg: folder_order.get(str(msg.get("folder_name") or ""), len(folder_order)))
    merged.sort(key=lambda msg: msg.get("received_time") or "", reverse=True)
    seen: set[str] = set()
    deduped: list[dict] = []
    for msg in merged:
        entry_id = msg.get("entry_id")
        if entry_id is not None:
            if entry_id in seen:
                continue
            seen.add(entry_id)
        deduped.append(msg)
        if len(deduped) >= limit:
            break
    return deduped


def list_messages(folder_name: str = "Inbox", top: int | None = None, unread_only: bool = False, since: str | None = None, until: str | None = None, has_attachments: bool | None = None, sort_desc: bool = True, account_email: str | None = None) -> list[dict]:
    _assert_allowed(folder_name, account_email)
    cfg = get_effective_config(account_email)
    limit = _resolve_top_limit(top, cfg.max_items)
    folder = _folders._folder_by_name_for_account(folder_name, account_email=account_email) if account_email is not None else _folders._folder_by_name(folder_name)
    items = folder.Items
    items.Sort("[ReceivedTime]", sort_desc)
    filters = []
    if unread_only:
        filters.append("[UnRead] = True")
    if since:
        dt, _date_only = _parse_date(since)
        filters.append(f"[ReceivedTime] >= '{_format_outlook_date(dt)}'")
    if until:
        dt, date_only = _parse_date(until)
        if date_only:
            filters.append(f"[ReceivedTime] < '{_format_outlook_date(dt + timedelta(days=1))}'")
        else:
            filters.append(f"[ReceivedTime] <= '{_format_outlook_date(dt)}'")
    if has_attachments is True:
        filters.append("[Attachments].Count > 0")
    elif has_attachments is False:
        filters.append("[Attachments].Count = 0")
    if filters:
        restriction = " AND ".join(f"({f})" for f in filters)
        try:
            items = items.Restrict(restriction)
            items.Sort("[ReceivedTime]", sort_desc)
        except Exception as exc:
            raise RuntimeError("Outlook could not apply the requested message filter; no unfiltered results returned.") from exc
    result = []
    count = 0
    for msg in items:
        if count >= limit:
            break
        try:
            result.append(_msg_header(msg, account_email=account_email))
        except Exception as exc:
            logger.debug("Skipping message: %s", exc)
        count += 1
    return result


def search_messages(folder_name: str = "Inbox", subject: str | None = None, sender: str | None = None, body_contains: str | None = None, top: int | None = None, account_email: str | None = None, scan_limit: int | None = None, include_scan_metadata: bool = False) -> list[dict] | dict[str, object]:
    _assert_allowed(folder_name, account_email)
    cfg = get_effective_config(account_email)
    limit = _resolve_top_limit(top, cfg.max_items)
    resolved_scan_limit = _resolve_scan_limit(scan_limit)
    folder = _folders._folder_by_name_for_account(folder_name, account_email=account_email) if account_email is not None else _folders._folder_by_name(folder_name)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    result = []
    match_count = 0
    inspected_count = 0
    scan_capped = False
    for msg in items:
        if match_count >= limit:
            break
        if resolved_scan_limit > 0 and inspected_count >= resolved_scan_limit:
            scan_capped = True
            break
        inspected_count += 1
        try:
            hdr = _msg_header(msg, account_email=account_email)
            if subject and subject.lower() not in (hdr.get("subject") or "").lower():
                continue
            if sender and sender.lower() not in (hdr.get("sender_email") or "").lower() and sender.lower() not in (hdr.get("sender_name") or "").lower():
                continue
            if body_contains:
                body = (getattr(msg, "Body", "") or "").lower()
                if body_contains.lower() not in body:
                    continue
            result.append(hdr)
            match_count += 1
        except Exception as exc:
            logger.debug("Skipping message in search: %s", exc)
    if include_scan_metadata:
        return {"messages": result, "scan_capped": scan_capped}
    return result


def search_all_folders(subject: str | None = None, sender: str | None = None, body_contains: str | None = None, query: str | None = None, top: int | None = None, scan_limit: int | None = None) -> list[dict]:
    return search_all_folders_detailed(subject=subject, sender=sender, body_contains=body_contains, query=query, top=top, scan_limit=scan_limit)["messages"]


def search_all_folders_detailed(subject: str | None = None, sender: str | None = None, body_contains: str | None = None, query: str | None = None, top: int | None = None, scan_limit: int | None = None, timeout_seconds: int = 30) -> dict:
    if isinstance(timeout_seconds, bool) or not isinstance(timeout_seconds, (int, float)) or not math.isfinite(timeout_seconds) or timeout_seconds <= 0:
        raise ValueError("timeout_seconds must be a positive finite number.")
    cfg = get_config()
    limit = _resolve_top_limit(top, cfg.max_items)
    if query and not any([subject, sender, body_contains]):
        body_contains = query
    if not any([subject, sender, body_contains]):
        raise ValueError("At least one search filter (query, subject, sender, body_contains) must be provided.")
    folder_results: list[list[dict]] = []
    folder_errors: list[dict[str, str]] = []
    folder_order = {folder_name: index for index, folder_name in enumerate(cfg.allowlist_folders)}
    max_workers = _search_all_folders_max_workers(len(cfg.allowlist_folders))
    partial = False
    any_folder_capped = False
    start_time = time.monotonic()
    folders_searched = 0
    from mailmcp import outlook_com as ol
    executor = ThreadPoolExecutor(max_workers=max_workers)
    future_to_folder = {}
    try:
        future_to_folder = {
            executor.submit(ol.search_all_folders_worker, folder_name, subject, sender, body_contains, limit, scan_limit=scan_limit): folder_name
            for folder_name in cfg.allowlist_folders
        }
        for future in as_completed(future_to_folder, timeout=max(0.0, timeout_seconds - (time.monotonic() - start_time))):
            folder_name = future_to_folder[future]
            try:
                results = future.result(timeout=0)
                folder_results.append(results)
                if getattr(results, "scan_capped", False):
                    any_folder_capped = True
            except Exception as exc:
                folder_errors.append({"folder_name": folder_name, "error": str(exc)})
            folders_searched += 1
    except TimeoutError:
        partial = True
    finally:
        for pending in future_to_folder:
            if not pending.done():
                pending.cancel()
        executor.shutdown(wait=False, cancel_futures=True)
    partial = partial or bool(folder_errors) or any_folder_capped or folders_searched < len(future_to_folder)
    folder_errors.sort(key=lambda item: folder_order.get(item["folder_name"], len(folder_order)))
    if folder_errors and len(folder_errors) == len(cfg.allowlist_folders):
        detail = "; ".join(f"{item['folder_name']}: {item['error']}" for item in folder_errors)
        raise RuntimeError(f"search_all_folders failed for all folders: {detail}")
    return {
        "messages": _merge_search_all_folders_results(folder_results, cfg.allowlist_folders, limit),
        "errors": folder_errors,
        "partial_results": partial,
        "any_folder_capped": any_folder_capped,
        "folders_searched": folders_searched,
        "folders_total": len(future_to_folder),
    }
