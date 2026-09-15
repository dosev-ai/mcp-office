"""Deterministic scheduler tests; no threads or Outlook calls are started."""
from __future__ import annotations

import sys
from concurrent.futures import Future
from types import ModuleType
from unittest.mock import Mock

import pytest

import mailmcp
from mailmcp import _core, _messages


class FakeExecutor:
    def __init__(self, results):
        self.results = results
        self.futures = []
        self.shutdown = Mock()

    def submit(self, worker, folder, *args, **kwargs):
        future = Future()
        value = self.results.get(folder)
        if isinstance(value, Exception):
            future.set_exception(value)
        elif value is not None:
            future.set_result(value)
        self.futures.append(future)
        return future


def scheduler(monkeypatch, results):
    executor = FakeExecutor(results)
    monkeypatch.setattr(_messages, "ThreadPoolExecutor", lambda **kwargs: executor)
    shim = ModuleType("mailmcp.outlook_com")
    shim.search_all_folders_worker = Mock()
    monkeypatch.setitem(sys.modules, "mailmcp.outlook_com", shim)
    monkeypatch.setattr(mailmcp, "outlook_com", shim, raising=False)
    _core.set_config(_core.OutlookConfig(allowlist_folders=list(results)))
    return executor


def test_global_top_waits_for_later_folder(monkeypatch):
    scheduler(monkeypatch, {
        "Inbox": [{"entry_id": "old", "received_time": "2026-01-01", "folder_name": "Inbox"}],
        "Archive": [{"entry_id": "new", "received_time": "2026-02-01", "folder_name": "Archive"}],
    })
    monkeypatch.setattr(_messages, "as_completed", lambda futures, **kwargs: iter(futures))
    result = _messages.search_all_folders_detailed(subject="synthetic", top=1)
    assert [row["entry_id"] for row in result["messages"]] == ["new"]
    assert result["folders_searched"] == 2
    assert result["partial_results"] is False


def test_wait_deadline_cancels_pending_without_blocking(monkeypatch):
    executor = scheduler(monkeypatch, {"Inbox": None})
    def timed_out(futures, timeout=None):
        assert timeout is not None and 0 <= timeout <= 0.5
        raise TimeoutError("Synthetic wait deadline")
    monkeypatch.setattr(_messages, "as_completed", timed_out)
    result = _messages.search_all_folders_detailed(subject="synthetic", timeout_seconds=0.5)
    assert result["partial_results"] is True
    assert result["folders_searched"] == 0
    assert result["messages"] == []
    assert executor.futures[0].cancelled()
    executor.shutdown.assert_called_once_with(wait=False, cancel_futures=True)
    _core._get_outlook.assert_not_called()


@pytest.mark.parametrize("timeout", [0, -1, True, float("nan"), float("inf"), "30"])
def test_invalid_deadline_rejected_before_executor(monkeypatch, timeout):
    factory = Mock(side_effect=AssertionError("Must not start executor"))
    monkeypatch.setattr(_messages, "ThreadPoolExecutor", factory)
    with pytest.raises(ValueError, match="timeout_seconds"):
        _messages.search_all_folders_detailed(subject="synthetic", timeout_seconds=timeout)
    factory.assert_not_called()


def test_failed_folder_marks_results_partial(monkeypatch):
    scheduler(monkeypatch, {"Inbox": [], "Archive": RuntimeError("synthetic failure")})
    monkeypatch.setattr(_messages, "as_completed", lambda futures, **kwargs: iter(futures))
    result = _messages.search_all_folders_detailed(subject="synthetic")
    assert result["partial_results"] is True
    assert len(result["errors"]) == 1


def test_restrict_failure_does_not_return_unfiltered_messages(monkeypatch, mailbox):
    items = Mock()
    items.Restrict.side_effect = RuntimeError("Unsupported filter")
    mailbox.folders[("a", "Inbox")].Items = items
    monkeypatch.setattr(_messages._folders, "_folder_by_name", lambda name: mailbox.folders[("a", "Inbox")])
    with pytest.raises(RuntimeError, match="filter"):
        _messages.list_messages(unread_only=True)
