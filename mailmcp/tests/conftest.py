"""Synthetic Outlook fixtures. Tests never connect to a real Outlook profile."""
from __future__ import annotations

import os
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from mailmcp import _core


class Recipients:
    def __init__(self, addresses=()):
        self.rows = [self._row(address, 1) for address in addresses]
        self.mutations = []

    @staticmethod
    def _row(address, recipient_type):
        return SimpleNamespace(
            AddressEntry=SimpleNamespace(Address=address),
            Type=recipient_type,
        )

    @property
    def Count(self):
        return len(self.rows)

    def Item(self, index):
        return self.rows[index - 1]

    def Add(self, address):
        self.mutations.append(("add", address))
        row = self._row(address, 1)
        self.rows.append(row)
        return row

    def Remove(self, index):
        self.mutations.append(("remove", index))
        del self.rows[index - 1]

    def ResolveAll(self):
        return True


@pytest.fixture(autouse=True)
def isolated_outlook_policy(monkeypatch):
    for key in os.environ:
        if key.startswith("OUTLOOK_"):
            monkeypatch.delenv(key)
    config = _core.OutlookConfig()
    monkeypatch.setattr(_core, "_config", config)
    monkeypatch.setattr(_core, "_startup_config", config)
    monkeypatch.setattr(_core, "_account_overrides", {})
    monkeypatch.setattr(_core, "_startup_account_overrides", {})
    monkeypatch.setattr(
        _core, "_get_outlook", Mock(side_effect=AssertionError("Real COM is forbidden"))
    )
    _core._invalidate_folder_cache()
    yield
    _core._invalidate_folder_cache()


@pytest.fixture
def mailbox(monkeypatch):
    names = {3: "Deleted Items", 6: "Inbox", 9: "Calendar", 10: "Contacts", 16: "Drafts", 23: "Junk Email"}
    stores = []
    accounts = []
    folders = {}
    for suffix in ("a", "b"):
        store = SimpleNamespace(StoreID=f"store-{suffix}", DisplayName=f"owner-{suffix}@example.com")
        per_store = {}
        for folder_id, name in names.items():
            folder = SimpleNamespace(
                EntryID=f"folder-{suffix}-{folder_id}", StoreID=store.StoreID,
                Store=store, Name=name, Items=SimpleNamespace(),
            )
            per_store[folder_id] = folder
            folders[(suffix, name)] = folder
        store.GetDefaultFolder = lambda number, f=per_store: f[number]
        store.GetRootFolder = lambda f=per_store: SimpleNamespace(Folders=list(f.values()))
        accounts.append(SimpleNamespace(
            SmtpAddress=store.DisplayName, DisplayName=store.DisplayName, DeliveryStore=store,
        ))
        stores.append(store)
    item = SimpleNamespace(
        Class=43, Sent=False, EntryID="synthetic-item", Parent=folders[("a", "Drafts")],
        Subject="Synthetic original", HTMLBody="Original body", Body="Original body",
        UnRead=True, FlagStatus=0, Recipients=Recipients(["person@example.com"]),
        Save=Mock(), Send=Mock(), Delete=Mock(), Move=Mock(),
    )
    item.Move.return_value = item
    mapi = SimpleNamespace(
        GetItemFromID=Mock(return_value=item), GetDefaultFolder=stores[0].GetDefaultFolder,
        Accounts=accounts, Stores=stores, CreateRecipient=Mock(),
    )
    getter = Mock(return_value=mapi)
    monkeypatch.setattr(_core, "_mapi", getter)
    _core.set_config(_core.OutlookConfig(
        allowlist_folders=list(names.values()), enable_write=True, enable_send=True,
        enable_delete=True, allowlist_domains=["example.com"],
    ))
    return SimpleNamespace(item=item, mapi=mapi, get_mapi=getter, folders=folders, stores=stores)
