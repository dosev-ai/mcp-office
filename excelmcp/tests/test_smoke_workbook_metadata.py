import asyncio

import openpyxl
import pytest

from excelmcp import server

from ._smoke_helpers import _run_tool


@pytest.mark.smoke
def test_smoke_workbook_metadata_tool_round_trip(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))

    path = tmp_path / "smoke_metadata.xlsx"
    openpyxl.Workbook().save(path)

    write_result = _run_tool(server.mcp, "workbook_metadata", {
        "operation": "write",
        "path": str(path),
        "payload": {"suite": "smoke"},
        "confirm": True,
    })
    read_result = _run_tool(server.mcp, "workbook_metadata", {
        "operation": "read",
        "path": str(path),
    })

    assert write_result["ok"] is True
    assert write_result["created"] is True
    assert read_result["present"] is True
    assert read_result["metadata"] == {"suite": "smoke"}


@pytest.mark.smoke
def test_smoke_workbook_metadata_resource_template_registered():
    templates = asyncio.run(server.mcp.list_resource_templates())
    uris = {getattr(t, "uri_template", getattr(t, "uriTemplate", "")) for t in templates}
    assert any(str(uri).endswith("/metadata") for uri in uris)
