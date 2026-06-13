from __future__ import annotations

import pytest
from fastmcp.exceptions import ToolError


def _call_tool(tool, *args, **kwargs):
    fn = getattr(tool, "fn", tool)
    return fn(*args, **kwargs)


@pytest.mark.unit
def test_server_entrypoint_preserves_live_doc_lookup(monkeypatch) -> None:
    import wordmcp.document_docx as dd
    from wordmcp import server as srv

    monkeypatch.setattr(dd, "read_document", lambda path: {"path": path, "sentinel": True})

    result = _call_tool(srv.read_document, "fake.docx")

    assert result == {"path": "fake.docx", "sentinel": True}


@pytest.mark.unit
def test_server_entrypoint_preserves_live_com_backend_lookup(monkeypatch) -> None:
    from wordmcp import server as srv

    calls: list[dict] = []

    class StubCom:
        @staticmethod
        def export_document(**kwargs):
            calls.append(kwargs)
            return {"status": "ok", "kwargs": kwargs}

    monkeypatch.setattr(srv, "_COM_LOADED", True)
    monkeypatch.setattr(srv, "_com", StubCom)

    result = _call_tool(
        srv.export_document,
        "fake.docx",
        output_path="fake.pdf",
        format="pdf",
        confirm=True,
    )

    assert result["status"] == "ok"
    assert calls == [{"path": "fake.docx", "output_path": "fake.pdf", "format": "pdf", "confirm": True}]


@pytest.mark.unit
def test_server_entrypoint_preserves_live_com_guard(monkeypatch) -> None:
    from wordmcp import server as srv

    monkeypatch.setattr(srv, "_COM_LOADED", False)

    with pytest.raises(ToolError, match="COM tools require pywin32"):
        _call_tool(srv.export_document, "fake.docx", output_path="fake.pdf", format="pdf", confirm=True)
