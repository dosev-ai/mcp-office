from __future__ import annotations

import pytest
from fastmcp.exceptions import ToolError


def _call_tool(tool, *args, **kwargs):
    fn = getattr(tool, "fn", tool)
    return fn(*args, **kwargs)


@pytest.mark.unit
def test_server_exports_expected_public_contract() -> None:
    from wordmcp import server as srv

    expected_symbols = [
        "mcp",
        "main",
        "capabilities",
        "read_document",
        "get_document_metadata",
        "list_paragraphs",
        "read_paragraph",
        "list_tables",
        "read_table",
        "add_paragraph",
        "add_heading",
        "add_page_break",
        "add_table",
        "insert_image",
        "find_replace",
        "save",
        "manage_comments",
        "create_document",
        "list_headings",
        "read_section",
        "search_text",
        "list_styles",
        "manage_tracked_changes",
        "export_document",
        "apply_style",
        "update_paragraph",
        "delete_paragraph",
        "update_table_cell",
        "set_document_properties",
        "add_list",
        "insert_paragraph",
        "bulk_add_paragraphs",
        "bulk_update_paragraphs",
        "bulk_update_table_cells",
        "get_headers_footers",
        "add_hyperlink",
        "add_footnote",
    ]

    for name in expected_symbols:
        assert hasattr(srv, name), name


@pytest.mark.unit
def test_capabilities_uses_live_doc_lookup(monkeypatch) -> None:
    import wordmcp.document_docx as dd
    from wordmcp import server as srv

    monkeypatch.setattr(
        dd,
        "capabilities",
        lambda: {
            "phase": "1.0",
            "backend": "python-docx",
            "tools": ["sentinel"],
            "governance": {},
        },
    )

    result = _call_tool(srv.capabilities)

    assert result["tools"][0] == "sentinel"


@pytest.mark.unit
def test_capabilities_uses_live_com_loaded_lookup(monkeypatch) -> None:
    from wordmcp import server as srv

    monkeypatch.setattr(srv, "_COM_LOADED", False)

    result = _call_tool(srv.capabilities)

    assert result["com_tools"]["loaded"] is False


@pytest.mark.unit
def test_manage_tracked_changes_uses_live_com_backend_lookup(monkeypatch) -> None:
    from wordmcp import server as srv

    calls: list[dict] = []

    class StubCom:
        @staticmethod
        def manage_tracked_changes(**kwargs):
            calls.append(kwargs)
            return {"status": "ok", "kwargs": kwargs}

    monkeypatch.setattr(srv, "_COM_LOADED", True)
    monkeypatch.setattr(srv, "_com", StubCom)

    result = _call_tool(
        srv.manage_tracked_changes,
        "fake.docx",
        operation="list",
    )

    assert result["status"] == "ok"
    assert calls == [{"path": "fake.docx", "operation": "list", "revision_index": None, "confirm": False}]


@pytest.mark.unit
def test_export_as_pdf_uses_live_com_loaded_guard(monkeypatch) -> None:
    from wordmcp import server as srv

    monkeypatch.setattr(srv, "_COM_LOADED", False)

    with pytest.raises(ToolError, match="COM tools require pywin32"):
        _call_tool(srv.export_document, "fake.docx", output_path="fake.pdf", format="pdf", confirm=True)
