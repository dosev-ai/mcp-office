from __future__ import annotations

from collections import OrderedDict

import pytest
from fastmcp.exceptions import ToolError

import wordmcp.document_docx as dd


def _call_tool(tool, *args, **kwargs):
    fn = getattr(tool, "fn", tool)
    return fn(*args, **kwargs)


@pytest.mark.unit
def test_facade_exports_compatibility_surface() -> None:
    expected = [
        "WordMCPError",
        "ValidationError",
        "NotAllowedError",
        "NotFoundError",
        "FileError",
        "OfficeCOMError",
        "_check_path",
        "_check_image_path",
        "_check_write",
        "_check_confirm",
        "_load_doc",
        "_atomic_save",
        "_doc_cache",
        "_DOC_CACHE_MAX",
        "_MAX_DOCX_BYTES",
        "read_document",
        "export_as_text",
        "add_paragraph",
        "manage_comments",
        "export_to_markdown",
    ]
    for name in expected:
        assert hasattr(dd, name), name


@pytest.mark.unit
def test_exception_identity_is_canonical() -> None:
    from wordmcp import document_docx as dd_again

    assert dd.WordMCPError is dd_again.WordMCPError
    assert dd.ValidationError is dd_again.ValidationError
    assert dd.NotAllowedError is dd_again.NotAllowedError
    assert dd.NotFoundError is dd_again.NotFoundError
    assert dd.FileError is dd_again.FileError
    assert dd.OfficeCOMError is dd_again.OfficeCOMError


@pytest.mark.unit
def test_cache_singleton_identity_is_preserved() -> None:
    from wordmcp._docx import core

    assert dd._doc_cache is core._doc_cache
    assert isinstance(dd._doc_cache, OrderedDict)


@pytest.mark.unit
def test_document_com_imports_canonical_facade_symbols() -> None:
    import wordmcp.document_com as document_com

    assert document_com.ValidationError is dd.ValidationError
    assert document_com.NotAllowedError is dd.NotAllowedError
    assert document_com.OfficeCOMError is dd.OfficeCOMError
    assert document_com._check_path is dd._check_path
    assert document_com._check_write is dd._check_write
    assert document_com._check_confirm is dd._check_confirm


@pytest.mark.unit
def test_server_catches_canonical_wordmcp_error(monkeypatch) -> None:
    from wordmcp import server as srv

    class Boom(dd.WordMCPError):
        pass

    monkeypatch.setattr(dd, "capabilities", lambda: (_ for _ in ()).throw(Boom("boom")))

    with pytest.raises(ToolError, match="boom"):
        _call_tool(srv.capabilities)


@pytest.mark.unit
def test_capabilities_manifest_is_unchanged() -> None:
    caps = dd.capabilities()
    assert caps["phase"] == "1.0"
    assert caps["backend"] == "python-docx"
    assert len(caps["tools"]) == 14
    assert caps["tools"] == [
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
    ]
