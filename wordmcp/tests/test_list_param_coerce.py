"""Tests for list parameter JSON string coercion in add_list, bulk_add_paragraphs,
bulk_update_paragraphs, and bulk_update_table_cells."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest


def _make_runtime():
    runtime = MagicMock()
    runtime.call_doc = MagicMock(return_value={"status": "ok"})
    return runtime


# ---------------------------------------------------------------------------
# add_list — items parameter
# ---------------------------------------------------------------------------
class TestAddListItemsCoercion:
    def test_list_passes_through(self):
        from wordmcp._server_docx_tools import add_list

        rt = _make_runtime()
        add_list(rt, "test.docx", items=["a", "b"], confirm=True)
        _, kwargs = rt.call_doc.call_args
        assert kwargs["items"] == ["a", "b"]

    def test_json_string_coerced(self):
        from wordmcp._server_docx_tools import add_list

        rt = _make_runtime()
        add_list(rt, "test.docx", items='["a","b"]', confirm=True)
        _, kwargs = rt.call_doc.call_args
        assert kwargs["items"] == ["a", "b"]

    def test_malformed_json_raises(self):
        from fastmcp.exceptions import ToolError
        from wordmcp._server_docx_tools import add_list

        rt = _make_runtime()
        with pytest.raises(ToolError, match="invalid JSON string"):
            add_list(rt, "test.docx", items="not json", confirm=True)

    def test_json_non_array_raises(self):
        from fastmcp.exceptions import ToolError
        from wordmcp._server_docx_tools import add_list

        rt = _make_runtime()
        with pytest.raises(ToolError, match="expected a JSON array"):
            add_list(rt, "test.docx", items='{"a": 1}', confirm=True)


# ---------------------------------------------------------------------------
# bulk_add_paragraphs — paragraphs parameter
# ---------------------------------------------------------------------------
class TestBulkAddParagraphsCoercion:
    def test_list_passes_through(self):
        from wordmcp._server_docx_tools import bulk_add_paragraphs

        rt = _make_runtime()
        paras = [{"text": "hello"}]
        bulk_add_paragraphs(rt, "test.docx", paragraphs=paras, confirm=True)
        _, kwargs = rt.call_doc.call_args
        assert kwargs["paragraphs"] == [{"text": "hello"}]

    def test_json_string_coerced(self):
        from wordmcp._server_docx_tools import bulk_add_paragraphs

        rt = _make_runtime()
        bulk_add_paragraphs(
            rt, "test.docx", paragraphs='[{"text":"hello"}]', confirm=True
        )
        _, kwargs = rt.call_doc.call_args
        assert kwargs["paragraphs"] == [{"text": "hello"}]

    def test_malformed_json_raises(self):
        from fastmcp.exceptions import ToolError
        from wordmcp._server_docx_tools import bulk_add_paragraphs

        rt = _make_runtime()
        with pytest.raises(ToolError, match="invalid JSON string"):
            bulk_add_paragraphs(rt, "test.docx", paragraphs="bad", confirm=True)

    def test_json_non_array_raises(self):
        from fastmcp.exceptions import ToolError
        from wordmcp._server_docx_tools import bulk_add_paragraphs

        rt = _make_runtime()
        with pytest.raises(ToolError, match="expected a JSON array"):
            bulk_add_paragraphs(rt, "test.docx", paragraphs='"hello"', confirm=True)


# ---------------------------------------------------------------------------
# bulk_update_paragraphs — updates parameter
# ---------------------------------------------------------------------------
class TestBulkUpdateParagraphsCoercion:
    def test_list_passes_through(self):
        from wordmcp._server_docx_tools import bulk_update_paragraphs

        rt = _make_runtime()
        updates = [{"paragraph_index": 0, "new_text": "x"}]
        bulk_update_paragraphs(rt, "test.docx", updates=updates, confirm=True)
        _, kwargs = rt.call_doc.call_args
        assert kwargs["updates"] == updates

    def test_json_string_coerced(self):
        from wordmcp._server_docx_tools import bulk_update_paragraphs

        rt = _make_runtime()
        bulk_update_paragraphs(
            rt,
            "test.docx",
            updates='[{"paragraph_index":0,"new_text":"x"}]',
            confirm=True,
        )
        _, kwargs = rt.call_doc.call_args
        assert kwargs["updates"] == [{"paragraph_index": 0, "new_text": "x"}]

    def test_malformed_json_raises(self):
        from fastmcp.exceptions import ToolError
        from wordmcp._server_docx_tools import bulk_update_paragraphs

        rt = _make_runtime()
        with pytest.raises(ToolError, match="invalid JSON string"):
            bulk_update_paragraphs(rt, "test.docx", updates="bad", confirm=True)

    def test_json_non_array_raises(self):
        from fastmcp.exceptions import ToolError
        from wordmcp._server_docx_tools import bulk_update_paragraphs

        rt = _make_runtime()
        with pytest.raises(ToolError, match="expected a JSON array"):
            bulk_update_paragraphs(rt, "test.docx", updates='{"a":1}', confirm=True)


# ---------------------------------------------------------------------------
# bulk_update_table_cells — updates parameter
# ---------------------------------------------------------------------------
class TestBulkUpdateTableCellsCoercion:
    def test_list_passes_through(self):
        from wordmcp._server_docx_tools import bulk_update_table_cells

        rt = _make_runtime()
        updates = [{"table_index": 0, "row": 0, "col": 0, "new_text": "x"}]
        bulk_update_table_cells(rt, "test.docx", updates=updates, confirm=True)
        _, kwargs = rt.call_doc.call_args
        assert kwargs["updates"] == updates

    def test_json_string_coerced(self):
        from wordmcp._server_docx_tools import bulk_update_table_cells

        rt = _make_runtime()
        bulk_update_table_cells(
            rt,
            "test.docx",
            updates='[{"table_index":0,"row":0,"col":0,"new_text":"x"}]',
            confirm=True,
        )
        _, kwargs = rt.call_doc.call_args
        assert kwargs["updates"] == [
            {"table_index": 0, "row": 0, "col": 0, "new_text": "x"}
        ]

    def test_malformed_json_raises(self):
        from fastmcp.exceptions import ToolError
        from wordmcp._server_docx_tools import bulk_update_table_cells

        rt = _make_runtime()
        with pytest.raises(ToolError, match="invalid JSON string"):
            bulk_update_table_cells(rt, "test.docx", updates="bad", confirm=True)

    def test_json_non_array_raises(self):
        from fastmcp.exceptions import ToolError
        from wordmcp._server_docx_tools import bulk_update_table_cells

        rt = _make_runtime()
        with pytest.raises(ToolError, match="expected a JSON array"):
            bulk_update_table_cells(rt, "test.docx", updates='{"a":1}', confirm=True)
