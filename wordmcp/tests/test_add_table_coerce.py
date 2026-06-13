"""Tests for add_table data parameter JSON string coercion."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest


def _make_runtime():
    runtime = MagicMock()
    runtime.call_doc = MagicMock(return_value={"status": "ok"})
    return runtime


class TestAddTableDataCoercion:
    def test_list_of_lists_passes_through(self):
        from wordmcp._server_docx_tools import add_table

        rt = _make_runtime()
        data = [["a", "b"], ["c", "d"]]
        add_table(rt, "test.docx", 2, 2, data=data, confirm=True)
        _, kwargs = rt.call_doc.call_args
        assert kwargs["data"] == [["a", "b"], ["c", "d"]]

    def test_json_string_coerced_to_list(self):
        from wordmcp._server_docx_tools import add_table

        rt = _make_runtime()
        add_table(rt, "test.docx", 2, 2, data='[["a","b"],["c","d"]]', confirm=True)
        _, kwargs = rt.call_doc.call_args
        assert kwargs["data"] == [["a", "b"], ["c", "d"]]

    def test_malformed_json_raises_tool_error(self):
        from fastmcp.exceptions import ToolError

        from wordmcp._server_docx_tools import add_table

        rt = _make_runtime()
        with pytest.raises(ToolError, match="invalid JSON string"):
            add_table(rt, "test.docx", 2, 2, data="not json", confirm=True)

    def test_json_non_array_raises_tool_error(self):
        from fastmcp.exceptions import ToolError

        from wordmcp._server_docx_tools import add_table

        rt = _make_runtime()
        with pytest.raises(ToolError, match="expected a JSON array"):
            add_table(rt, "test.docx", 2, 2, data='{"a": 1}', confirm=True)

    def test_none_passes_through(self):
        from wordmcp._server_docx_tools import add_table

        rt = _make_runtime()
        add_table(rt, "test.docx", 2, 2, data=None, confirm=True)
        _, kwargs = rt.call_doc.call_args
        assert kwargs["data"] is None

    def test_flat_json_array_raises_tool_error(self):
        from fastmcp.exceptions import ToolError

        from wordmcp._server_docx_tools import add_table

        rt = _make_runtime()
        with pytest.raises(ToolError, match=r"data\[0\]:.*expected a JSON array"):
            add_table(rt, "test.docx", 2, 1, data='["a","b"]', confirm=True)

    def test_dict_in_array_raises_tool_error(self):
        from fastmcp.exceptions import ToolError

        from wordmcp._server_docx_tools import add_table

        rt = _make_runtime()
        with pytest.raises(ToolError, match=r"data\[0\]:.*expected a JSON array"):
            add_table(rt, "test.docx", 1, 2, data='[{"a":1,"b":2}]', confirm=True)

    def test_empty_string_raises_tool_error(self):
        from fastmcp.exceptions import ToolError

        from wordmcp._server_docx_tools import add_table

        rt = _make_runtime()
        with pytest.raises(ToolError, match="invalid JSON string"):
            add_table(rt, "test.docx", 2, 2, data="", confirm=True)

    def test_whitespace_string_raises_tool_error(self):
        from fastmcp.exceptions import ToolError

        from wordmcp._server_docx_tools import add_table

        rt = _make_runtime()
        with pytest.raises(ToolError, match="invalid JSON string"):
            add_table(rt, "test.docx", 2, 2, data="   ", confirm=True)
