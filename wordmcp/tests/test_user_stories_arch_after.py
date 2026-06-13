"""
User Story Tests — Word MCP ARCH Design Sprint (AFTER delivery) — AS01–AS10

Behavioural tests exercising the contracts delivered in P1-P8:
  P2: export_document format Literal + validation
  P3: manage_tracked_changes no accept_one / reject_one on MCP surface
  P4: manage_comments CRUD with confirm gate
  P5: bulk_update_paragraphs max=500
  P6: bulk_update_table_cells max=200; write + confirm gates
  P7: format/operation routing (structural gating via code paths, not grep)
  P8: Exactly 39 tools registered

These tests are BEHAVIOURAL — they import and call real code paths.
All marked @pytest.mark.unit; no COM / win32com required.

Sprint action: (internal reference removed)
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pytest
from fastmcp.exceptions import ToolError

# Ensure word/src is on sys.path before any wordmcp imports.
WORD_SRC = Path(__file__).parents[1] / "src"
sys.path.insert(0, str(WORD_SRC))


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------


def _noop_bind(fn: Any, name: str) -> Any:  # noqa: D401
    """Passthrough bind_tool — returns the inner handler function unwrapped."""
    return fn


def _make_runtime() -> Any:
    """Construct a ServerRuntime backed by document_docx with no COM support."""
    from wordmcp._server_runtime import ServerRuntime
    from wordmcp import document_docx
    from wordmcp.document_docx import WordMCPError

    def _require_com_stub() -> None:
        raise ToolError("COM not available in unit tests")

    return ServerRuntime(
        get_doc=lambda: document_docx,
        get_word_error=lambda: WordMCPError,
        get_com=lambda: None,
        get_com_loaded=lambda: False,
        require_com=_require_com_stub,
    )


# ---------------------------------------------------------------------------
# AS01 — export_document rejects unknown format (P2)
# ---------------------------------------------------------------------------


@pytest.mark.unit
def test_as01_export_document_rejects_unknown_format() -> None:
    """AS a Word MCP user, I WANT export_document to reject unsupported formats like 'docx'
    SO THAT I get an immediate ToolError rather than confusing downstream failure."""
    from wordmcp._server_handlers_docx_export import register

    tools = register(_noop_bind, _make_runtime())
    export_fn = tools["export_document"]

    with pytest.raises(ToolError) as exc_info:
        export_fn(path="test.docx", format="docx")

    # The error message must name the offending value so the caller can self-correct.
    assert "docx" in str(exc_info.value)


# ---------------------------------------------------------------------------
# AS02 — bulk_update_paragraphs rejects empty updates list (P5)
# ---------------------------------------------------------------------------


@pytest.mark.unit
def test_as02_bulk_update_paragraphs_rejects_empty_updates(monkeypatch: pytest.MonkeyPatch) -> None:
    """AS a Word MCP user, I WANT bulk_update_paragraphs to reject an empty updates list
    SO THAT I get a clear diagnostics error rather than a silent no-op."""
    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    from wordmcp import _server_docx_tools

    runtime = _make_runtime()
    with pytest.raises(ToolError) as exc_info:
        _server_docx_tools.bulk_update_paragraphs(
            runtime,
            path="irrelevant.docx",
            updates=[],
            confirm=True,
        )
    assert "empty" in str(exc_info.value).lower()


# ---------------------------------------------------------------------------
# AS03 — bulk_update_paragraphs rejects batch > 500 (P5)
# ---------------------------------------------------------------------------


@pytest.mark.unit
def test_as03_bulk_update_paragraphs_rejects_oversized_batch(monkeypatch: pytest.MonkeyPatch) -> None:
    """AS a Word MCP user, I WANT bulk_update_paragraphs to enforce a max of 500 per call
    SO THAT the server is protected from accidental very large mutations in a single request."""
    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    from wordmcp import _server_docx_tools

    runtime = _make_runtime()
    big = [{"paragraph_index": i, "new_text": "x"} for i in range(501)]
    with pytest.raises(ToolError) as exc_info:
        _server_docx_tools.bulk_update_paragraphs(
            runtime,
            path="irrelevant.docx",
            updates=big,
            confirm=True,
        )
    assert "500" in str(exc_info.value)


# ---------------------------------------------------------------------------
# AS04 — bulk_update_table_cells rejects batch > 200 (P6)
# ---------------------------------------------------------------------------


@pytest.mark.unit
def test_as04_bulk_update_table_cells_rejects_oversized_batch(monkeypatch: pytest.MonkeyPatch) -> None:
    """AS a Word MCP user, I WANT bulk_update_table_cells to enforce a max of 200 per call
    SO THAT runaway table-fill operations are bounded at the MCP contract level."""
    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    from wordmcp import _server_docx_tools

    runtime = _make_runtime()
    big = [{"table_index": 0, "row": 0, "col": 0, "new_text": "x"} for _ in range(201)]
    with pytest.raises(ToolError) as exc_info:
        _server_docx_tools.bulk_update_table_cells(
            runtime,
            path="irrelevant.docx",
            updates=big,
            confirm=True,
        )
    assert "200" in str(exc_info.value)


# ---------------------------------------------------------------------------
# AS05 — manage_comments resolve requires confirm=True (P4/P6)
# ---------------------------------------------------------------------------


@pytest.mark.unit
def test_as05_manage_comments_resolve_requires_confirm(monkeypatch: pytest.MonkeyPatch) -> None:
    """AS a Word MCP user, I WANT manage_comments resolve to require confirm=True
    SO THAT comments cannot be resolved without an explicit confirmation flag."""
    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    from wordmcp import _server_docx_tools

    runtime = _make_runtime()
    with pytest.raises(ToolError) as exc_info:
        _server_docx_tools.manage_comments(
            runtime,
            path="irrelevant.docx",
            operation="resolve",
            paragraph_index=0,
            confirm=False,           # deliberately omitted
        )
    assert "confirm" in str(exc_info.value).lower()


# ---------------------------------------------------------------------------
# AS06 — manage_comments delete requires confirm=True (P4/P6)
# ---------------------------------------------------------------------------


@pytest.mark.unit
def test_as06_manage_comments_delete_requires_confirm(monkeypatch: pytest.MonkeyPatch) -> None:
    """AS a Word MCP user, I WANT manage_comments delete to require confirm=True
    SO THAT comments cannot be permanently deleted without an explicit safeguard."""
    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    from wordmcp import _server_docx_tools

    runtime = _make_runtime()
    with pytest.raises(ToolError) as exc_info:
        _server_docx_tools.manage_comments(
            runtime,
            path="irrelevant.docx",
            operation="delete",
            paragraph_index=0,
            confirm=False,
        )
    assert "confirm" in str(exc_info.value).lower()


# ---------------------------------------------------------------------------
# AS07 — manage_tracked_changes rejects unknown operation like "accept_one" (P3)
# ---------------------------------------------------------------------------


@pytest.mark.unit
def test_as07_manage_tracked_changes_rejects_accept_one() -> None:
    """AS a Word MCP user, I WANT manage_tracked_changes to raise a ToolError for 'accept_one'
    SO THAT per-revision operations not on the MCP surface fail fast with a clear message."""
    from wordmcp._server_handlers_com import register

    tools = register(_noop_bind, _make_runtime())
    handler = tools["manage_tracked_changes"]

    with pytest.raises(ToolError) as exc_info:
        handler(path="test.docx", operation="accept_one")

    # Error must name the offending operation so the caller knows what went wrong.
    assert "accept_one" in str(exc_info.value)


# ---------------------------------------------------------------------------
# AS08 — export_document format="txt" routes to docx path, NOT COM (P2)
# ---------------------------------------------------------------------------


@pytest.mark.unit
def test_as08_export_txt_uses_docx_path_not_com(monkeypatch: pytest.MonkeyPatch) -> None:
    """AS a Word MCP user, I WANT export_document with format='txt' to use the python-docx path
    SO THAT plain-text export works without COM / Word desktop being installed."""
    import wordmcp._server_docx_tools as docx_tools
    import wordmcp._server_com_tools as com_tools
    from wordmcp._server_handlers_docx_export import register

    DOCX_SENTINEL: dict[str, Any] = {"text": "_txt_from_docx_path_", "chars": 20}

    def _assert_not_called(*a: Any, **kw: Any) -> None:
        raise AssertionError("COM export_document must NOT be called for format='txt'")

    monkeypatch.setattr(docx_tools, "export_as_text", lambda rt, **kw: DOCX_SENTINEL)
    monkeypatch.setattr(com_tools, "export_document", _assert_not_called)

    tools = register(_noop_bind, _make_runtime())
    result = tools["export_document"](path="any.docx", format="txt")

    assert result is DOCX_SENTINEL, "Expected docx-path sentinel, got a different value"


# ---------------------------------------------------------------------------
# AS09 — export_document format="md" routes to docx path, NOT COM (P2)
# ---------------------------------------------------------------------------


@pytest.mark.unit
def test_as09_export_md_uses_docx_path_not_com(monkeypatch: pytest.MonkeyPatch) -> None:
    """AS a Word MCP user, I WANT export_document with format='md' to use the python-docx path
    SO THAT Markdown export works without COM / Word desktop being installed."""
    import wordmcp._server_docx_tools as docx_tools
    import wordmcp._server_com_tools as com_tools
    from wordmcp._server_handlers_docx_export import register

    MD_SENTINEL: dict[str, Any] = {"markdown": "# _md_from_docx_path_", "chars": 21}

    def _assert_not_called(*a: Any, **kw: Any) -> None:
        raise AssertionError("COM export_document must NOT be called for format='md'")

    monkeypatch.setattr(docx_tools, "export_to_markdown", lambda rt, **kw: MD_SENTINEL)
    monkeypatch.setattr(com_tools, "export_document", _assert_not_called)

    tools = register(_noop_bind, _make_runtime())
    result = tools["export_document"](path="any.docx", format="md")

    assert result is MD_SENTINEL, "Expected docx-path sentinel, got a different value"


# ---------------------------------------------------------------------------
# AS10 — Server registers exactly 40 tools in the core handlers (P8)
# Note: paragraph dispatcher and context handler add further tools on top.
# ---------------------------------------------------------------------------

_EXPECTED_TOOL_COUNT = 40


@pytest.mark.unit
def test_as10_server_registers_exactly_39_tools() -> None:
    """AS a Word MCP user, I WANT the Word MCP surface to expose at most 40 core tools
    in the primary handlers SO THAT the tool count fits model context windows.
    (Updated from 39 to 40 to include set_paragraph_format; read_document moved
    from _server_handlers_docx_core to _server_handlers_read_document dispatcher.)"""
    from wordmcp import (
        _server_handlers_com,
        _server_handlers_docx_advanced,
        _server_handlers_docx_core,
        _server_handlers_docx_edit,
        _server_handlers_docx_export,
        _server_handlers_meta,
        _server_handlers_read_document,
        _server_handlers_review,
    )

    runtime = _make_runtime()
    all_tools: dict[str, Any] = {}
    for module in (
        _server_handlers_meta,
        _server_handlers_read_document,
        _server_handlers_docx_core,
        _server_handlers_docx_edit,
        _server_handlers_docx_advanced,
        _server_handlers_com,
        _server_handlers_docx_export,
        _server_handlers_review,
    ):
        all_tools.update(module.register(_noop_bind, runtime))

    count = len(all_tools)
    assert count == _EXPECTED_TOOL_COUNT, (
        f"Expected exactly {_EXPECTED_TOOL_COUNT} registered tools, got {count}. "
        f"Tools found: {sorted(all_tools)}"
    )
