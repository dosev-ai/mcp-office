from __future__ import annotations

from collections.abc import Callable
from typing import Any

from wordmcp import _server_docx_tools
from wordmcp._server_runtime import ServerRuntime


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _add_list(
        path: str,
        items: list | str,
        list_type: str = "bullet",
        level: int = 0,
        confirm: bool = False,
    ) -> dict:
        """
        Append a bulleted or numbered list to the document.
        list_type: 'bullet' (default) | 'number'.
        items: list of string values. Max 500 items per call.
        level: indent level (0-8). 0 = top level, 1+ = nested. Default 0.
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        """
        return _server_docx_tools.add_list(
            runtime,
            path=path,
            items=items,
            list_type=list_type,
            level=level,
            confirm=confirm,
        )

    def _insert_paragraph(
        path: str,
        paragraph_index: int,
        text: str,
        style: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """
        Insert a new paragraph at a specific position in the document body.
        paragraph_index=0 inserts before the first paragraph.
        paragraph_index=len(paragraphs) appends to the end.
        style: optional Word paragraph style name (e.g. 'Heading 1', 'Normal').
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        """
        return _server_docx_tools.insert_paragraph(
            runtime,
            path=path,
            paragraph_index=paragraph_index,
            text=text,
            style=style,
            confirm=confirm,
        )

    def _bulk_add_paragraphs(
        path: str,
        paragraphs: list | str,
        confirm: bool = False,
    ) -> dict:
        """
        Add multiple paragraphs to the document in a single operation (one save).
        paragraphs: list of dicts with keys: text (str), style (optional str). Max 500.
        More efficient than calling add_paragraph repeatedly.
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        """
        return _server_docx_tools.bulk_add_paragraphs(
            runtime,
            path=path,
            paragraphs=paragraphs,
            confirm=confirm,
        )

    def _get_headers_footers(path: str) -> list:
        """
        Return the header and footer text for each document section.
        Includes whether each is linked to the previous section.
        Returns a list of {section_index, header_text, footer_text,
        header_linked_to_previous, footer_linked_to_previous}.
        """
        return _server_docx_tools.get_headers_footers(runtime, path=path)

    def _add_hyperlink(
        path: str,
        paragraph_index: int,
        text: str,
        url: str,
        confirm: bool = False,
    ) -> dict:
        """
        Add a clickable hyperlink to a paragraph.
        url must begin with http:// or https:// (javascript: and file:// schemes are blocked).
        text: the display text (defaults to url if empty).
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        """
        return _server_docx_tools.add_hyperlink(
            runtime,
            path=path,
            paragraph_index=paragraph_index,
            text=text,
            url=url,
            confirm=confirm,
        )

    def _add_footnote(
        path: str,
        paragraph_index: int,
        text: str,
        confirm: bool = False,
    ) -> dict:
        """
        Add a footnote reference to a paragraph.
        Requires document to have an existing footnotes part
        (open in Word and add one footnote manually first if absent).
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        """
        return _server_docx_tools.add_footnote(
            runtime,
            path=path,
            paragraph_index=paragraph_index,
            text=text,
            confirm=confirm,
        )

    def _bulk_update_paragraphs(
        path: str,
        updates: list | str,
        confirm: bool = False,
    ) -> dict:
        """
        Update multiple paragraphs by index in a single operation (one save).
        updates: list of dicts with keys: paragraph_index (int), new_text (str). Max 500.
        Per-item errors are collected and returned; partial success is supported.
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        """
        return _server_docx_tools.bulk_update_paragraphs(
            runtime,
            path=path,
            updates=updates,
            confirm=confirm,
        )

    def _bulk_update_table_cells(
        path: str,
        updates: list | str,
        confirm: bool = False,
    ) -> dict:
        """
        Update multiple table cells in a single operation (one save).
        updates: list of dicts with keys: table_index (int), row (int), col (int), new_text (str).
        Max 200 updates per call. Per-item errors are collected; partial success supported.
        Requires WORD_ENABLE_WRITE=true and confirm=True.

        DEPRECATED: Use table(operation="bulk_update_cells", path=path, updates=[...]) instead.
        This alias is kept for backward compatibility and returns identical output.
        """
        return _server_docx_tools.bulk_update_table_cells(
            runtime,
            path=path,
            updates=updates,
            confirm=confirm,
        )

    return {
        "add_list": bind_tool(_add_list, "add_list"),
        "insert_paragraph": bind_tool(_insert_paragraph, "insert_paragraph"),
        "bulk_add_paragraphs": bind_tool(_bulk_add_paragraphs, "bulk_add_paragraphs"),
        "get_headers_footers": bind_tool(_get_headers_footers, "get_headers_footers"),
        "add_hyperlink": bind_tool(_add_hyperlink, "add_hyperlink"),
        "add_footnote": bind_tool(_add_footnote, "add_footnote"),
        "bulk_update_paragraphs": bind_tool(_bulk_update_paragraphs, "bulk_update_paragraphs"),
        "bulk_update_table_cells": bind_tool(_bulk_update_table_cells, "bulk_update_table_cells"),
    }
