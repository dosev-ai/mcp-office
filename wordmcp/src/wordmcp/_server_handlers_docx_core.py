from __future__ import annotations

from collections.abc import Callable
from typing import Any

from wordmcp import _server_docx_tools
from wordmcp._server_runtime import ServerRuntime


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _get_document_metadata(path: str) -> dict:
        """Return full document metadata: title, author, subject, keywords, dates, revision.

        Times are as stored in the .docx file (typically UTC, no timezone suffix).

        DEPRECATED: Use read_document(path, scope='metadata') instead.
        This alias is kept for backward compatibility and returns identical output.
        """
        return _server_docx_tools.get_document_metadata(runtime, path)

    def _list_paragraphs(path: str) -> list[dict]:
        """Return summary of all body paragraphs: index, style, text_preview, outline_level.

        NOTE: Table-cell paragraphs are excluded. For full run detail, use read_paragraph.
        Word_count covers body paragraphs only; table cell text is excluded.
        Large documents may return many entries (no pagination in Phase 1).
        """
        return _server_docx_tools.list_paragraphs(runtime, path)

    def _read_paragraph(path: str, paragraph_index: int) -> dict:
        """Return full detail for one body paragraph including all run properties."""
        return _server_docx_tools.read_paragraph(runtime, path, paragraph_index)

    def _list_tables(path: str) -> list[dict]:
        """Return summary of all tables in the document: index, rows, cols, style.

        DEPRECATED: Use table(operation="list", path=path) instead.
        This alias is kept for backward compatibility and returns identical output.
        """
        return _server_docx_tools.list_tables(runtime, path)

    def _read_table(path: str, table_index: int) -> dict:
        """Return full cell data for one table as a 2D row-major array.

        NOTE: Merged cells are repeated (python-docx behaviour); data shows duplicated
        values for merged cells. Large tables may produce oversized responses (Phase 1
        limitation, no pagination).

        DEPRECATED: Use table(operation="read", path=path, table_index=N) instead.
        This alias is kept for backward compatibility and returns identical output.
        """
        return _server_docx_tools.read_table(runtime, path, table_index)

    def _add_paragraph(
        path: str,
        text: str,
        style: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """Append a paragraph to the end of the document.

        Requires WORD_ENABLE_WRITE=true and confirm=True.
        style: Word paragraph style name (e.g. 'Normal', 'Body Text'). None = 'Normal'.
        """
        return _server_docx_tools.add_paragraph(
            runtime,
            path,
            text,
            style=style,
            confirm=confirm,
        )

    def _add_heading(
        path: str,
        text: str,
        level: int = 1,
        confirm: bool = False,
    ) -> dict:
        """Append a heading paragraph. level 0='Title', 1-9='Heading N'.

        Requires WORD_ENABLE_WRITE=true and confirm=True.
        level must be 0-9; level 10+ raises an error.
        """
        return _server_docx_tools.add_heading(
            runtime,
            path,
            text,
            level=level,
            confirm=confirm,
        )

    def _add_page_break(path: str, confirm: bool = False) -> dict:
        """Insert a page break at the end of the document body.

        Requires WORD_ENABLE_WRITE=true and confirm=True.
        """
        return _server_docx_tools.add_page_break(runtime, path, confirm=confirm)

    def _add_table(
        path: str,
        rows: int,
        cols: int,
        data: list[list[str]] | str | None = None,
        style: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """Append a table to the end of the document.

        Requires WORD_ENABLE_WRITE=true and confirm=True.
        data: optional 2D list (rows x cols) of string cell values. None = empty table.
        style: Word table style name. None = 'Normal Table'.

        DEPRECATED: Use table(operation="add", path=path, rows=N, cols=N) instead.
        This alias is kept for backward compatibility and returns identical output.
        """
        return _server_docx_tools.add_table(
            runtime,
            path,
            rows,
            cols,
            data=data,
            style=style,
            confirm=confirm,
        )

    def _insert_image(
        path: str,
        image_path: str,
        width_inches: float | None = None,
        confirm: bool = False,
    ) -> dict:
        """Insert an inline image at the end of the document body.

        Requires WORD_ENABLE_WRITE=true and confirm=True.
        image_path must be under WORD_ALLOWLIST_ROOTS.
        Supported formats: PNG, JPEG, GIF, BMP, TIFF. SVG not supported.
        width_inches: resize width (aspect ratio preserved). None = native size.
        """
        return _server_docx_tools.insert_image(
            runtime,
            path,
            image_path,
            width_inches=width_inches,
            confirm=confirm,
        )

    def _find_replace(
        path: str,
        find_text: str,
        replace_text: str,
        paragraph_index: int | None = None,
        occurrence: int | None = None,
        confirm: bool = False,
    ) -> dict:
        """Find and replace text across body paragraphs and table cells.

        Supports cross-run matching: text that spans multiple <w:r> run elements
        within the same paragraph is correctly found and replaced.
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        Case-sensitive. find_text must not be empty.
        paragraph_index: if provided, limit replacement to that body paragraph only.
        occurrence: if provided (1-based), replace only the Nth occurrence in the document.
        """
        return _server_docx_tools.find_replace(
            runtime,
            path,
            find_text,
            replace_text,
            paragraph_index=paragraph_index,
            occurrence=occurrence,
            confirm=confirm,
        )

    def _save(path: str, confirm: bool = False) -> dict:
        """Explicitly persist the document to disk.

        Requires WORD_ENABLE_WRITE=true and confirm=True.
        In Phase 1, all write tools auto-save after every mutation. save() is provided
        for future batched-edit workflows (Phase 2+). Calling it after a Phase 1 write
        tool is a harmless double-save.
        """
        return _server_docx_tools.save(runtime, path, confirm=confirm)

    return {
        "get_document_metadata": bind_tool(
            _get_document_metadata,
            "get_document_metadata",
        ),
        "list_paragraphs": bind_tool(_list_paragraphs, "list_paragraphs"),
        "read_paragraph": bind_tool(_read_paragraph, "read_paragraph"),
        "list_tables": bind_tool(_list_tables, "list_tables"),
        "read_table": bind_tool(_read_table, "read_table"),
        "add_paragraph": bind_tool(_add_paragraph, "add_paragraph"),
        "add_heading": bind_tool(_add_heading, "add_heading"),
        "add_page_break": bind_tool(_add_page_break, "add_page_break"),
        "add_table": bind_tool(_add_table, "add_table"),
        "insert_image": bind_tool(_insert_image, "insert_image"),
        "find_replace": bind_tool(_find_replace, "find_replace"),
        "save": bind_tool(_save, "save"),
    }
