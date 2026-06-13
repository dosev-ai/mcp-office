from __future__ import annotations

from collections.abc import Callable
from typing import Any

from wordmcp import _server_docx_tools
from wordmcp._server_runtime import ServerRuntime


def register(
    bind_tool: Callable[[Callable[..., Any], str], Any],
    runtime: ServerRuntime,
) -> dict[str, Any]:
    def _manage_comments(
        path: str,
        operation: str,
        text: str | None = None,
        author: str | None = None,
        paragraph_index: int | None = None,
        confirm: bool = False,
    ) -> dict:
        """
        Manage document comments (oxml implementation).

        Operations:
          'list'    — read-only; returns list of {id, author, date, text} dicts.
          'add'     — adds a comment. Requires WORD_ENABLE_WRITE=true and confirm=True.
                      text is required. author defaults to 'MCP'.
                      paragraph_index anchors the comment to that paragraph (optional).
          'resolve' — marks comment resolved (w:done=1). paragraph_index = comment_id.
                      Requires WORD_ENABLE_WRITE=true and confirm=True.
          'delete'  — removes comment + body anchors. paragraph_index = comment_id.
                      Requires WORD_ENABLE_WRITE=true and confirm=True.
        Invalid operation raises ToolError.
        """
        return _server_docx_tools.manage_comments(
            runtime,
            path=path,
            operation=operation,
            text=text,
            author=author,
            paragraph_index=paragraph_index,
            confirm=confirm,
        )

    def _create_document(
        path: str,
        title: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """
        Create a new empty Word document at the given path.
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        Raises ToolError if the file already exists.
        """
        return _server_docx_tools.create_document(
            runtime,
            path=path,
            title=title,
            confirm=confirm,
        )

    def _list_headings(path: str) -> list:
        """
        List all headings in the document with their outline level (1-9), paragraph index, and style name.
        Returns an empty list if the document has no headings.

        DEPRECATED: Use read_document(path, scope='headings') instead.
        This alias is kept for backward compatibility and returns identical output.
        """
        return _server_docx_tools.list_headings(runtime, path=path)

    def _read_section(path: str, heading_index: int) -> dict:
        """
        Read the content of a section identified by its heading paragraph index.
        Returns the heading and all body paragraphs until the next same-or-higher-level heading.
        Use list_headings() to discover heading indices.

        DEPRECATED: Use read_document(path, scope='section', section_index=N) instead.
        This alias is kept for backward compatibility and returns identical output.
        """
        return _server_docx_tools.read_section(
            runtime,
            path=path,
            heading_index=heading_index,
        )

    def _search_text(
        path: str,
        query: str,
        case_sensitive: bool = True,
        max_results: int = 100,
    ) -> dict:
        """
        Search for text in document paragraphs and table cells.
        Returns matching paragraph indices, context (first 200 chars), and match count per paragraph.
        max_results: cap on total matches returned (1-1000, default 100).
        """
        return _server_docx_tools.search_text(
            runtime,
            path=path,
            query=query,
            case_sensitive=case_sensitive,
            max_results=max_results,
        )

    def _list_styles(path: str, style_type: str | None = None) -> list:
        """
        List all styles available in the document.
        style_type: optional filter — 'paragraph' | 'character' | 'table' | 'numbering'.
        Returns name, type, and whether the style is built-in.
        """
        return _server_docx_tools.list_styles(runtime, path=path, style_type=style_type)

    def _apply_style(
        path: str,
        paragraph_index: int,
        style: str,
        confirm: bool = False,
    ) -> dict:
        """
        Apply a named style to a paragraph by index.
        Use list_styles() to discover available styles.
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        Raises ToolError for unknown style or out-of-range index.
        """
        return _server_docx_tools.apply_style(
            runtime,
            path=path,
            paragraph_index=paragraph_index,
            style=style,
            confirm=confirm,
        )

    def _update_paragraph(
        path: str,
        paragraph_index: int,
        new_text: str,
        confirm: bool = False,
    ) -> dict:
        """
        Replace the text content of a paragraph identified by index.
        Preserves paragraph style (heading, normal, etc.) — only the runs (text) are replaced.
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        """
        return _server_docx_tools.update_paragraph(
            runtime,
            path=path,
            paragraph_index=paragraph_index,
            new_text=new_text,
            confirm=confirm,
        )

    def _delete_paragraph(
        path: str,
        paragraph_index: int,
        confirm: bool = False,
    ) -> dict:
        """
        Delete a paragraph from the document by index.
        WARNING: All subsequent paragraph indices shift down by 1 after deletion.
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        """
        return _server_docx_tools.delete_paragraph(
            runtime,
            path=path,
            paragraph_index=paragraph_index,
            confirm=confirm,
        )

    def _update_table_cell(
        path: str,
        table_index: int,
        row: int,
        col: int,
        new_text: str,
        confirm: bool = False,
    ) -> dict:
        """
        Set the text content of a specific table cell.
        Use list_tables() to discover table count; use read_table() to see row/col structure.
        Requires WORD_ENABLE_WRITE=true and confirm=True.

        DEPRECATED: Use table(operation="update_cell", path=path, table_index=N,
            row_index=R, col_index=C, new_text=T) instead.
        This alias is kept for backward compatibility and returns identical output.
        """
        return _server_docx_tools.update_table_cell(
            runtime,
            path=path,
            table_index=table_index,
            row=row,
            col=col,
            new_text=new_text,
            confirm=confirm,
        )

    def _set_paragraph_format(
        path: str,
        paragraph_index: int,
        space_before: float | None = None,
        space_after: float | None = None,
        line_spacing: float | None = None,
        table_cell: dict | None = None,
        confirm: bool = False,
    ) -> dict:
        """
        DEPRECATED: Use paragraph(operation="set_format", ...) instead.
        This standalone tool is retained as a deprecated delegate for
        backward compatibility and will be removed in a future phase.

        Apply spacing and line-spacing to a paragraph's ParagraphFormat.

        paragraph_index: 0-based index of the body paragraph to format.
        space_before: space before paragraph in points (pt). None = no change.
        space_after: space after paragraph in points (pt). None = no change.
        line_spacing: line spacing multiplier (1.0 = single, 1.5 = one-and-a-half,
            2.0 = double). None = no change.
        table_cell: optional dict with 'row' (int) and 'col' (int) keys. When
            provided, targets a paragraph inside that cell of the first table,
            and paragraph_index refers to the paragraph index within the cell.
        At least one of space_before, space_after, or line_spacing must be provided.
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        """
        return _server_docx_tools.set_paragraph_format(
            runtime,
            path=path,
            paragraph_index=paragraph_index,
            space_before=space_before,
            space_after=space_after,
            line_spacing=line_spacing,
            table_cell=table_cell,
            confirm=confirm,
        )

    def _set_document_properties(
        path: str,
        title: str | None = None,
        author: str | None = None,
        subject: str | None = None,
        keywords: str | None = None,
        category: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """
        Set document core properties (metadata).
        At least one property must be provided.
        Requires WORD_ENABLE_WRITE=true and confirm=True.
        Properties: title, author, subject, keywords, category.
        """
        return _server_docx_tools.set_document_properties(
            runtime,
            path=path,
            title=title,
            author=author,
            subject=subject,
            keywords=keywords,
            category=category,
            confirm=confirm,
        )

    return {
        "manage_comments": bind_tool(_manage_comments, "manage_comments"),
        "create_document": bind_tool(_create_document, "create_document"),
        "list_headings": bind_tool(_list_headings, "list_headings"),
        "read_section": bind_tool(_read_section, "read_section"),
        "search_text": bind_tool(_search_text, "search_text"),
        "list_styles": bind_tool(_list_styles, "list_styles"),
        "apply_style": bind_tool(_apply_style, "apply_style"),
        "update_paragraph": bind_tool(_update_paragraph, "update_paragraph"),
        "delete_paragraph": bind_tool(_delete_paragraph, "delete_paragraph"),
        "update_table_cell": bind_tool(_update_table_cell, "update_table_cell"),
        "set_document_properties": bind_tool(
            _set_document_properties,
            "set_document_properties",
        ),
        "set_paragraph_format": bind_tool(_set_paragraph_format, "set_paragraph_format"),
    }
