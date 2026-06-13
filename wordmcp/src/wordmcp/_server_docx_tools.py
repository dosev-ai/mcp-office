from __future__ import annotations

import json

from fastmcp.exceptions import ToolError

from wordmcp._server_runtime import ServerRuntime


def read_document(runtime: ServerRuntime, path: str) -> dict:
    return runtime.call_doc("read_document", path)


def get_document_metadata(runtime: ServerRuntime, path: str) -> dict:
    return runtime.call_doc("get_document_metadata", path)


def list_paragraphs(runtime: ServerRuntime, path: str) -> list[dict]:
    return runtime.call_doc("list_paragraphs", path)


def read_paragraph(runtime: ServerRuntime, path: str, paragraph_index: int) -> dict:
    return runtime.call_doc("read_paragraph", path, paragraph_index)


def list_tables(runtime: ServerRuntime, path: str) -> list[dict]:
    return runtime.call_doc("list_tables", path)


def read_table(runtime: ServerRuntime, path: str, table_index: int) -> dict:
    return runtime.call_doc("read_table", path, table_index)


def export_as_text(runtime: ServerRuntime, path: str) -> dict:
    return runtime.call_doc("export_as_text", path)


def add_paragraph(
    runtime: ServerRuntime,
    path: str,
    text: str,
    style: str | None = None,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc("add_paragraph", path, text, style=style, confirm=confirm)


def add_heading(
    runtime: ServerRuntime,
    path: str,
    text: str,
    level: int = 1,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc("add_heading", path, text, level=level, confirm=confirm)


def add_page_break(runtime: ServerRuntime, path: str, confirm: bool = False) -> dict:
    return runtime.call_doc("add_page_break", path, confirm=confirm)


def add_table(
    runtime: ServerRuntime,
    path: str,
    rows: int,
    cols: int,
    data: list[list[str]] | str | None = None,
    style: str | None = None,
    confirm: bool = False,
) -> dict:
    if data is not None:
        data = _coerce_list_param(data, "data")
        for i, row in enumerate(data):
            if not isinstance(row, list):
                raise ToolError(
                    f"data[{i}]: expected a JSON array, got {type(row).__name__}"
                )
    return runtime.call_doc(
        "add_table",
        path,
        rows,
        cols,
        data=data,
        style=style,
        confirm=confirm,
    )


def insert_image(
    runtime: ServerRuntime,
    path: str,
    image_path: str,
    width_inches: float | None = None,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "insert_image",
        path,
        image_path,
        width_inches=width_inches,
        confirm=confirm,
    )


def find_replace(
    runtime: ServerRuntime,
    path: str,
    find_text: str,
    replace_text: str,
    paragraph_index: int | None = None,
    occurrence: int | None = None,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "find_replace",
        path,
        find_text,
        replace_text,
        paragraph_index=paragraph_index,
        occurrence=occurrence,
        confirm=confirm,
    )


def save(runtime: ServerRuntime, path: str, confirm: bool = False) -> dict:
    return runtime.call_doc("save", path, confirm=confirm)


def manage_comments(
    runtime: ServerRuntime,
    path: str,
    operation: str,
    text: str | None = None,
    author: str | None = None,
    paragraph_index: int | None = None,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "manage_comments",
        path=path,
        operation=operation,
        text=text,
        author=author,
        paragraph_index=paragraph_index,
        confirm=confirm,
    )


def create_document(
    runtime: ServerRuntime,
    path: str,
    title: str | None = None,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc("create_document", path=path, title=title, confirm=confirm)


def list_headings(runtime: ServerRuntime, path: str) -> list:
    return runtime.call_doc("list_headings", path=path)


def read_section(runtime: ServerRuntime, path: str, heading_index: int) -> dict:
    return runtime.call_doc("read_section", path=path, heading_index=heading_index)


def search_text(
    runtime: ServerRuntime,
    path: str,
    query: str,
    case_sensitive: bool = True,
    max_results: int = 100,
) -> dict:
    return runtime.call_doc(
        "search_text",
        path=path,
        query=query,
        case_sensitive=case_sensitive,
        max_results=max_results,
    )


def list_styles(
    runtime: ServerRuntime,
    path: str,
    style_type: str | None = None,
) -> list:
    return runtime.call_doc("list_styles", path=path, style_type=style_type)


def apply_style(
    runtime: ServerRuntime,
    path: str,
    paragraph_index: int,
    style: str,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "apply_style",
        path=path,
        paragraph_index=paragraph_index,
        style=style,
        confirm=confirm,
    )


def update_paragraph(
    runtime: ServerRuntime,
    path: str,
    paragraph_index: int,
    new_text: str,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "update_paragraph",
        path=path,
        paragraph_index=paragraph_index,
        new_text=new_text,
        confirm=confirm,
    )


def delete_paragraph(
    runtime: ServerRuntime,
    path: str,
    paragraph_index: int,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "delete_paragraph",
        path=path,
        paragraph_index=paragraph_index,
        confirm=confirm,
    )


def update_table_cell(
    runtime: ServerRuntime,
    path: str,
    table_index: int,
    row: int,
    col: int,
    new_text: str,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "update_table_cell",
        path=path,
        table_index=table_index,
        row=row,
        col=col,
        new_text=new_text,
        confirm=confirm,
    )


def set_document_properties(
    runtime: ServerRuntime,
    path: str,
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    category: str | None = None,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "set_document_properties",
        path=path,
        title=title,
        author=author,
        subject=subject,
        keywords=keywords,
        category=category,
        confirm=confirm,
    )


def _coerce_list_param(value: object, name: str) -> list:
    """Coerce a JSON-string list param to a real list (FastMCP Pydantic workaround)."""
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except (json.JSONDecodeError, TypeError) as exc:
            raise ToolError(f"{name}: invalid JSON string: {exc}")
        if not isinstance(value, list):
            raise ToolError(
                f"{name}: expected a JSON array, got {type(value).__name__}"
            )
    return value  # type: ignore[return-value]


def add_list(
    runtime: ServerRuntime,
    path: str,
    items: list | str,
    list_type: str = "bullet",
    level: int = 0,
    confirm: bool = False,
) -> dict:
    items = _coerce_list_param(items, "items")
    return runtime.call_doc(
        "add_list",
        path=path,
        items=items,
        list_type=list_type,
        level=level,
        confirm=confirm,
    )


def insert_paragraph(
    runtime: ServerRuntime,
    path: str,
    paragraph_index: int,
    text: str,
    style: str | None = None,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "insert_paragraph",
        path=path,
        paragraph_index=paragraph_index,
        text=text,
        style=style,
        confirm=confirm,
    )


def bulk_add_paragraphs(
    runtime: ServerRuntime,
    path: str,
    paragraphs: list | str,
    confirm: bool = False,
) -> dict:
    paragraphs = _coerce_list_param(paragraphs, "paragraphs")
    return runtime.call_doc(
        "bulk_add_paragraphs",
        path=path,
        paragraphs=paragraphs,
        confirm=confirm,
    )


def export_to_markdown(runtime: ServerRuntime, path: str) -> dict:
    return runtime.call_doc("export_to_markdown", path=path)


def get_headers_footers(runtime: ServerRuntime, path: str) -> list:
    return runtime.call_doc("get_headers_footers", path=path)


def add_hyperlink(
    runtime: ServerRuntime,
    path: str,
    paragraph_index: int,
    text: str,
    url: str,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "add_hyperlink",
        path=path,
        paragraph_index=paragraph_index,
        text=text,
        url=url,
        confirm=confirm,
    )


def add_footnote(
    runtime: ServerRuntime,
    path: str,
    paragraph_index: int,
    text: str,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "add_footnote",
        path=path,
        paragraph_index=paragraph_index,
        text=text,
        confirm=confirm,
    )


def bulk_update_paragraphs(
    runtime: ServerRuntime,
    path: str,
    updates: list | str,
    confirm: bool = False,
) -> dict:
    updates = _coerce_list_param(updates, "updates")
    return runtime.call_doc(
        "bulk_update_paragraphs", path=path, updates=updates, confirm=confirm
    )


def set_paragraph_format(
    runtime: ServerRuntime,
    path: str,
    paragraph_index: int,
    space_before: float | None = None,
    space_after: float | None = None,
    line_spacing: float | None = None,
    table_cell: dict | None = None,
    confirm: bool = False,
) -> dict:
    return runtime.call_doc(
        "set_paragraph_format",
        path=path,
        paragraph_index=paragraph_index,
        space_before=space_before,
        space_after=space_after,
        line_spacing=line_spacing,
        table_cell=table_cell,
        confirm=confirm,
    )


def bulk_update_table_cells(
    runtime: ServerRuntime,
    path: str,
    updates: list | str,
    confirm: bool = False,
) -> dict:
    updates = _coerce_list_param(updates, "updates")
    return runtime.call_doc(
        "bulk_update_table_cells", path=path, updates=updates, confirm=confirm
    )
