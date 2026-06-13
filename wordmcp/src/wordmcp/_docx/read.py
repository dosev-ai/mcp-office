from __future__ import annotations

from typing import Any

from wordmcp._docx import _facade
from wordmcp._docx.core import _outline_level, _paragraph_to_dict


def read_document(path: str) -> dict:
    facade = _facade()
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)

    image_count = sum(1 for rel in doc.part.rels.values() if rel.reltype == facade.IMAGE_RELTYPE)
    title = doc.core_properties.title or None
    author = doc.core_properties.author or None
    word_count = sum(len(paragraph.text.split()) for paragraph in doc.paragraphs)

    return {
        "path": str(resolved),
        "paragraph_count": len(doc.paragraphs),
        "table_count": len(doc.tables),
        "section_count": len(doc.sections),
        "image_count": image_count,
        "title": title,
        "author": author,
        "word_count": word_count,
    }


def get_document_metadata(path: str) -> dict:
    facade = _facade()
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)
    cp = doc.core_properties

    def _dt(value: Any) -> str | None:
        return value.isoformat() if value else None

    return {
        "path": str(resolved),
        "title": cp.title or None,
        "author": cp.author or None,
        "subject": cp.subject or None,
        "keywords": cp.keywords or None,
        "category": cp.category or None,
        "last_modified_by": cp.last_modified_by or None,
        "revision": cp.revision if cp.revision else None,
        "version": cp.version or None,
        "created": _dt(cp.created),
        "modified": _dt(cp.modified),
    }


def list_paragraphs(path: str) -> list[dict]:
    facade = _facade()
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)
    return [
        {
            "index": index,
            "style": paragraph.style.name,
            "text_preview": paragraph.text[:120],
            "outline_level": _outline_level(paragraph),
            "run_count": len(paragraph.runs),
        }
        for index, paragraph in enumerate(doc.paragraphs)
    ]


def read_paragraph(path: str, paragraph_index: int) -> dict:
    facade = _facade()
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)

    paragraphs = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paragraphs):
        raise facade.NotFoundError(
            f"paragraph_index {paragraph_index} out of range (0-{len(paragraphs) - 1})"
        )

    return _paragraph_to_dict(paragraphs[paragraph_index], paragraph_index)


def list_tables(path: str) -> list[dict]:
    facade = _facade()
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)
    return [
        {
            "index": index,
            "rows": len(table.rows),
            "cols": len(table.columns),
            "style": table.style.name,
        }
        for index, table in enumerate(doc.tables)
    ]


def read_table(path: str, table_index: int) -> dict:
    facade = _facade()
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)

    tables = doc.tables
    if table_index < 0 or table_index >= len(tables):
        raise facade.NotFoundError(
            f"table_index {table_index} out of range (0-{len(tables) - 1})"
        )

    table = tables[table_index]
    data = [
        [" ".join(paragraph.text for paragraph in cell.paragraphs).strip() for cell in row.cells]
        for row in table.rows
    ]

    return {
        "index": table_index,
        "rows": len(table.rows),
        "cols": len(table.columns),
        "style": table.style.name,
        "data": data,
    }


def list_headings(path: str) -> list:
    facade = _facade()
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)
    headings = []
    for index, paragraph in enumerate(doc.paragraphs):
        level = _outline_level(paragraph)
        if level is not None:
            headings.append(
                {
                    "index": index,
                    "level": level,
                    "text": paragraph.text,
                    "style": paragraph.style.name,
                }
            )
    return headings


def read_section(path: str, heading_index: int) -> dict:
    facade = _facade()
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)
    paragraphs = doc.paragraphs
    if heading_index < 0 or heading_index >= len(paragraphs):
        raise facade.NotFoundError(
            f"Paragraph index {heading_index} out of range (0-{len(paragraphs) - 1})"
        )
    start_level = _outline_level(paragraphs[heading_index])
    if start_level is None:
        raise facade.ValidationError(f"Paragraph {heading_index} is not a heading")

    body = []
    for index in range(heading_index + 1, len(paragraphs)):
        level = _outline_level(paragraphs[index])
        if level is not None and level <= start_level:
            break
        body.append(_paragraph_to_dict(paragraphs[index], index))

    return {
        "heading": _paragraph_to_dict(paragraphs[heading_index], heading_index),
        "body": body,
    }


def search_text(
    path: str,
    query: str,
    case_sensitive: bool = True,
    max_results: int = 100,
) -> dict:
    facade = _facade()
    if not query:
        raise facade.ValidationError("query must not be empty")
    if max_results < 1 or max_results > 1000:
        raise facade.ValidationError("max_results must be between 1 and 1000")

    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)
    q = query if case_sensitive else query.lower()
    results: list[dict] = []
    seen = 0

    def _scan_para(para: Any, idx: int) -> None:
        nonlocal seen
        if seen >= max_results:
            return
        text = para.text if case_sensitive else para.text.lower()
        count = text.count(q)
        if count > 0:
            results.append(
                {
                    "paragraph_index": idx,
                    "context": para.text[:200],
                    "match_count": count,
                }
            )
            seen += count

    for index, paragraph in enumerate(doc.paragraphs):
        _scan_para(paragraph, index)

    para_offset = len(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _scan_para(paragraph, para_offset)
                    para_offset += 1

    return {
        "results": results,
        "total_matches": sum(result["match_count"] for result in results),
        "query": query,
    }
