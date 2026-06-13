from __future__ import annotations

import hashlib
import re
from pathlib import Path
from typing import Any

from wordmcp._docx.core import _check_path, _load_doc, _outline_level

# Parity: must equal workflowruntime.contracts.constants.ARTIFACT_ID_PREFIX_DOC
_ARTIFACT_ID_PREFIX_DOC = "doc:"

_FORMULA_LEADING = re.compile(r"^[=+\-@\t\r\n]+")


def _sanitize_stem(stem: str) -> str:
    """Strip formula-injection leading chars from a file stem. OWASP A03."""
    result = _FORMULA_LEADING.sub("", stem)
    if not result:
        # Degenerate input (all chars stripped) — use hash fallback for non-empty artifact_id
        result = f"_sanitized_{hashlib.md5(stem.encode()).hexdigest()[:8]}"
    return result


_METADATA_FIELDS = (
    "title",
    "author",
    "subject",
    "category",
    "keywords",
    "last_modified_by",
    "revision",
    "created",
    "modified",
)


def get_document_context(path: str) -> dict[str, Any]:
    resolved = _check_path(path)
    doc = _load_doc(resolved)

    metadata = _build_metadata(doc.core_properties)
    headings = _extract_headings(doc)
    content: dict[str, Any] = {
        **metadata,
        "paragraph_count": len(doc.paragraphs),
        "table_count": len(doc.tables),
        "section_count": len(doc.sections),
        "heading_count": len(headings),
        "headings": headings,
    }

    return {
        "packet_type": "focused",
        "adapter_type": "wordmcp",
        "artifact_type": "document",
        # artifact_id is session-stable; stems change on rename by design
        "artifact_id": f"{_ARTIFACT_ID_PREFIX_DOC}{_sanitize_stem(resolved.stem)}",
        "source_ref": str(resolved),
        "summary": _build_summary(resolved, content),
        "content": content,
    }


def _build_metadata(core_properties: Any) -> dict[str, Any]:
    values = {
        "title": core_properties.title or None,
        "author": core_properties.author or None,
        "subject": core_properties.subject or None,
        "category": core_properties.category or None,
        "keywords": core_properties.keywords or None,
        "last_modified_by": core_properties.last_modified_by or None,
        "revision": core_properties.revision if core_properties.revision else None,
        "created": _to_iso(core_properties.created),
        "modified": _to_iso(core_properties.modified),
    }
    return {field: values[field] for field in _METADATA_FIELDS}


def _extract_headings(doc: Any) -> list[str]:
    headings: list[str] = []
    for paragraph in doc.paragraphs:
        if _outline_level(paragraph) is None:
            continue
        text = paragraph.text.strip()
        if text:
            headings.append(text)
    return headings


def _build_summary(resolved: Path, content: dict[str, Any]) -> str:
    summary_parts = [
        f"{resolved.name}: {content['paragraph_count']} paragraphs",
        f"{content['table_count']} tables",
        f"{content['section_count']} sections",
        f"{content['heading_count']} headings",
    ]
    if content["title"]:
        summary_parts.append(f"title={content['title']}")
    if content["author"]:
        summary_parts.append(f"author={content['author']}")
    if content["headings"]:
        preview = ", ".join(content["headings"][:3])
        summary_parts.append(f"headings={preview}")
    return "; ".join(summary_parts)


def _to_iso(value: Any) -> str | None:
    return value.isoformat() if value else None
