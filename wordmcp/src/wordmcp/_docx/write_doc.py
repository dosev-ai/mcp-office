from __future__ import annotations

from docx import Document

from wordmcp._docx import _facade


def create_document(path: str, title: str | None = None, confirm: bool = False) -> dict:
    facade = _facade()
    facade._check_write()
    facade._check_confirm(confirm)
    resolved = facade._check_path(path, must_exist=False)
    if resolved.exists():
        raise facade.ValidationError(
            f"File already exists: {resolved}. Delete it first or use a different path."
        )
    new_doc = Document()
    if title:
        new_doc.core_properties.title = title
    facade._atomic_save(new_doc, resolved)
    facade._evict_doc(resolved)
    facade._audit_log("create_document", resolved)
    return {"created": True, "path": str(resolved)}


def save(path: str, confirm: bool = False) -> dict:
    facade = _facade()
    facade._check_write()
    facade._check_confirm(confirm)
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)

    facade._atomic_save(doc, resolved)
    facade._evict_doc(resolved)
    facade._audit_log("save", resolved)
    return {"saved": True, "path": str(resolved)}


def set_document_properties(
    path: str,
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    category: str | None = None,
    confirm: bool = False,
) -> dict:
    facade = _facade()
    facade._check_write()
    facade._check_confirm(confirm)
    if title is None and author is None and subject is None and keywords is None and category is None:
        raise facade.ValidationError("At least one property must be provided")

    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)
    cp = doc.core_properties
    if title is not None:
        cp.title = title
    if author is not None:
        cp.author = author
    if subject is not None:
        cp.subject = subject
    if keywords is not None:
        cp.keywords = keywords
    if category is not None:
        cp.category = category

    facade._atomic_save(doc, resolved)
    facade._evict_doc(resolved)
    facade._audit_log("set_document_properties", resolved)
    updated = {
        key: value
        for key, value in {
            "title": title,
            "author": author,
            "subject": subject,
            "keywords": keywords,
            "category": category,
        }.items()
        if value is not None
    }
    return {"updated": True, "properties": updated}
