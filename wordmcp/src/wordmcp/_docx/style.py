from __future__ import annotations

from docx.enum.style import WD_STYLE_TYPE

from wordmcp._docx import _facade


def list_styles(path: str, style_type: str | None = None) -> list:
    facade = _facade()
    valid_types = {
        "paragraph": WD_STYLE_TYPE.PARAGRAPH,
        "character": WD_STYLE_TYPE.CHARACTER,
        "table": WD_STYLE_TYPE.TABLE,
        "numbering": WD_STYLE_TYPE.LIST,
    }
    if style_type and style_type not in valid_types:
        raise facade.ValidationError(
            f"Unknown style_type {style_type!r}. Valid: {', '.join(valid_types)}"
        )

    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)
    filter_type = valid_types.get(style_type) if style_type else None

    def _type_str(style_value) -> str:
        if hasattr(style_value, "name"):
            return style_value.name
        if hasattr(style_value, "_member_name"):
            return style_value._member_name
        return str(int(style_value)) if isinstance(style_value, int) else str(style_value)

    return [
        {"name": style.name, "type": _type_str(style.type), "builtin": style.builtin}
        for style in doc.styles
        if filter_type is None or style.type == filter_type
    ]


def apply_style(path: str, paragraph_index: int, style: str, confirm: bool = False) -> dict:
    facade = _facade()
    facade._check_write()
    facade._check_confirm(confirm)
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)
    paragraphs = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paragraphs):
        raise facade.NotFoundError(
            f"Paragraph index {paragraph_index} out of range (0-{len(paragraphs)-1})"
        )
    try:
        doc.styles[style]
    except KeyError:
        raise facade.ValidationError(f"Unknown style: {style!r}")

    paragraphs[paragraph_index].style = doc.styles[style]
    facade._atomic_save(doc, resolved)
    facade._evict_doc(resolved)
    facade._audit_log("apply_style", resolved)
    return {"paragraph_index": paragraph_index, "style": style}
