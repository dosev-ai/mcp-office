from __future__ import annotations

# Re-export shim — write.py split into sub-modules.
from wordmcp._docx.write_doc import create_document, save, set_document_properties
from wordmcp._docx.write_paragraph import (
    add_paragraph, add_heading, add_page_break, bulk_add_paragraphs,
    bulk_update_paragraphs, delete_paragraph, insert_paragraph, update_paragraph,
)
from wordmcp._docx.write_table import add_table, bulk_update_table_cells, update_table_cell
from wordmcp._docx.write_content import add_list, find_replace, insert_image, set_paragraph_format

__all__ = [
    "add_heading", "add_list", "add_page_break", "add_paragraph", "add_table",
    "bulk_add_paragraphs", "bulk_update_paragraphs", "bulk_update_table_cells",
    "create_document", "delete_paragraph", "find_replace", "insert_image",
    "insert_paragraph", "save", "set_document_properties", "set_paragraph_format",
    "update_paragraph", "update_table_cell",
]
