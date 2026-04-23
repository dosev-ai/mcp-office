"""Field lineage and normalization for WorkbookMetadata.

Pure stdlib module — no FastMCP imports, no server imports, no COM imports.
WorkbookMetadata is only imported under TYPE_CHECKING to avoid runtime circular imports.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from excelmcp._metadata_contract import WorkbookMetadata

FIELD_ALIASES: dict[str, str] = {
    # legacy key      : canonical key
    "name":            "workbook_name",
    "sheets":          "sheet_names",
    "active":          "active_sheet",
    "formula":         "has_formulas",
    "formulas":        "has_formulas",
    "modified":        "last_modified",
    "bytes":           "size_bytes",
    "ext":             "format",
    "version":         "contract_version",
}

_REQUIRED_CANONICAL: tuple[str, ...] = (
    "path",
    "workbook_name",
    "sheet_names",
    "active_sheet",
    "sheet_count",
    "has_formulas",
    "last_modified",
    "size_bytes",
    "format",
    "contract_version",
)


def map_lineage_fields(raw: dict) -> "WorkbookMetadata":
    """Normalize a dict that may use legacy key aliases to canonical WorkbookMetadata shape.

    Rules:
      - Known aliases (FIELD_ALIASES) are renamed to canonical names.
      - Unknown/extra keys are silently dropped.
      - After normalization, if any of _REQUIRED_CANONICAL fields is missing,
        raises ValueError("Missing required field: <field>").
    """
    # 1. Build normalized dict: canonical key -> value
    out: dict = {}
    for key, value in raw.items():
        canonical = FIELD_ALIASES.get(key, key)   # remap or keep as-is
        if canonical in _REQUIRED_CANONICAL:       # drop anything not in canonical set
            out[canonical] = value

    # 2. Validate all required fields present after normalization
    for field in _REQUIRED_CANONICAL:
        if field not in out:
            raise ValueError(f"Missing required field: {field}")

    return out  # type: ignore[return-value]  -- shape matches WorkbookMetadata


def normalize_metadata(raw: dict) -> "WorkbookMetadata":
    """Public alias for map_lineage_fields. Preferred entry point for callers."""
    return map_lineage_fields(raw)
