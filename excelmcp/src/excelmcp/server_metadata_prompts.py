"""MCP metadata prompts and contract resource.

Import side-effect: importing this module registers 2 prompts and 1 resource
on the shared mcp instance from excelmcp._server_instance.
Do not instantiate FastMCP here.
"""
from __future__ import annotations

import json

from excelmcp._metadata_contract import CANONICAL_FIELDS, CONTRACT_VERSION
from excelmcp._server_instance import mcp


@mcp.resource("excelmcp://meta/contract")
def resource_metadata_contract() -> str:
    """Return the WorkbookMetadata contract schema as JSON — no live workbook required.

    The returned object describes each canonical field: its name, type, and description.
    Use this resource to understand what build_workbook_metadata() will return.
    """
    schema = {
        "contract_version": CONTRACT_VERSION,
        "fields": [
            {"name": "path",             "type": "str",       "description": "Absolute, allowlist-validated path to workbook"},
            {"name": "workbook_name",    "type": "str",       "description": "os.path.basename(path)"},
            {"name": "sheet_names",      "type": "list[str]", "description": "All worksheet names in order"},
            {"name": "active_sheet",     "type": "str",       "description": "wb.active.title"},
            {"name": "sheet_count",      "type": "int",       "description": "len(sheet_names)"},
            {"name": "has_formulas",     "type": "bool",      "description": "True if any cell value starts with '='"},
            {"name": "last_modified",    "type": "str",       "description": "File mtime as ISO-8601 UTC"},
            {"name": "size_bytes",       "type": "int",       "description": "os.path.getsize in bytes"},
            {"name": "format",           "type": "str",       "description": "Suffix: .xlsx | .xlsm | .xls"},
            {"name": "contract_version", "type": "str",       "description": f"Always '{CONTRACT_VERSION}'"},
        ],
        "canonical_field_order": list(CANONICAL_FIELDS),
    }
    return json.dumps(schema, indent=2)


@mcp.prompt()
def workbook_metadata_workflow(path: str) -> str:
    """4-step metadata-first workflow: load -> review -> act on policy -> confirm.

    Uses only canonical field names from CANONICAL_FIELDS. No aliases.
    """
    fields_list = "\n".join(f"  - {f}" for f in CANONICAL_FIELDS)
    return (
        f"## Workbook Metadata Workflow\n\n"
        f"**Path:** `{path}`\n\n"
        f"### Step 1 — Load structural metadata\n"
        f"Call the resource `excelmcp://workbook/{{url_encoded_path}}/metadata` to retrieve:\n"
        f"(where `{{url_encoded_path}}` = `urllib.parse.quote(path, safe='')`)`\n"
        f"{fields_list}\n\n"
        f"### Step 2 — Review metadata fields\n"
        f"Verify: `sheet_count` matches expected sheets, `has_formulas` is correct, "
        f"`last_modified` is recent, `size_bytes` is within expected range.\n\n"
        f"### Step 3 — Act on policy\n"
        f"Check the `policy` field in the result.\n"
        f"- If `strict`: any missing or unreadable field is a hard error — stop and report.\n"
        f"- If `lenient`: optional fields may be empty; required fields (`path`, `workbook_name`) "
        f"must still be present.\n\n"
        f"### Step 4 — Confirm and proceed\n"
        f"Only proceed with subsequent tool calls (read_sheet_all, range_io, etc.) "
        f"after metadata review passes. Report `contract_version` in your summary."
    )


@mcp.prompt()
def metadata_contract_reference() -> str:
    """Return CONTRACT_VERSION, all canonical field names, and their types as markdown."""
    lines = [
        f"# WorkbookMetadata Contract v{CONTRACT_VERSION}",
        "",
        "| Field | Type | Description |",
        "|---|---|---|",
        "| `path` | `str` | Absolute allowlist-validated path |",
        "| `workbook_name` | `str` | `os.path.basename(path)` |",
        "| `sheet_names` | `list[str]` | All worksheet names in workbook order |",
        "| `active_sheet` | `str` | `wb.active.title` |",
        "| `sheet_count` | `int` | `len(sheet_names)` |",
        "| `has_formulas` | `bool` | Any cell value string starting with `=` |",
        "| `last_modified` | `str` | File mtime as ISO-8601 UTC |",
        "| `size_bytes` | `int` | `os.path.getsize()` |",
        "| `format` | `str` | `.xlsx` \\| `.xlsm` \\| `.xls` |",
        "| `contract_version` | `str` | Always `\"1.0\"` |",
        "",
        f"**Canonical field order:** `{'`, `'.join(CANONICAL_FIELDS)}`",
        "",
        "**Policy key** (injected by server layer, not in TypedDict):",
        "| `policy` | `str` | `strict` \\| `lenient` — from `EXCEL_METADATA_POLICY` env |",
    ]
    return "\n".join(lines)
