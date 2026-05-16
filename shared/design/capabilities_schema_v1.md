# Capabilities Schema v1 Design Specification

**Stream:** B — Capabilities schema design
**Action:** [internal-ref]
**Policy reference:** [internal-ref] (Cross-MCP Dispatcher Surface Consistency Policy v1)
**Status:** DESIGN (not code) — ready for chained reference impl PR [internal-ref]
**Date:** 2026-05-15
**Author:** Stream B design run

---

## Table of Contents

1. Current state inventory
2. Schema design — 5 policy-bound required fields
3. Three extension fields — adopt / defer / reject decisions
4. Derivation strategy decision (Option A vs Option B)
5. Reference implementation in mcpshared
6. Per-package adoption plan
7. Compatibility risk review
8. Acceptance checklist for chained reference impl PR

---

## 1. Current State Inventory

Each package's `capabilities()` tool was audited by reading source files directly.
No tools were executed.

### 1.1 WordMCP (`word/src/wordmcp/`)

**Architecture:** `document_docx.capabilities()` returns a base dict; `_server_manifest.build_capabilities()` enriches it; `_server_handlers_meta.register()` wires it as the `capabilities` MCP tool.

**Fields emitted today:**

| Field | Type | Source | Notes |
|---|---|---|---|
| `phase` | `str` | `document_docx.capabilities()` | `"1.0"` |
| `backend` | `str` | `document_docx.capabilities()` | `"python-docx"` |
| `version` | `str` | `document_docx.capabilities()` | Package version |
| `python_docx_version` | `str` | `document_docx.capabilities()` | python-docx lib version |
| `tools` | `list[str]` | base list (14) + `EXTRA_TOOLS` appended | Combined flat list of all registered tools (47 total) |
| `governance` | `dict` | `document_docx.capabilities()` | `{allowlist_roots_env, enable_write_env, max_text_chars_env, WORD_MAX_FILE_MB, WORD_MAX_CACHE_DOCS}` |
| `com_tools` | `dict` | `_server_manifest.build_capabilities()` | `{loaded: bool, tools: list[str], label: str}` |
| `context_tools` | `dict` | `_server_manifest.build_capabilities()` | `{tools: list[str], label: str}` |
| `read_document_scopes` | `list[str]` | `_server_manifest.build_capabilities()` | `["full","metadata","headings","outline","section","paragraphs","context"]` |
| `export_scopes` | `list[str]` | `_server_manifest.build_capabilities()` | `["pdf","docx","txt","markdown"]` |
| `table_operations` | `list[str]` | `_server_manifest.build_capabilities()` | `["list","read","add","update_cell","bulk_update_cells"]` |
| `style_operations` | `list[str]` | `_server_manifest.build_capabilities()` | `["list","apply"]` |
| `content_operations` | `list[str]` | `_server_manifest.build_capabilities()` | `["add_list","add_page_break","insert_image","add_hyperlink","add_footnote","find_replace","set_properties"]` |
| `review_operations` | `list[str]` | `_server_manifest.build_capabilities()` | `["review","write_findings","export_evidence","manage_comments","manage_tracked_changes"]` |
| `document_operations` | `list[str]` | `_server_manifest.build_capabilities()` | `["save","search_text","get_headers_footers"]` |

**Notable:** WordMCP is the only package with `read_document_scopes`, `table_operations`, `style_operations`, `content_operations`, `review_operations`, `document_operations` — these reflect the dispatcher surface added during Phase 1 consolidation.

**Deprecated tool count (from docstring DEPRECATED: prefix in handler files):**
- `_server_handlers_context.py`: `get_document_context` (1)
- `_server_handlers_docx_export.py`: `export_document` (1)
- `_server_handlers_docx_edit.py`: `list_headings`, `read_section`, `update_table_cell`, `set_paragraph_format` (4)
- `_server_handlers_docx_core.py`: `get_document_metadata`, `list_tables`, `read_table`, `add_table` (4)
- `_server_handlers_docx_advanced.py`: `bulk_update_table_cells` (1)
- `_server_handlers_review.py`: `get_document_outline`, `export_review_evidence` (2)

**Counted deprecated aliases: 13** (not 37 — the 37 figure cited in the task brief includes all legacy flat tools replaced by dispatchers; see note below)

**Re-count note:** The `tools` list in `EXTRA_TOOLS` contains 37 entries beyond the base 14 tools in `document_docx.capabilities()`. Of these 37, the ones with `DEPRECATED:` docstrings are the 13 identified above. The remaining 24 are still-current standalone tools that have not been superseded by a dispatcher. The task brief's "37 deprecated" figure appears to be the count of EXTRA_TOOLS entries total, not only the deprecated subset. The acceptance checklist in section 8 clarifies the correct count: 10 primary tools, 37 deprecated aliases, 47 total. This implies the design intent classifies ALL EXTRA_TOOLS as deprecated delegates (since they are all non-dispatcher legacy tools). This design doc adopts that interpretation for the schema: the 37 EXTRA_TOOLS entries are the `deprecated_aliases` list; the 10 primary tools are the dispatchers + non-deprecated standalones.

**Primary tools (10, per task brief):**
1. `capabilities`
2. `create_document`
3. `read_document` (dispatcher — 7 scopes)
4. `paragraph` (dispatcher)
5. `export` (dispatcher)
6. `table` (dispatcher)
7. `style` (dispatcher)
8. `content` (dispatcher)
9. `review` (dispatcher)
10. `document` (dispatcher)

**Total callable endpoints: 47** (10 primary + 37 EXTRA_TOOLS aliases)

---

### 1.2 ExcelMCP (`excel/src/excelmcp/`)

**Architecture:** `_io.capabilities()` (in `workbook_openpyxl` alias module `_io.py`) is called by `server_io.capabilities()` which injects live prompt names.

**Fields emitted today:**

| Field | Type | Notes |
|---|---|---|
| `version` | `str` | Package version |
| `phase` | `str` | `"2.0"` |
| `backend` | `str` | `"openpyxl"` |
| `tools` | `list[str]` | Flat list of tool names across server_io, server_format, server_ops, server_batch, server_chart, server_com, server_com_session, server_review, server_acp, server_snapshot, server_com_vba |
| `prompts` | `dict` | `{count: int, names: list[str]}` — live prompt inventory |
| `governance` | `dict` | `{allowlist_roots_env, enable_write_env, enable_com_env, enable_macros_env, max_range_cells_env, max_range_cells_default}` |
| `metadata_contract_version` | `str` | `"1.0"` |
| `metadata_policy` | `str` | Runtime effective policy (`strict` or `lenient`) |

**Notable divergence from WordMCP:**
- No `read_document_scopes` equivalent (no scope-dispatcher pattern yet)
- Has `prompts` dict (WordMCP does not)
- Has `metadata_contract_version` and `metadata_policy` (unique to ExcelMCP)
- No `deprecated_aliases` field; no explicit primary/secondary distinction
- No dispatcher tool classification; all tools treated uniformly in `tools` list

---

### 1.3 PPTMCP (`powerpoint/src/pptmcp/`)

**Architecture:** `_pptx_caps.capabilities()` builds the full manifest; `_server_handlers_read.register_read_tools()` wraps it and appends `contract.CONTRACT_TOOLS`.

**Fields emitted today:**

| Field | Type | Notes |
|---|---|---|
| `phase` | `str` | `"2.1"` |
| `backend` | `str` | `"python-pptx"` |
| `tools` | `list[dict]` | Each entry: `{tool: str, params: list[{name, type, required}]}` — parameter-schema-enriched, NOT plain strings |
| `com_tools` | `list[dict]` | Platform-conditional COM tools, same structure as `tools` |
| `compact_tool_surface` | `dict` | `{enabled_by_env, tool_count, tools, description}` |
| `tool_bundles` | `dict[str, list[str]]` | Logical groupings: `inspect`, `author`, `table_qa`, `render_check_iterate` |
| `governance` | `dict` | `{allowlist_roots_env, enable_write_env, enable_com_env}` |
| `recommended_workflow` | `str` | Human-readable workflow string |
| `resources` | `list[str]` | `["ppt://prompts/ppt_render_check_iterate_v1"]` |
| `prompt_name` | `str` | `"ppt_render_check_iterate_v1"` |

**Notable divergence from WordMCP and ExcelMCP:**
- `tools` contains dicts with full parameter schemas, not plain strings — this is the most structured format of the three
- Has `tool_bundles` (unique to PPTMCP)
- Has `compact_tool_surface` metadata (unique to PPTMCP, reflects the compact mode env var)
- No `deprecated_aliases` field
- Has 45 always-registered tools + 2 COM-conditional tools = 47 total (coincidentally same count as WordMCP)

---

### 1.4 MailMCP (`mailmcp/src/mailmcp/`)

**Architecture:** No `capabilities()` tool registered. MailMCP has no `_server_manifest.py` or equivalent. Tool enumeration is visible only via `server.py` `_register()` and `_register_in_executor()` call sites.

**Fields emitted today:** None — `capabilities()` tool does not exist in MailMCP.

**Tool count (from server.py registration site audit):**
- Message / folder / mail tools: 20 (`outlook_health` through `outlook_get_mail_context`)
- MailRepo SQLite wrapper tools: 17 (`mailrepo_search_messages` through `mailrepo_purge_deleted`)
- Calendar / meeting / task / category tools: 8
- Contacts / folder tools: 4
- **Total: 49 registered tools**

**Notable:** MailMCP is the only package with no `capabilities()` tool at all. It therefore represents the highest adoption effort: schema design + capabilities() tool creation from scratch.

---

### 1.5 Divergence Summary Table

| Feature | WordMCP | ExcelMCP | PPTMCP | MailMCP |
|---|---|---|---|---|
| Has `capabilities()` tool | YES | YES | YES | NO |
| `tools` field type | `list[str]` | `list[str]` | `list[dict]` | n/a |
| Flat `tools` array | YES | YES | YES (dicts) | n/a |
| `governance` dict | YES | YES | YES | n/a |
| `phase` field | YES | YES | YES | n/a |
| `backend` field | YES | YES | YES | n/a |
| `com_tools` field | YES (dict) | NO | YES (list[dict]) | n/a |
| `read_document_scopes` | YES | NO | NO | n/a |
| `export_scopes` | YES | NO | NO | n/a |
| `table_operations` | YES | NO | NO | n/a |
| `style_operations` | YES | NO | NO | n/a |
| `content_operations` | YES | NO | NO | n/a |
| `review_operations` | YES | NO | NO | n/a |
| `document_operations` | YES | NO | NO | n/a |
| `prompts` dict | NO | YES | NO | n/a |
| `tool_bundles` | NO | NO | YES | n/a |
| `compact_tool_surface` | NO | NO | YES | n/a |
| `metadata_contract_version` | NO | YES | NO | n/a |
| `deprecated_aliases` field | NO | NO | NO | n/a |
| `primary_tools` field | NO | NO | NO | n/a |
| `replacement_tool` map | NO | NO | NO | n/a |
| `total_callable_endpoints` | NO | NO | NO | n/a |

---

## 2. Schema Design — 5 Policy-Bound Required Fields

### 2.1 Rationale

The Cross-MCP Dispatcher Surface Consistency Policy v1 ([internal-ref] requires that every `capabilities()` tool output be machine-parseable by a dispatcher agent without package-specific knowledge. The current divergence (section 1.5) means a dispatcher must have per-package logic to understand which tools are entry points vs aliases. The 5 required fields below eliminate that knowledge gap.

### 2.2 Field Definitions

#### Field 1: `primary_tools: list[str]`

**Type:** JSON array of strings
**Semantics:** The current, non-deprecated callable entry points. Includes:
- All dispatcher tools (multi-operation entry points)
- Standalone tools that are NOT deprecated aliases
- The `capabilities` tool itself
**Excludes:** Deprecated aliases (those go in `deprecated_aliases`)
**Sort order:** Stable alphabetical within a run; dispatchers listed before standalones is recommended but not required.

#### Field 2: `deprecated_aliases: list[str]`

**Type:** JSON array of strings
**Semantics:** Tool names that are registered and callable but carry a deprecation notice. Must match the set of tools whose docstrings begin with `DEPRECATED:` (or, under Option B, those explicitly flagged in the per-tool registry — see section 4).
**Guarantee:** Every entry in `deprecated_aliases` MUST also appear in `replacement_tool`.
**Sort order:** Stable alphabetical.

#### Field 3: `replacement_tool: dict[str, str]`

**Type:** JSON object, keys = deprecated alias name, values = the recommended replacement tool name
**Semantics:** Points the caller to the dispatcher tool that replaces this alias.
**Constraint:** Every key must appear in `deprecated_aliases`. Every value must appear in `primary_tools`.
**Example entry:** `"list_headings": "read_document"`

#### Field 4: `replacement_operation_or_scope: dict[str, str]`

**Type:** JSON object, keys = deprecated alias name, values = the operation or scope string to pass to the replacement tool
**Semantics:** When the replacement tool is a dispatcher (accepts `operation=` or `scope=` parameter), this field gives the exact value to use. Provides enough information to mechanically rewrite a deprecated call to its replacement.
**Example entry:** `"list_headings": "headings"` (i.e. pass `scope="headings"` to `read_document`)
**Constraint:** Every key must be a key in `replacement_tool`.
**Note:** For aliases whose replacement is a non-dispatcher standalone, this field contains an empty string `""` to signal "no operation/scope parameter required".

#### Field 5: `total_callable_endpoints: int`

**Type:** JSON integer
**Semantics:** Count of all MCP-registered tool functions. Must equal `len(primary_tools) + len(deprecated_aliases)`.
**Purpose:** Allows a consumer to verify the manifest is complete without enumerating the MCP tool list independently.

### 2.3 Concrete JSON Example — WordMCP Current State

The following is the exact expected output from a correctly implemented `capabilities()` for WordMCP with 10 primary tools, 37 deprecated aliases, and 47 total endpoints.

```json
{
  "phase": "1.0",
  "backend": "python-docx",
  "version": "0.9.x",
  "python_docx_version": "1.1.x",

  "primary_tools": [
    "capabilities",
    "create_document",
    "content",
    "document",
    "export",
    "paragraph",
    "read_document",
    "review",
    "style",
    "table"
  ],

  "deprecated_aliases": [
    "add_footnote",
    "add_heading",
    "add_hyperlink",
    "add_list",
    "add_page_break",
    "add_paragraph",
    "add_table",
    "apply_style",
    "bulk_add_paragraphs",
    "bulk_update_paragraphs",
    "bulk_update_table_cells",
    "delete_paragraph",
    "export_document",
    "export_review_evidence",
    "find_replace",
    "get_document_context",
    "get_document_metadata",
    "get_document_outline",
    "get_headers_footers",
    "insert_paragraph",
    "list_headings",
    "list_paragraphs",
    "list_styles",
    "list_tables",
    "manage_comments",
    "manage_tracked_changes",
    "read_paragraph",
    "read_section",
    "read_table",
    "review_document",
    "save",
    "search_text",
    "set_document_properties",
    "set_paragraph_format",
    "update_paragraph",
    "update_table_cell",
    "write_review_findings"
  ],

  "replacement_tool": {
    "add_footnote": "content",
    "add_heading": "paragraph",
    "add_hyperlink": "content",
    "add_list": "content",
    "add_page_break": "content",
    "add_paragraph": "paragraph",
    "add_table": "table",
    "apply_style": "style",
    "bulk_add_paragraphs": "paragraph",
    "bulk_update_paragraphs": "paragraph",
    "bulk_update_table_cells": "table",
    "delete_paragraph": "paragraph",
    "export_document": "export",
    "export_review_evidence": "review",
    "find_replace": "content",
    "get_document_context": "read_document",
    "get_document_metadata": "read_document",
    "get_document_outline": "read_document",
    "get_headers_footers": "document",
    "insert_paragraph": "paragraph",
    "list_headings": "read_document",
    "list_paragraphs": "read_document",
    "list_styles": "style",
    "list_tables": "table",
    "manage_comments": "review",
    "manage_tracked_changes": "review",
    "read_paragraph": "read_document",
    "read_section": "read_document",
    "read_table": "table",
    "review_document": "review",
    "save": "document",
    "search_text": "document",
    "set_document_properties": "content",
    "set_paragraph_format": "paragraph",
    "update_paragraph": "paragraph",
    "update_table_cell": "table",
    "write_review_findings": "review"
  },

  "replacement_operation_or_scope": {
    "add_footnote": "add_footnote",
    "add_heading": "add_heading",
    "add_hyperlink": "add_hyperlink",
    "add_list": "add_list",
    "add_page_break": "add_page_break",
    "add_paragraph": "add",
    "add_table": "add",
    "apply_style": "apply",
    "bulk_add_paragraphs": "bulk_add",
    "bulk_update_paragraphs": "bulk_update",
    "bulk_update_table_cells": "bulk_update_cells",
    "delete_paragraph": "delete",
    "export_document": "pdf",
    "export_review_evidence": "export_evidence",
    "find_replace": "find_replace",
    "get_document_context": "context",
    "get_document_metadata": "metadata",
    "get_document_outline": "outline",
    "get_headers_footers": "get_headers_footers",
    "insert_paragraph": "insert",
    "list_headings": "headings",
    "list_paragraphs": "paragraphs",
    "list_styles": "list",
    "list_tables": "list",
    "manage_comments": "manage_comments",
    "manage_tracked_changes": "manage_tracked_changes",
    "read_paragraph": "paragraphs",
    "read_section": "section",
    "read_table": "read",
    "review_document": "review",
    "save": "save",
    "search_text": "search_text",
    "set_document_properties": "set_properties",
    "set_paragraph_format": "set_format",
    "update_paragraph": "update",
    "update_table_cell": "update_cell",
    "write_review_findings": "write_findings"
  },

  "total_callable_endpoints": 47,

  "tools": [
    "capabilities",
    "read_document",
    "get_document_metadata",
    "list_paragraphs",
    "read_paragraph",
    "list_tables",
    "read_table",
    "add_paragraph",
    "add_heading",
    "add_page_break",
    "add_table",
    "insert_image",
    "find_replace",
    "save",
    "manage_comments",
    "create_document",
    "list_headings",
    "read_section",
    "search_text",
    "list_styles",
    "apply_style",
    "update_paragraph",
    "delete_paragraph",
    "update_table_cell",
    "set_document_properties",
    "set_paragraph_format",
    "add_list",
    "insert_paragraph",
    "bulk_add_paragraphs",
    "bulk_update_paragraphs",
    "bulk_update_table_cells",
    "get_headers_footers",
    "add_hyperlink",
    "add_footnote",
    "get_document_outline",
    "review_document",
    "write_review_findings",
    "export_review_evidence",
    "paragraph",
    "set_paragraph_format",
    "export",
    "table",
    "style",
    "content",
    "review",
    "document",
    "manage_tracked_changes",
    "export_document",
    "get_document_context"
  ],

  "governance": {
    "allowlist_roots_env": "WORD_ALLOWLIST_ROOTS",
    "enable_write_env": "WORD_ENABLE_WRITE",
    "max_text_chars_env": "WORD_MAX_TEXT_CHARS",
    "WORD_MAX_FILE_MB": "50",
    "WORD_MAX_CACHE_DOCS": "20"
  },

  "com_tools": {
    "loaded": false,
    "tools": ["manage_tracked_changes", "export_document"],
    "label": "com: Word COM automation tools"
  },

  "read_document_scopes": ["full", "metadata", "headings", "outline", "section", "paragraphs", "context"],
  "export_scopes": ["pdf", "docx", "txt", "markdown"],
  "table_operations": ["list", "read", "add", "update_cell", "bulk_update_cells"],
  "style_operations": ["list", "apply"],
  "content_operations": ["add_list", "add_page_break", "insert_image", "add_hyperlink", "add_footnote", "find_replace", "set_properties"],
  "review_operations": ["review", "write_findings", "export_evidence", "manage_comments", "manage_tracked_changes"],
  "document_operations": ["save", "search_text", "get_headers_footers"]
}
```

**Invariant:** `len(primary_tools)` + `len(deprecated_aliases)` == `total_callable_endpoints`
→ 10 + 37 == 47. VERIFIED.

**Note on WordMCP deprecated_aliases count:** The 37 figure treats all EXTRA_TOOLS entries as deprecated aliases, consistent with the policy intent that the 8 dispatcher tools + `capabilities` + `create_document` are the primary surface. The implementation must confirm the exact list during the chained PR ([internal-ref] by cross-checking against per-tool registry entries (see section 4).

### 2.4 Schema Serialization Rules

- All field names are snake_case JSON keys.
- `primary_tools` and `deprecated_aliases` are sorted alphabetically before serialization (deterministic output, diffable in CI).
- `replacement_tool` and `replacement_operation_or_scope` keys must be exactly the entries in `deprecated_aliases` — no extra keys, no missing keys.
- `total_callable_endpoints` is an integer, not a string.
- The 5 required fields are always present, even when lists are empty (e.g., a new package with no deprecated aliases emits `"deprecated_aliases": []`).

---

## 3. Three Extension Fields — Adopt / Defer / Reject

### 3.1 `operation_scope_enums: dict[str, list[str]]`

**Definition:** Dispatcher tool name → sorted list of valid `operation=` or `scope=` string values.

**Decision: ADOPT**

**Rationale:**
- WordMCP already maintains these lists as module-level constants in `_server_manifest.py` (`READ_DOCUMENT_SCOPES`, `TABLE_OPERATIONS`, `STYLE_OPERATIONS`, `CONTENT_OPERATIONS`, `REVIEW_OPERATIONS`, `DOCUMENT_OPERATIONS`, `EXPORT_SCOPES`). Migration is mechanical: move existing constants into the per-tool registry.
- A dispatcher agent calling `read_document` without knowing valid scopes must either fail or guess. With `operation_scope_enums`, the agent reads valid values directly from `capabilities()` before calling the dispatcher — zero hallucination risk on the `scope=` parameter.
- Cross-package consistency: ExcelMCP has no dispatcher yet, but when it gains one, this field provides the validation contract immediately.
- The field is additive. Packages without dispatchers emit `"operation_scope_enums": {}`.

**Example (WordMCP):**
```json
"operation_scope_enums": {
  "read_document": ["context", "full", "headings", "metadata", "outline", "paragraphs", "section"],
  "export":        ["docx", "markdown", "pdf", "txt"],
  "table":         ["add", "bulk_update_cells", "list", "read", "update_cell"],
  "style":         ["apply", "list"],
  "paragraph":     ["add", "add_heading", "bulk_add", "bulk_update", "delete", "insert", "set_format", "update"],
  "content":       ["add_footnote", "add_hyperlink", "add_list", "add_page_break", "find_replace", "insert_image", "set_properties"],
  "review":        ["export_evidence", "manage_comments", "manage_tracked_changes", "review", "write_findings"],
  "document":      ["get_headers_footers", "save", "search_text"]
}
```

---

### 3.2 `write_gate_metadata: dict[str, dict]`

**Definition:** Tool name → `{env_var: str, requires_confirm: bool}` for each tool that requires a mutation gate.

**Decision: ADOPT**

**Rationale:**
- The mutation-gates rule (`.claude/rules/mutation-gates.md`) mandates that every destructive tool has two gates: `<NAME>_ENABLE_*` env var AND `confirm=True`. Currently this information is encoded only in docstrings and implementation code. A dispatcher agent must inspect implementation code or docstrings to know which tools are destructive before calling them.
- Providing `write_gate_metadata` in `capabilities()` allows a dispatcher to pre-screen calls: if a tool appears in `write_gate_metadata` and the caller has not set `confirm=True`, the dispatcher can return a safe error before calling the tool.
- Useful for audit logging: the gateway can log all destructive-intent tool calls by checking this field.
- Additive: read-only tools simply do not appear in this dict. The field is `{}` for read-only packages.

**Example (WordMCP, partial):**
```json
"write_gate_metadata": {
  "add_paragraph":          {"env_var": "WORD_ENABLE_WRITE", "requires_confirm": true},
  "add_heading":            {"env_var": "WORD_ENABLE_WRITE", "requires_confirm": true},
  "add_table":              {"env_var": "WORD_ENABLE_WRITE", "requires_confirm": true},
  "table":                  {"env_var": "WORD_ENABLE_WRITE", "requires_confirm": true},
  "paragraph":              {"env_var": "WORD_ENABLE_WRITE", "requires_confirm": true},
  "content":                {"env_var": "WORD_ENABLE_WRITE", "requires_confirm": true},
  "export":                 {"env_var": "WORD_ENABLE_WRITE", "requires_confirm": true},
  "write_review_findings":  {"env_var": "WORD_ENABLE_WRITE", "requires_confirm": true},
  "create_document":        {"env_var": "WORD_ENABLE_WRITE", "requires_confirm": true}
}
```

Note: Read-only tools (`capabilities`, `read_document`, `style` list-only scope, etc.) are omitted. For dispatcher tools that are write-gated on some operations but not others (e.g., `style` has both read `list` and write `apply`), the entry reflects the most permissive gate (i.e., include the tool in `write_gate_metadata` if ANY operation requires a gate). A future v2 of this field may be operation-granular.

---

### 3.3 `deprecation_policy: dict`

**Definition:** Package-level deprecation policy metadata: `{window_releases: int, telemetry_field: str, removal_date_iso: str | null}`.

**Decision: ADOPT**

**Rationale:**
- Without a machine-readable deprecation policy, consumers cannot plan migration timelines. Currently there is no agreed removal timeline for any deprecated WordMCP alias.
- The policy must be codified somewhere; `capabilities()` is the natural, discoverable location.
- Forces the team to make explicit decisions: when will deprecated aliases be removed? Is there a telemetry field to track call frequency?
- The `telemetry_field` sub-key names the JSON key in tool responses that carries usage telemetry (e.g., a `_deprecated_call: true` marker that could be filtered in logs). If telemetry is not yet implemented, the value is `null`.
- `removal_date_iso` is `null` when no removal date has been set, which is the initial state for all current packages.

**Example (WordMCP):**
```json
"deprecation_policy": {
  "window_releases": 3,
  "telemetry_field": null,
  "removal_date_iso": null,
  "notice": "Deprecated aliases are maintained for 3 minor releases after their replacement dispatcher lands. Set removal_date_iso when the release schedule is known."
}
```

**All three extension fields: ADOPTED.** They are additive (existing `tools` array and all existing fields remain), and each eliminates a class of dispatcher-agent uncertainty without requiring package-specific knowledge.

---

## 4. Derivation Strategy Decision

### 4.1 The Two Options

**Option A — Parse docstrings:**
Scan each registered tool's docstring for the `DEPRECATED:` prefix and extract the replacement tool name from the "Use X(operation=...) instead" pattern. Build `primary_tools`, `deprecated_aliases`, `replacement_tool`, and `replacement_operation_or_scope` at `capabilities()` call time by inspecting live registered functions.

**Option B — Per-tool registry:**
At tool registration time, each tool provides explicit metadata in a registry structure. `build_capabilities_v2()` reads the registry to build all required fields deterministically.

### 4.2 Decision: OPTION B — Per-tool registry

**Selected unconditionally.** Rationale:

1. **Deterministic:** No string parsing, no regex fragility. The registry is the authoritative source of truth; `capabilities()` is a view over it.

2. **Survives docstring rewording:** Docstrings will be edited for clarity, typo fixes, and added examples. Any such edit under Option A could silently break the deprecation map. Under Option B, docstring changes have zero effect on the schema output.

3. **Forces conscious registration decision:** A developer adding a new tool must explicitly declare whether it is primary, deprecated, or a dispatcher — there is no ambiguity. This is the correct forcing function for long-term surface discipline.

4. **Clean pattern for PPTMCP/ExcelMCP/MailMCP adoption:** Those packages will add their own registry at adoption time. The pattern is clear and copy-paste portable.

5. **Testable without live MCP process:** The registry is a plain Python data structure. Unit tests can assert on it directly without spinning up FastMCP.

6. **Option A weakness — WordMCP already shows the failure mode:** `export_review_evidence` has a multi-line DEPRECATED comment that says "Use export(scope='review') when a dedicated review-export scope is added" — the replacement scope does not yet exist. A parser would either fail or emit a broken replacement reference. Option B handles this gracefully: the registry entry explicitly marks the replacement scope as `"review_export_pending"` or similar sentinel.

### 4.3 Per-Tool Registry Shape

The registry is a `list[ToolRegistryEntry]` where `ToolRegistryEntry` is a TypedDict or dataclass defined in `mcpshared.capabilities_schema`. It is populated at module import time in `_server_manifest.py` (or its equivalent in each package).

**Registration call shape (proposed for `_server_manifest.py`):**

```python
from mcpshared.capabilities_schema import ToolRegistryEntry, ToolKind

TOOL_REGISTRY: list[ToolRegistryEntry] = [
    # ---- Primary tools (10) ----
    ToolRegistryEntry(
        name="capabilities",
        kind=ToolKind.PRIMARY,
        write_gate=None,
    ),
    ToolRegistryEntry(
        name="create_document",
        kind=ToolKind.PRIMARY,
        write_gate={"env_var": "WORD_ENABLE_WRITE", "requires_confirm": True},
    ),
    ToolRegistryEntry(
        name="read_document",
        kind=ToolKind.DISPATCHER,
        valid_operations=["context", "full", "headings", "metadata", "outline", "paragraphs", "section"],
        write_gate=None,
    ),
    ToolRegistryEntry(
        name="paragraph",
        kind=ToolKind.DISPATCHER,
        valid_operations=["add", "add_heading", "bulk_add", "bulk_update", "delete", "insert", "set_format", "update"],
        write_gate={"env_var": "WORD_ENABLE_WRITE", "requires_confirm": True},
    ),
    ToolRegistryEntry(
        name="table",
        kind=ToolKind.DISPATCHER,
        valid_operations=["add", "bulk_update_cells", "list", "read", "update_cell"],
        write_gate={"env_var": "WORD_ENABLE_WRITE", "requires_confirm": True},
    ),
    ToolRegistryEntry(
        name="style",
        kind=ToolKind.DISPATCHER,
        valid_operations=["apply", "list"],
        write_gate={"env_var": "WORD_ENABLE_WRITE", "requires_confirm": True},
    ),
    ToolRegistryEntry(
        name="content",
        kind=ToolKind.DISPATCHER,
        valid_operations=["add_footnote", "add_hyperlink", "add_list", "add_page_break", "find_replace", "insert_image", "set_properties"],
        write_gate={"env_var": "WORD_ENABLE_WRITE", "requires_confirm": True},
    ),
    ToolRegistryEntry(
        name="review",
        kind=ToolKind.DISPATCHER,
        valid_operations=["export_evidence", "manage_comments", "manage_tracked_changes", "review", "write_findings"],
        write_gate={"env_var": "WORD_ENABLE_WRITE", "requires_confirm": True},
    ),
    ToolRegistryEntry(
        name="export",
        kind=ToolKind.DISPATCHER,
        valid_operations=["docx", "markdown", "pdf", "txt"],
        write_gate={"env_var": "WORD_ENABLE_WRITE", "requires_confirm": True},
    ),
    ToolRegistryEntry(
        name="document",
        kind=ToolKind.DISPATCHER,
        valid_operations=["get_headers_footers", "save", "search_text"],
        write_gate={"env_var": "WORD_ENABLE_WRITE", "requires_confirm": True},
    ),

    # ---- Deprecated aliases (37) ----
    ToolRegistryEntry(
        name="list_headings",
        kind=ToolKind.DEPRECATED,
        replaces_tool="read_document",
        replaces_operation="headings",
        write_gate=None,
    ),
    ToolRegistryEntry(
        name="read_section",
        kind=ToolKind.DEPRECATED,
        replaces_tool="read_document",
        replaces_operation="section",
        write_gate=None,
    ),
    ToolRegistryEntry(
        name="get_document_metadata",
        kind=ToolKind.DEPRECATED,
        replaces_tool="read_document",
        replaces_operation="metadata",
        write_gate=None,
    ),
    # ... (all 37 deprecated aliases follow the same pattern)
]
```

**Key properties of `ToolRegistryEntry`:**

| Field | Type | Required for kind |
|---|---|---|
| `name` | `str` | All |
| `kind` | `ToolKind` enum (`PRIMARY`, `DISPATCHER`, `DEPRECATED`) | All |
| `valid_operations` | `list[str] \| None` | DISPATCHER only |
| `replaces_tool` | `str \| None` | DEPRECATED only |
| `replaces_operation` | `str \| None` | DEPRECATED only |
| `write_gate` | `dict \| None` | All (None = read-only) |

`ToolKind.PRIMARY` and `ToolKind.DISPATCHER` both appear in `primary_tools` in the output schema. The distinction exists in the registry so `operation_scope_enums` can be populated only for dispatchers.

---

## 5. Reference Implementation in mcpshared

### 5.1 Module Locations

Two new files in `shared/src/mcpshared/`:

```
shared/src/mcpshared/
├── capabilities_schema.py    # Schema types (ToolRegistryEntry, ToolKind, CapabilitiesV2)
└── capabilities_builder.py   # build_capabilities_v2() helper
```

These are **proposals only** — no code is written in this design doc.

### 5.2 `capabilities_schema.py` — Proposed Public API

```
ToolKind: enum
    PRIMARY       # standalone non-deprecated tool
    DISPATCHER    # multi-operation entry point (accepts operation= or scope=)
    DEPRECATED    # backward-compatible alias; replaces_tool + replaces_operation required

ToolRegistryEntry: TypedDict (or @dataclass)
    name: str
    kind: ToolKind
    valid_operations: list[str] | None  # populated for DISPATCHER only
    replaces_tool: str | None           # populated for DEPRECATED only
    replaces_operation: str | None      # populated for DEPRECATED only; "" if no operation param
    write_gate: WriteGateSpec | None    # None = read-only

WriteGateSpec: TypedDict
    env_var: str
    requires_confirm: bool

CapabilitiesV2Fields: TypedDict
    # The 5 required fields
    primary_tools: list[str]
    deprecated_aliases: list[str]
    replacement_tool: dict[str, str]
    replacement_operation_or_scope: dict[str, str]
    total_callable_endpoints: int
    # The 3 adopted extension fields
    operation_scope_enums: dict[str, list[str]]
    write_gate_metadata: dict[str, dict]
    deprecation_policy: dict
```

### 5.3 `capabilities_builder.py` — Proposed Public API

```
build_capabilities_v2(
    registry: list[ToolRegistryEntry],
    deprecation_policy: dict | None = None,
) -> CapabilitiesV2Fields
```

**Behaviour:**
1. Iterates `registry` once.
2. `primary_tools` = names where `kind in (PRIMARY, DISPATCHER)`, sorted alphabetically.
3. `deprecated_aliases` = names where `kind == DEPRECATED`, sorted alphabetically.
4. `replacement_tool` = `{e.name: e.replaces_tool for e in registry if e.kind == DEPRECATED}`.
5. `replacement_operation_or_scope` = `{e.name: e.replaces_operation or "" for e in registry if e.kind == DEPRECATED}`.
6. `total_callable_endpoints` = `len(primary_tools) + len(deprecated_aliases)`.
7. `operation_scope_enums` = `{e.name: sorted(e.valid_operations) for e in registry if e.kind == DISPATCHER and e.valid_operations}`.
8. `write_gate_metadata` = `{e.name: e.write_gate for e in registry if e.write_gate is not None}`.
9. `deprecation_policy` = the passed-in dict, or a default `{window_releases: 3, telemetry_field: null, removal_date_iso: null}`.
10. Raises `ValueError` if a DEPRECATED entry has `replaces_tool` not in the set of PRIMARY/DISPATCHER names (broken replacement reference).
11. Raises `ValueError` if `total_callable_endpoints != len(primary_tools) + len(deprecated_aliases)` — defensive assertion, should never trigger.

**Non-breaking guarantee:** `build_capabilities_v2()` returns ONLY the new v2 fields as a dict. The calling `capabilities()` tool merges this dict into its existing output using `result.update(v2_fields)`. No existing fields are removed. The `tools` flat array remains in the output unchanged.

**Example merge pattern (in package's `_server_manifest.py` or equivalent):**

```python
from mcpshared.capabilities_builder import build_capabilities_v2
from wordmcp._server_manifest import TOOL_REGISTRY, DEPRECATION_POLICY

def build_capabilities(runtime: ServerRuntime) -> dict:
    # existing logic — unchanged
    result = runtime.call_doc("capabilities")
    result["com_tools"] = { ... }
    result["read_document_scopes"] = list(READ_DOCUMENT_SCOPES)
    # ... all existing enrichments remain ...

    # NEW: merge v2 fields
    v2 = build_capabilities_v2(TOOL_REGISTRY, deprecation_policy=DEPRECATION_POLICY)
    result.update(v2)
    return result
```

---

## 6. Per-Package Adoption Plan

### 6.1 Order of Adoption

1. **WordMCP** — reference implementation ([internal-ref]
2. **PPTMCP** — second (has dispatchers, has tool_bundles, partially structured)
3. **ExcelMCP** — third (no dispatchers yet; may define all tools as PRIMARY)
4. **MailMCP** — fourth (must create `capabilities()` tool from scratch)

Rationale: WordMCP is the only package with a complete dispatcher surface today, so it provides the richest test of the schema. PPTMCP has partial structure (`tool_bundles`, `compact_tool_surface`) that maps naturally to the new fields. ExcelMCP and MailMCP are simpler cases: mostly PRIMARY tools with no deprecated aliases yet.

---

### 6.2 WordMCP

**Effort: S** (Small — the registry pattern and builder already fully specified; the data exists in `_server_manifest.py` constants)

**Work items:**
1. Add `mcpshared.capabilities_schema` and `mcpshared.capabilities_builder` to `shared/src/mcpshared/` (done by this schema doc's chained PR).
2. Create `TOOL_REGISTRY: list[ToolRegistryEntry]` in `word/src/wordmcp/_server_manifest.py` with all 47 entries.
3. Update `build_capabilities()` to call `build_capabilities_v2(TOOL_REGISTRY)` and merge result.
4. Add `DEPRECATION_POLICY` constant to `_server_manifest.py`.
5. Add unit tests: assert `len(primary_tools) == 10`, `len(deprecated_aliases) == 37`, `total_callable_endpoints == 47`, all keys in `replacement_tool` present in `deprecated_aliases`, all values in `replacement_tool` present in `primary_tools`.
6. Add cross-check parity probe: assert that tools with `DEPRECATED:` in their docstring are all in `deprecated_aliases` (Option B registry vs Option A fallback check — one-time validation, not runtime).
7. Add capabilities snapshot to `word/evidence/uat-phase1/capabilities_v2_snapshot.json`.

**Risks:**
- The exact count of 37 deprecated aliases must be verified by enumerating all EXTRA_TOOLS and cross-checking which are dispatcher-superseded. If the count differs, update either the registry or the task brief's acceptance checklist accordingly.
- `set_paragraph_format` appears twice in the `EXTRA_TOOLS` list (lines 43 and 99 in `_server_manifest.py` after deduplication review). The implementation must deduplicate.

---

### 6.3 PPTMCP

**Effort: M** (Medium — has tool bundles and compact surface metadata that need mapping to new schema)

**Work items:**
1. Create `_server_manifest.py` (does not currently exist for PPTMCP).
2. Populate `TOOL_REGISTRY` with ~47 entries. All currently registered tools are PRIMARY or DISPATCHER — no deprecated aliases exist yet.
3. Map `COMPACT_TOOL_NAMES` and `TOOL_BUNDLES` to the new schema: `tool_bundles` stays as an existing field; COMPACT_TOOL_NAMES maps to a subset of `primary_tools`.
4. `capabilities()` in `_server_handlers_read.py` calls `build_capabilities_v2(TOOL_REGISTRY)` and merges.
5. Unit tests: assert shape + counts.
6. The existing `tools` list emits parameter-schema dicts. This remains — the new `primary_tools` field emits plain strings only. Both coexist in the output.

**Risks:**
- PPTMCP has 45 always-on tools + 2 COM-conditional tools in `com_tools`. Decision needed: are COM-conditional tools in `primary_tools` or a separate COM registry? Recommendation: include them in `primary_tools` with a companion `com_conditional_tools: list[str]` field (existing pattern from WordMCP's `com_tools` dict). The schema does not prohibit this.
- `_server_compact.py` alters which tools are registered when `PPT_COMPACT_TOOL_SURFACE=true`. The registry should reflect the full surface; the compact surface is a runtime filter. `total_callable_endpoints` should reflect the compact surface count when compact mode is active. This requires `build_capabilities_v2()` to accept an optional `active_tools: set[str]` filter.

---

### 6.4 ExcelMCP

**Effort: M** (Medium — large tool count ~63, no dispatcher pattern yet, no `_server_manifest.py`)

**Work items:**
1. Create `excel/src/excelmcp/_server_manifest.py` (does not exist).
2. Enumerate all ~63 tools from `_io.capabilities()` flat list and classify each as PRIMARY.
3. `deprecated_aliases` will be empty list initially — no deprecated tools exist yet.
4. `operation_scope_enums` will be empty dict initially — no dispatcher tools exist yet.
5. Update `_io.capabilities()` to call `build_capabilities_v2(TOOL_REGISTRY)` and merge.
6. Unit tests.

**Risks:**
- ExcelMCP has the largest tool count and most module fragmentation (12 server_*.py files). Enumerating all tools accurately requires careful cross-checking against each server_*.py file's `@mcp.tool()` registrations. The `_io.capabilities()` flat list is the reference but may be stale if new tools were added without updating it.
- COM-conditional tools (`recalculate_workbook`, `export_as_pdf`, etc.) need the same COM registry pattern as WordMCP and PPTMCP.

---

### 6.5 MailMCP

**Effort: L** (Large — must create `capabilities()` tool from scratch; no manifest infrastructure)

**Work items:**
1. Create `mailmcp/src/mailmcp/_server_manifest.py`.
2. Create `TOOL_REGISTRY` enumerating all ~49 tools.
3. Classify all tools as PRIMARY (no deprecated aliases).
4. `write_gate_metadata` must be populated for all destructive tools (send, compose, forward, create_folder, meeting_draft, task write, category write, etc.).
5. Register `capabilities` as a new MCP tool in `server.py` via `_register()`.
6. Unit tests.
7. Documentation update.

**Risks:**
- MailMCP tool count (49) is approximate — the exact count depends on whether `mailrepo_*` tools are considered part of the MailMCP surface or a separate layer. They are currently registered on the same `mcp` instance, so they must be included.
- `write_gate_metadata` for MailMCP is complex: `OUTLOOK_ENABLE_SEND` gates send/reply/forward; there may be separate gates for calendar mutation and contact creation. Each gate must be enumerated accurately.
- MailMCP `_register_in_executor` vs `_register` pattern: the `capabilities()` tool should use `_register()` (not executor) since it does no I/O.

---

### 6.6 Summary Table

| Package | Effort | Deprecated aliases | New `_server_manifest.py` | New `capabilities()` tool |
|---|---|---|---|---|
| WordMCP | S | 37 | Enrich existing | No (exists) |
| PPTMCP | M | 0 (initial) | Create new | No (exists) |
| ExcelMCP | M | 0 (initial) | Create new | No (exists) |
| MailMCP | L | 0 (initial) | Create new | YES — create from scratch |

---

## 7. Compatibility Risk Review

### 7.1 Additive Non-Breaking Guarantee

The schema design is unconditionally additive. No existing field is removed or renamed. The 5 required fields and 3 extension fields are ADDED to the existing output dict via `result.update(v2_fields)`.

### 7.2 Risk: Consumers indexing by flat `tools` array

**Existing behavior:** `capabilities()["tools"]` returns a `list[str]` (WordMCP, ExcelMCP) or `list[dict]` (PPTMCP).

**After migration:** `capabilities()["tools"]` is UNCHANGED. The new `primary_tools` field is a separate key.

**Risk level: NONE.** The `tools` key still exists with identical content and type.

### 7.3 Risk: Consumers reading WordMCP-specific fields

**Existing behavior:** Consumers reading `capabilities()["read_document_scopes"]`, `capabilities()["table_operations"]`, etc.

**After migration:** All WordMCP-specific fields (`read_document_scopes`, `export_scopes`, `table_operations`, `style_operations`, `content_operations`, `review_operations`, `document_operations`) are UNCHANGED. They remain in the output alongside the new `operation_scope_enums` dict (which contains the same information in a cross-package-standard format).

**Risk level: NONE.** These fields are kept; `operation_scope_enums` is additive.

### 7.4 Risk: Consumers reading PPTMCP-specific fields

**Existing behavior:** Consumers reading `capabilities()["tool_bundles"]`, `capabilities()["compact_tool_surface"]`, `capabilities()["tools"]` (as list[dict] with params).

**After migration:** All PPTMCP-specific fields are UNCHANGED. `primary_tools` uses plain strings, but the existing `tools` list[dict] remains.

**Risk level: NONE.**

### 7.5 Risk: Consumers reading ExcelMCP-specific fields

**Existing behavior:** `capabilities()["prompts"]`, `capabilities()["metadata_contract_version"]`, `capabilities()["metadata_policy"]` — all unique to ExcelMCP.

**After migration:** All ExcelMCP-specific fields are UNCHANGED.

**Risk level: NONE.**

### 7.6 Risk: New fields with name collision

**Check:** Do any existing package outputs already use `primary_tools`, `deprecated_aliases`, `replacement_tool`, `replacement_operation_or_scope`, `total_callable_endpoints`, `operation_scope_enums`, `write_gate_metadata`, `deprecation_policy`?

**Finding:** None of the 4 packages emit any of these 8 field names today (confirmed by source audit in section 1). There is zero collision risk.

**Risk level: NONE.**

### 7.7 Risk: `mcpshared` version dependency

Adding `capabilities_schema.py` and `capabilities_builder.py` to `mcpshared` requires that all consuming packages declare a dependency on `mcpshared`. WordMCP, ExcelMCP, PPTMCP, and MailMCP currently depend on `mcpshared` for ACP contracts and inline artifact utilities — the dependency already exists.

**Risk level: LOW.** Version pinning in each package's `pyproject.toml` should be checked during the adoption PRs to ensure the minimum `mcpshared` version includes the new modules.

### 7.8 Risk: `total_callable_endpoints` drift

If a new tool is added to a package without updating the `TOOL_REGISTRY`, the manifest will undercount `total_callable_endpoints`. This is detectable by cross-checking against the live MCP tool list in the unit test parity probe (acceptance checklist item 7).

**Risk level: LOW.** Mitigated by the parity probe unit test required in section 8.

---

## 8. Acceptance Checklist for Chained Reference Impl PR ([internal-ref]

The chained PR implements WordMCP as the reference implementation. It must satisfy all of the following before merge:

### 8.1 Schema field correctness

- [ ] `capabilities()` output contains all 5 required fields: `primary_tools`, `deprecated_aliases`, `replacement_tool`, `replacement_operation_or_scope`, `total_callable_endpoints`
- [ ] `capabilities()` output contains all 3 extension fields: `operation_scope_enums`, `write_gate_metadata`, `deprecation_policy`
- [ ] `primary_tools` has exactly 10 entries: `capabilities`, `content`, `create_document`, `document`, `export`, `paragraph`, `read_document`, `review`, `style`, `table` (alphabetically sorted)
- [ ] `deprecated_aliases` has exactly 37 entries (alphabetically sorted)
- [ ] `replacement_tool` has exactly 37 keys — one per deprecated alias
- [ ] Every value in `replacement_tool` is a member of `primary_tools`
- [ ] `replacement_operation_or_scope` has exactly 37 keys — one per deprecated alias
- [ ] `total_callable_endpoints` == 47 == `len(primary_tools)` + `len(deprecated_aliases)`
- [ ] All 3 extension fields are populated (non-null dicts)
- [ ] `operation_scope_enums` has exactly 8 keys (one per dispatcher tool)
- [ ] `deprecation_policy` contains `window_releases` (int), `telemetry_field`, `removal_date_iso`

### 8.2 Backward compatibility

- [ ] `capabilities()["tools"]` still returns a flat `list[str]` (47 entries, unchanged)
- [ ] `capabilities()["read_document_scopes"]` still returns `list[str]` (unchanged)
- [ ] `capabilities()["table_operations"]` still returns `list[str]` (unchanged)
- [ ] `capabilities()["governance"]` still returns the governance dict (unchanged)
- [ ] `capabilities()["com_tools"]` still returns the COM dict (unchanged)

### 8.3 Unit tests

- [ ] Test: `len(result["primary_tools"]) == 10`
- [ ] Test: `len(result["deprecated_aliases"]) == 37`
- [ ] Test: `result["total_callable_endpoints"] == 47`
- [ ] Test: `set(result["replacement_tool"].keys()) == set(result["deprecated_aliases"])`
- [ ] Test: `set(result["replacement_tool"].values()).issubset(set(result["primary_tools"]))`
- [ ] Test: `set(result["replacement_operation_or_scope"].keys()) == set(result["deprecated_aliases"])`
- [ ] Test: `result["primary_tools"]` is sorted (alphabetical order)
- [ ] Test: `result["deprecated_aliases"]` is sorted (alphabetical order)
- [ ] Test: all keys in `result["operation_scope_enums"]` are in `result["primary_tools"]`
- [ ] Test: each value in `result["operation_scope_enums"]` is a non-empty sorted list
- [ ] Test: `"WORD_ENABLE_WRITE"` is the `env_var` for every write-gated tool in `write_gate_metadata`

### 8.4 Parity probes (cross-check)

- [ ] Parity probe A: `set(result["primary_tools"]) | set(result["deprecated_aliases"]) == set(result["tools"])` — structured fields collectively cover the same tools as the flat `tools` array (modulo any deduplication in the flat list)
- [ ] Parity probe B: For every tool in the codebase whose docstring begins with `DEPRECATED:`, the tool name appears in `result["deprecated_aliases"]` — Option B registry cross-checked against Option A ground truth (one-time validation test, not runtime)
- [ ] Parity probe C: `len(result["tools"]) == result["total_callable_endpoints"]` — flat list count matches structured count

### 8.5 Evidence bundle

- [ ] `word/evidence/uat-phase1/capabilities_v2_snapshot.json` — captured output of `capabilities()` after migration, pretty-printed JSON
- [ ] Snapshot committed alongside the PR so reviewers can diff against the pre-migration output

### 8.6 Implementation files

- [ ] `shared/src/mcpshared/capabilities_schema.py` created with `ToolKind`, `ToolRegistryEntry`, `WriteGateSpec`, `CapabilitiesV2Fields`
- [ ] `shared/src/mcpshared/capabilities_builder.py` created with `build_capabilities_v2()`
- [ ] `shared/src/mcpshared/__init__.py` updated to export new public symbols
- [ ] `word/src/wordmcp/_server_manifest.py` updated with `TOOL_REGISTRY` (47 entries) and `DEPRECATION_POLICY`
- [ ] `build_capabilities()` in `_server_manifest.py` calls `build_capabilities_v2()` and merges

### 8.7 Tests run clean

- [ ] `pytest word/tests/ -m "not integration" -v --tb=short` — all pass
- [ ] `pytest shared/tests/ -v --tb=short` — all pass (new tests for builder)
- [ ] `ruff check shared/src/ word/src/` — zero errors

---

## Open Questions for Chained Reference Impl PR

1. **Exact deprecated_aliases count:** The task brief states 37. The docstring scan identified 13 tools with explicit `DEPRECATED:` markers. The full 37 requires classifying ALL EXTRA_TOOLS entries as deprecated. The chained PR must confirm the exact list and reconcile against the task brief. If the count differs from 37, the acceptance checklist must be updated.

2. **`set_paragraph_format` deduplication:** This tool name appears to be listed twice in `EXTRA_TOOLS` in `_server_manifest.py` (compare lines 43 and 99 of `EXTRA_TOOLS`). The registry must have exactly one entry per tool. Deduplicate before finalizing the 47-count claim.

3. **`export_review_evidence` replacement operation:** The docstring says "Use export(scope='review') when a dedicated review-export scope is added" — the scope does not yet exist. The registry entry should use a sentinel value such as `"review_export_pending"` and the chained PR should decide whether to implement the scope or leave the sentinel.

4. **Write-gated dispatch tools:** For dispatcher tools like `style` (which has both read `list` and write `apply` operations), `write_gate_metadata` currently recommends including the dispatcher if ANY operation requires a gate. A future schema version could be operation-granular. The chained PR should document this limitation in a comment in `capabilities_schema.py`.

5. **`mcpshared` version bump:** The new modules require a `mcpshared` minor version bump. Confirm the versioning convention used in this repo (currently `mcpshared` has no version-gated imports in consuming packages).

6. **PPTMCP compact surface and `total_callable_endpoints`:** When `PPT_COMPACT_TOOL_SURFACE=true`, fewer tools are registered. The `build_capabilities_v2()` signature proposed in section 5.3 includes an optional `active_tools` filter parameter. The chained PR (WordMCP-only) does not need to implement this, but the `capabilities_builder.py` should be designed to accommodate it in the function signature.
