# mcpshared — Shared contracts for MCP Office

`mcpshared` is the shared library for the [MCP Office](../README.md) repository.
It defines the **Artifact Context Packet (ACP) v1** TypedDict hierarchy, validator, and
sanitisation helpers consumed by every MCP server in the suite (`excelmcp`, `pptmcp`, etc.).
It is not an MCP server itself and exposes no tools or transport; it is a pure Python
library imported as a dependency by the other packages.

---

## Contents

| Module | Provides |
|--------|---------|
| `_acp_contract.py` | ACP TypedDicts (6), annotation validator, formula-injection sanitiser |

---

## ACP Contract

The Artifact Context Packet is a structured payload that MCP servers attach to tool
responses to give the AI model structured context about the artefact it just operated on.
Packets nest three levels of detail so callers can request only as much context as needed.

```
ACPIndex  (Level 1 — always present)
  └─ ACPFocused  (Level 2 — adds optional ACPContent)
       └─ ACPDeep  (Level 3 — adds optional ACPDetail + ACPAnnotation list)
```

### TypedDicts

#### `ACPAnnotation` (`total=False`)

Free-form key/value tag attached to an `ACPDeep` payload.

| Field | Type | Constraint |
|-------|------|-----------|
| `key` | `str` | max 64 chars; no leading `= + - @ \t \r \n`; no NUL |
| `value` | `str` | max 1 024 chars; no formula prefix; no NUL |
| `source` | `str` | max 64 chars; e.g. `"user"`, `"rule-engine"`; no NUL |
| `rule_id` | `str` | max 64 chars; e.g. `"IBCS-1.3"`; optional |

#### `ACPIndex` (Level 1, required fields)

Minimum index always included in every ACP response.

| Field | Type | Notes |
|-------|------|-------|
| `acp_version` | `str` | Always `"1.0"` |
| `artifact_type` | `Literal["presentation", "spreadsheet"]` | |
| `artifact_id` | `str` | e.g. `"deck:report.pptx"` or `"wb:budget.xlsx"` |
| `tool_name` | `str` | `"pptmcp"` \| `"excelmcp"` |
| `summary` | `str` | One-line human description, max 120 chars |
| `timestamp` | `str` | ISO 8601 UTC, e.g. `"2025-01-15T14:00:00Z"` |

#### `ACPContent` (`total=False`)

Optional level-2 content summary embedded in `ACPFocused.content`.

| Field | Type | Populated by |
|-------|------|-------------|
| `slide_count` | `int` | pptmcp |
| `slide_titles` | `list[str]` | pptmcp |
| `sheet_count` | `int` | excelmcp |
| `sheet_names` | `list[str]` | excelmcp |
| `table_names` | `list[str]` | excelmcp |

#### `ACPDetail` (`total=False`)

Optional level-3 detail embedded in `ACPDeep.detail`.
**Security note S-001:** raw cell values are intentionally excluded to prevent leaking
sensitive workbook content through the context layer. Adapters MUST NOT embed cell
values in any `ACPDetail` field.

| Field | Type | Populated by |
|-------|------|-------------|
| `png_paths` | `list[str]` | pptmcp (export artefact paths) |
| `review_findings` | `list[dict]` | pptmcp (evidence bundle findings) |
| `metadata_policy` | `str` | excelmcp — `"strict"` \| `"lenient"` |
| `named_ranges` | `list[str]` | excelmcp — names only, not values |
| `lineage_summary` | `dict` | excelmcp — output of `normalize_metadata()` |

#### `ACPFocused(ACPIndex, total=False)` (Level 2)

Extends `ACPIndex` with:

| Field | Type |
|-------|------|
| `content` | `ACPContent` |

#### `ACPDeep(ACPFocused, total=False)` (Level 3)

Extends `ACPFocused` with:

| Field | Type |
|-------|------|
| `detail` | `ACPDetail` |
| `annotations` | `list[ACPAnnotation]` |

---

### `validate_acp_annotations(annotations)`

```python
def validate_acp_annotations(annotations: list[ACPAnnotation]) -> list[str]: ...
```

Validates every annotation in the list. Returns an empty list if all are valid; otherwise
returns a list of human-readable error strings (one per violation). Checks enforced per annotation:

- `key`: max 64 chars, no formula-prefix leading char, no NUL byte
- `value`: max 1 024 chars, no formula-prefix leading char, no NUL byte
- `source`: max 64 chars, no NUL byte
- `rule_id`: max 64 chars, no formula-prefix leading char, no NUL byte

**Raises:** nothing — all errors are returned as strings so the caller decides whether to
reject the payload or surface a warning.

---

### `_sanitize_text_field(value, max_len=255)`

```python
def _sanitize_text_field(value: str | None, max_len: int = 255) -> str: ...
```

Formula-injection sanitiser (OWASP A03). Applied to all user-controlled strings before
they are embedded in an ACP payload.

| Input | Behaviour |
|-------|-----------|
| `None` or non-`str` | Returns `""` |
| Starts with `= + - @ \t \r \n` | Prefix-escapes with `'`; output length ≤ `max_len` |
| Clean string | Truncated to `max_len` chars |

**Contract:** output is always a `str` and always `len(output) <= max_len`.

---

## Security Model

All user-controlled strings that enter an ACP payload MUST be passed through
`_sanitize_text_field` before assignment.

| Rule | Detail |
|------|--------|
| Formula-prefix strip | Characters `= + - @ \t \r \n` at position 0 are escaped with a `'` prefix |
| Length cap | Output is always ≤ `max_len` chars (default 255) |
| NUL bytes | `validate_acp_annotations` rejects annotations containing `\x00` in any field |
| S-001 (cell values) | `ACPDetail` fields MUST NOT contain raw cell values; adapters are responsible for this exclusion |

`excelmcp`'s `_acp_adapter.py` sanitises `sheet_names`, `table_names`, and
`lineage_summary["workbook_name"]` via `_sanitize_text_field` before building the packet.

---

## Installation

```powershell
# From the shared/ directory
pip install -e .

# Or from the repo root
pip install -e shared/

# With dev dependencies
pip install -e "shared/[dev]"
```

Dependencies: none (stdlib only). Python `>=3.11` required.

---

## Testing

```powershell
pytest shared/tests/ -q
```

Current: **33 tests** in `shared/tests/test_acp_contract.py` covering all TypedDict
constraints, validator edge cases, and sanitiser contracts.
