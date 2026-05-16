# Changelog

All notable changes to ExcelMCP are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

### Highlights

- **`capabilities_v2` schema** — `capabilities()` now emits a machine-readable
  `capabilities_v2` nested object. Dispatcher agents can discover tool families,
  operation enums, write gates, and tool modes without parsing docstrings.
  Schema defined in `shared/design/capabilities_schema_v1.md`.
- **Public-code hygiene** — internal domain-specific tooling removed from all public
  surfaces (source, tests, scripts, docs).

---

### ExcelMCP

#### Added
- `_server_manifest.py` — TOOL_REGISTRY (64 entries), 17 dispatcher vocabulary
  constants (`SHEET_OPERATIONS`, `CELL_OPERATIONS`, `RANGE_IO_OPERATIONS`, …),
  `GATE_CLASSES` (write / com / macros / allowlist), `DEPRECATION_POLICY`, and
  `build_capabilities_v2()`.
- `capabilities()` now emits a nested `capabilities_v2` key with:
  - `primary_tools` — sorted list of 64 primary tool names
  - `deprecated_aliases` — `[]` (no deprecated aliases at this phase)
  - `total_callable_endpoints` — 64
  - `operation_scope_enums` — 17 dispatchers with sorted vocabulary lists
  - `gate_metadata` — per-tool gate list with `kind`, `env_var`, `requires_confirm`
  - `write_gate_metadata` — backward-compatible gate subset
  - `gate_classes` — 4 entries (write / com / macros / allowlist) with scope annotations
  - `tool_modes` — maps all 64 tools to `openpyxl` / `com_conditional` / `com_session`
  - `deprecation_policy` — `window_releases=2`, `telemetry_field`, `removal_date_iso`
- Phase 3E: `format` dispatcher consolidates 8 standalone format tools
  (`apply_number_format`, `apply_style`, `apply_alignment`, `add_border`,
  `apply_format_to_sheet_list`, `copy_range_format`, `write_range_with_format`,
  `apply_conditional_format`) under `format(target=...)`.
- Backend decomposition: `_io.py` and `server_io.py` split into focused modules
  (`server_range`, `server_sheet`, `server_workbook`, `server_table_io`,
  `_range_read`, `_range_write`).
- `read_summary_routine.py` — COM-free openpyxl read routine with allowlist validation.

#### Removed
- Internal domain-specific tool removed from all public surfaces
  (source, tests, scripts, docs). `TOOL_REGISTRY` count: 65 → 64.

#### Fixed
- `session_mode: bool = False` stub parameter removed from `cell`, `range_io`,
  `append_rows`, and `create_workbook` — these stubs raised `ToolError` on use and
  delivered no value.
- Write gate pattern standardised: `_check_write()` + `_check_confirm()` used
  consistently across all write-gated tools.

---

### Shared

#### Added
- `shared/design/capabilities_schema_v1.md` — cross-package `capabilities_v2` schema
  specification. Documents the 8 required fields and 3 ExcelMCP extension fields.

---

### Test Coverage

| Suite | Result |
|---|---|
| excelmcp unit + smoke | 1354 passed, 0 failed |
