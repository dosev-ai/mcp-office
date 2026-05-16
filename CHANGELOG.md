# Changelog

All notable changes to ExcelMCP are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [v0.3] — `5cd6dad` · [PR #5](https://github.com/dosev-ai/mcp-office/pull/5) · May 16 2026

`capabilities_v2` machine-readable registry, `format()` dispatcher, server decomposition,
`session_mode` stub removal, public-code hygiene (TOOL_REGISTRY 65→64).

Full schema spec: [`shared/design/capabilities_schema_v1.md`](shared/design/capabilities_schema_v1.md)

---

## [v0.2] — `cbfe0b1` · Apr 24 2026

Namespace package shadowing fix — `python -m excelmcp.server` now works correctly
when invoked from the repo root.

---

## [v0.1] — `cea346c` · Apr 23 2026

Initial public export of ExcelMCP Phase 1: 60+ tools, openpyxl backend, COM-conditional
extensions, allowlist governance, CI scaffold, quickstart docs.

---
