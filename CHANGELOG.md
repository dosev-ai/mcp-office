# Changelog

All notable changes to MCP Office are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [v0.4] — `dddbc12` · Jun 14 2026 — pptmcp + wordmcp suite release

Added pptmcp (48 tools) and wordmcp (51 tools) to the public mcp-office suite.

**pptmcp:** Build, edit, review, and export PowerPoint presentations. Output Contract
framework for machine-verifiable slide specs. COM-conditional tools for PDF export,
chart recalculation, and live slide show. python-pptx backend for all structural tools.

**wordmcp:** Template assembly, tracked-changes support, and structural QA for Word
documents. COM-conditional tools for PDF/HTML export and tracked-changes management.
python-docx backend for all read/write/review tools (no Word required for most tools).

**Release validation — E2E UAT on fresh public install (Windows 11, Claude Desktop):**

| Package | Stories | Result | Install |
|---|---|---|---|
| wordmcp | 10/10 | ✅ All pass | `pip install -e ./wordmcp` |
| excelmcp | 9/10 + 1 partial | ✅ All pass (1 config note) | `pip install -e ./excelmcp` |
| pptmcp | 10/10 | ✅ All pass | `pip install -e ./shared && pip install -e ./pptmcp` |

Full UAT evidence recorded internally.

**excelmcp config note:** `export_range_as_csv` with an `output_path` requires `EXCEL_CSV_EXPORT_ROOTS` to be set in the server env (same pattern as `EXCEL_ALLOWLIST_ROOTS`). CSV data is returned inline without it — add `"EXCEL_CSV_EXPORT_ROOTS": "C:\\path\\to\\your\\files"` to your config to enable file-write output.

**Docs:** README, docs/quickstart.md, mcp.json.template all updated to cover all 3 packages.
Env var name fix: `EXCEL_ALLOWED_DIRS` → `EXCEL_ALLOWLIST_ROOTS` (was wrong in docs).

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
