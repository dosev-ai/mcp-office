# excelmcp — Excel MCP Server (Sprint P) v0.6.0

`excelmcp` is a Python MCP stdio server for reading and editing `.xlsx` files using
[openpyxl](https://openpyxl.readthedocs.io/) as the backend. It is part of the
[MCP Office](../README.md) repository and operates entirely at the file level — no
Excel application is required for the core openpyxl tool surface; optional COM-backed Sprint O tools add live recalculation, PDF export, template hydration, pivot-table operations, and live formula evaluation via `evaluate_formula(live=True)` on Windows when `EXCEL_ENABLE_COM=true`. Write operations are gated
behind `EXCEL_ENABLE_WRITE=true` (environment variable) and a `confirm=True` parameter
for destructive calls, preventing accidental mutations.

---

## Quick Start

```powershell
# Install the package from the public repo root
pip install -e "./excelmcp"

# Or install with dev dependencies
pip install -e "./excelmcp[dev]"

# Reproducible install using pinned versions (constraints only — avoids pywin32 on Linux)
pip install -c excelmcp/requirements-lock.txt -e "./excelmcp[dev]"

# Run tests
pytest excelmcp/tests/ -m "not integration" -v --tb=short
```

The lock file `excelmcp/requirements-lock.txt` pins exact runtime dependency versions for reproducible installs. The bundled sibling package `shared/` is installed automatically via the local file dependency declared in `excelmcp/pyproject.toml`.

### VS Code mcp.json Registration

Add the following entry to
`%APPDATA%\Code - Insiders\User\mcp.json` (or `Code\User\mcp.json` for stable VS Code):

```json
"excel-excelmcp": {
    "type": "stdio",
    "command": "python",
    "args": ["-m", "excelmcp.server"],
    "env": {
            "PYTHONPATH": "C:\\path\\to\\mcp-office\\excelmcp\\src",
            "EXCEL_ALLOWLIST_ROOTS": "C:\\Workbooks,D:\\ExcelData",
        "EXCEL_ENABLE_WRITE": "true"
    }
}
```

---

## Environment Variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `EXCEL_ALLOWLIST_ROOTS` | **yes** | — | Comma-separated absolute folder paths the server may access. Any path outside these roots is blocked with a `"Path not in allowlist"` error. |
| `EXCEL_ENABLE_WRITE` | no | `false` | Must be `true` to enable any mutating tool. If unset or `false`, all write calls return an error even when `confirm=True` is passed. |
| `EXCEL_MAX_RANGE_CELLS` | no | `5000` | Maximum number of cells returned by a single `read_range` or `write_range` call. Prevents accidental oversized reads. |
| `EXCEL_ALLOW_FILE_HYPERLINKS` | no | `false` | Allow `file://` scheme in `bulk_add_hyperlinks`. Disabled by default — Windows UNC paths (`file://attacker.com/share`) can trigger NTLM credential relay. Enable only in fully trusted environments. |
| `EXCEL_MAX_METADATA_BYTES` | no | `32768` | Maximum UTF-8 payload size accepted by `workbook_metadata` and the metadata resource. Payloads must also fit Excel's single-cell text limit (`32767` chars). |
| `EXCEL_DA_COM_PROBE` | no | `off` | Stage-4 COM validation pass after DA formula patching. `off` = skipped; `auto` = runs if win32com is available; `required` = runs and fails if win32com unavailable. Windows-only. |
| `EXCEL_DA_FINALIZE_MODE` | no | `warn` | DA finalization pipeline strictness. `warn` = log failures and continue; `strict` = raise `DAFinalizeError` on any failed stage. |
| `EXCEL_ENABLE_COM` | no | `false` | Must be `true` to activate COM-backed operations: `recalculate_workbook`, `export_as_pdf`, `hydrate_template`, `pivot_table`, and `evaluate_formula(live=True)`. Windows-only; requires `pywin32`. |
| `EXCEL_EXPORT_ROOTS` | no | — | Comma-separated absolute folder paths permitted as PDF output destinations. Required when `export_as_pdf` is used. |
| `EXCEL_METADATA_POLICY` | no | `strict` | Metadata validation policy. `strict` (default): any missing or unreadable field in `build_workbook_metadata()` is a hard error. `lenient`: optional fields may be empty; only `path` and `workbook_name` are required. Injected into the `/metadata` and `/meta` resource responses as the `policy` field. |
| `EXCEL_CSV_EXPORT_ROOTS` | no | — | Comma-separated absolute folder paths permitted as CSV output destinations when `export_range_as_csv` writes to disk. Required when `output_path` is supplied. Separate from `EXCEL_ALLOWLIST_ROOTS` (source read allowlist) and `EXCEL_EXPORT_ROOTS` (PDF output allowlist). |
| `EXCEL_SESSION_MODE` | no | `false` | Set to `true` to enable session-mode routing for the active-workbook bridge. When `true` but `EXCEL_ENABLE_COM` is not `true`, a `UserWarning` is emitted at runtime and the bridge tools return errors. Has no effect without `EXCEL_ENABLE_COM=true`; safe to set on non-Windows where COM is unavailable. |
| `EXCEL_ENABLE_MACROS` | no | `false` | Must be `true` to permit the `run_macro` tool. When unset or `false`, `run_macro` returns an error. `list_macros` does not require this variable. Used in `_ensure_macros_gate()`. Requires `EXCEL_ENABLE_COM=true`. |

---

## Dynamic Array Formula Support

openpyxl does not emit the OOXML metadata required by Excel 365 for dynamic array
(spill) formulas. Without it, Excel's repair dialog removes every spill formula on
first open. excelmcp applies a **post-save OOXML patch** automatically when saving
workbooks that contain any of the following functions:

`FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `RANDARRAY`, `EXPAND`, `TAKE`,
`DROP`, `TOROW`, `TOCOL`, `WRAPROWS`, `WRAPCOLS`, `LET`, `XLOOKUP`, `XMATCH`,
`MAKEARRAY`

The patch runs as part of the 4-stage finalization pipeline (`_da_finalize`):

| Stage | Description |
|-------|-------------|
| 1 — patch | OOXML metadata injection: `cm="1"` linkage on DA cells, `xl/metadata.xml`, Content-Type override, legacy CSE attribute strip |
| 2 — classify | Outcome classification (`patched` / `no_da_detected`) |
| 3 — validate | Invariant validation (skipped until a future sprint) |
| 4 — probe | Optional COM round-trip via `_da_com.probe_excel_reopen` — controlled by `EXCEL_DA_COM_PROBE`. Requires `EXCEL_ENABLE_COM=true`; raises `NotAllowedError` otherwise. Non-temp paths validated against `EXCEL_ALLOWLIST_ROOTS` (denied by default when unset). `AutomationSecurity = 3` enforced before any workbook open — macro execution blocked. |

### COM Formula2 Preservation (LET+FILTER Gap — Fixed)

**Previously a Known Open Gap.** `LET()` wrapping a spill function (e.g.
`=LET(src, FILTER(...), INDEX(src, SEQUENCE(ROWS(src)), {cols}))`) written via
Excel COM `.Formula2` now survives Excel MCP (openpyxl) save round-trips.

**Root cause:** openpyxl's load/save cycle silently drops the `t="array" aca="1"
ref="..." ca="1"` attributes from `<f>` elements. These attributes are required by
Excel 365 to recognise LET+FILTER cells as dynamic-array anchors. The `cm="1"`
approach alone is insufficient for deeply nested LET expressions.

**Fix implemented in `_core.py` + `_ooxml.py`:**

1. **Pre-save snapshot** — Before `wb.save()`, `snapshot_formula2_cells(path)` reads
   the OOXML zip and captures all cells with `aca="1"` on their `<f>` element (the
   Excel 365 Formula2 marker written by COM).
2. **Post-patcher restore** — After the DA patcher runs, `restore_formula2_f_attrs(path,
   snapshots)` re-injects the original `t="array" aca="1" ref ca` attributes into
   those cells. The `cm="1"` added by the patcher is preserved alongside them.
3. **Strip-pass guard** — `_strip_legacy_f_attrs` in `_ooxml.py` has an `aca="1"`
   early-return guard: cells carrying `aca="1"` are skipped entirely by the legacy-CSE
   strip pass, preventing accidental attribute removal before the restore runs.

Both `_save_and_evict` and `_flush_if_dirty` in `_core.py` follow this
snapshot → save → patch → restore sequence.

**Compatibility table:**

| Formula type | Engine | Result after MCP save |
|---|---|---|
| `=FILTER(...)` plain | openpyxl patcher | ✅ `cm="1"` + `_xlfn.FILTER` |
| `=LET(x, FILTER(...), x)` via COM `.Formula2` | snapshot/restore | ✅ `t="array" aca="1" ref` preserved |
| `=LET(src, ..., INDEX(FILTER(...), SEQUENCE(...), {...}))` via COM | snapshot/restore | ✅ full attributes preserved |
| Legacy CSE `{=SUM(...)}` | patcher strips `t="array"` | ✅ treated as scalar DA |

Plain spill functions (non-LET) continue to use the existing `cm="1"` path.

---

## Tools Reference

64 tools: **server_io.py** (18), **server_ops.py** (16), **server_format.py** (8), **server_batch.py** (6), **server_com.py** (7), **server_chart.py** (1), **server_review.py** (3), **server_acp.py** (1), **server_snapshot.py** (2), **server_com_vba.py** (2).

Unified dispatch tools (e.g. `sheet`, `cell`, `range_io`) consolidate related operations behind one endpoint — pass `operation=` to select the mode. Read modes need no write gate; write modes require `EXCEL_ENABLE_WRITE=true` + `confirm=True`.

`capabilities` also returns a live prompt inventory from the shared MCP registry via `prompts.count` and `prompts.names`.

### Read Tools (10)

No write-gate required. Safe to call without `EXCEL_ENABLE_WRITE`.

| Tool | Description | Key response fields |
|---|---|---|
| `capabilities` | Returns server version, phase, backend, tool list, live prompt inventory, and metadata governance fields | `version`, `phase`, `backend`, `tools[]`, `prompts.count`, `prompts.names[]`, `governance`, `metadata_contract_version`, `metadata_policy` |
| `get_used_range` | Bounding box of the used cells on a sheet; `empty: true` for blank sheets | `sheet`, `min_row`, `max_row`, `min_col`, `max_col`, `empty` |
| `read_sheet_all` | Full used range of a worksheet as a 2-D value array | `sheet`, `values[][]` |
| `get_cell_metadata` | Font, fill, number format, comment, and hyperlink for a single cell | `font`, `fill`, `number_format`, `comment`, `hyperlink` |
| `evaluate_formula` | Formula string and value for a cell: cached openpyxl value by default, or live Excel COM evaluation when `live=True` (`EXCEL_ENABLE_COM=true`, Windows-only) | cached path: `sheet`, `address`, `formula`, `cached_value`; live path also returns `ok`, `note`, `value` |
| `list_tables` | All Excel tables in the workbook; optionally filter by sheet | `tables[].name`, `tables[].sheet`, `tables[].ref` |
| `read_table` | Headers and all data rows from a named Excel table | `table`, `sheet`, `headers[]`, `rows[][]` |
| `validate_formula_syntax` | Check formula syntax without writing; `path` is optional | `valid`, `error` |
| `diff_workbooks` | Compare two workbooks cell-by-cell and return differences; read-only (`data_only=True`) | `path_a`, `path_b`, `sheets_compared[]`, `diffs[]`, `total_diffs`, `truncated` |
| `consolidate_ranges` | Read multiple ranges from workbooks and merge into one table; max 20 range tuples, 500 rows | `ok`, `ranges_read`, `total_rows`, `truncated`, `rows[]` |

### Unified Dispatch Tools (10)

Pass `operation=` to select the mode. Read operations need no write gate; write operations require `EXCEL_ENABLE_WRITE=true` + `confirm=True`.

| Tool | Operations | Replaces | Description |
|---|---|---|---|
| `sheet` | `list` · `add` · `rename` · `delete` | `list_sheets`, `add_sheet`, `rename_sheet`, `delete_sheet` | Manage workbook sheets; add/rename/delete require `confirm=True` |
| `cell` | `read` · `write` | `read_cell`, `write_cell` | Read or write a single cell value; write mode persists immediately on success. `session_mode=True` on the write path routes via COM to the active workbook (requires `EXCEL_ENABLE_COM=true` + `EXCEL_ENABLE_WRITE=true` + `confirm=True`; formula values rejected) |
| `range_io` | `read` · `write` | `read_range`, `write_range` | Read or write a rectangular range; capped at `EXCEL_MAX_RANGE_CELLS`; write mode persists immediately on success. `session_mode=True` on the write path routes via COM to the active workbook (same guards as `cell`; ragged rows padded to `n_cols` with `None` for COM SAFEARRAY safety) |
| `named_range` | `list` · `read` · `write` | `get_named_ranges`, `read_named_range`, `write_named_range` | List, read, or write named range definitions |
| `merge` | `list` · `merge` · `unmerge` | `list_merged_cells`, `merge_cells`, `unmerge_cells` | List, merge, or unmerge cell ranges |
| `cell_comment` | `get` · `delete` | `get_cell_comment`, `delete_cell_comment` | Get (read-only) or delete a cell comment |
| `sheet_properties` | `get` · `set` | `get_sheet_properties`, `set_sheet_properties` | Get or set tab color, visibility state, gridlines, zoom level |
| `data_validation` | `get` · `add` | `get_data_validation_info`, `bulk_add_data_validation` | Get or add data validation rules |
| `hyperlink` | `get` · `add` · `remove` | `get_hyperlinks`, `add_hyperlink`, `bulk_add_hyperlinks`, `remove_hyperlink` | Manage hyperlinks; URL scheme allowlist: `https`, `http`, `mailto`; `file://` requires `EXCEL_ALLOW_FILE_HYPERLINKS=true`; `javascript:`, `data:`, `vbscript:` always blocked |
| `workbook_metadata` | `read` · `write` | — | Read or write workbook-level metadata stored on hidden sheet `_MCP_META`; read returns presence, version, payload, and byte size; write requires `EXCEL_ENABLE_WRITE=true` + `confirm=True` |

### Write / Structural Tools (25)

All require `EXCEL_ENABLE_WRITE=true` + `confirm=True`.

> **Save semantics:**
> - `cell` (write), `range_io` (write), and `append_rows` persist immediately on success via the shared save-and-evict path.
> - `save` remains available as an explicit flush and returns a safe no-op when no write-mode workbook is cached.

| Tool | Save mode | Description |
|---|---|---|
| `create_workbook` | cache | Create a new `.xlsx`; optional `sheets` list; `overwrite=False` guards accidental replacement |
| `save` | — | Flush in-memory cache to disk; no-op if nothing pending |
| `append_rows` | auto | Append rows after the last used row; table-aware — expands `tbl.ref` + syncs autofilter |
| `find_replace` | auto | Find-and-replace strings across all cells; `whole_cell` + `match_case` options |
| `import_csv_to_sheet` | auto | Import CSV/TSV into a sheet; `skip_header=False` default; `overwrite=True` replaces data |
| `rows` | auto | Insert (`operation='insert'`) or delete (`operation='delete'`) rows |
| `cols` | auto | Insert (`operation='insert'`) or delete (`operation='delete'`) columns |
| `protect` | auto | Protect/unprotect a sheet (`scope='sheet'`) or workbook structure (`scope='workbook'`) |
| `set_print_area` | auto | Set or clear the print area; `address=None` to clear |
| `set_column_visibility` | auto | Show or hide columns by letter list; `hidden=True` to hide |
| `set_workbook_calc_mode` | auto | Set `auto`/`manual`/`autoNoTable` calc mode + `fullCalcOnLoad` flag |
| `copy_range` | auto | Copy values + formats from src_address to dst_address within the same sheet |
| `create_table` | auto | Create a native Excel table (ListObject); default style `TableStyleMedium9` |
| `freeze_panes` | auto | Freeze rows/columns at a cell; `address=None` to unfreeze |
| `auto_filter` | auto | Apply or clear Excel AutoFilter on a range; `address=None` to clear |
| `bulk_copy_sheet` | auto | Copy one source sheet to multiple named new sheets in a single call; note: charts/images not copied |
| `set_row_heights` | auto | Set row heights: uniform (`row_range` + `height`) or per-row (`row_specs` list) |
| `set_column_widths` | auto | Explicit `width` or `autofit=True` (content approximation); `col_range` or `columns` list |
| `apply_format_to_sheet_list` | auto | Apply style/number_format/alignment/`column_widths`/`row_heights` to same address across multiple sheets |
| `bulk_add_comments` | auto | Add multiple cell comments in one call; fail-fast validation |
| `fill_formula_range` | auto | Fill a formula across a range with automatic reference relativisation via `Translator` |
| `bulk_range_write` | auto | Write multiple non-contiguous cells in one call using a `writes` list of `{address, value}` items; rejects invalid/out-of-bounds addresses and formula-injection strings |
| `copy_range_format` | auto | Format-painter: copy formatting (not values) from source to destination range; cross-sheet capable |
| `write_range_with_format` | auto | Write a 2-D value array with style/number_format/alignment atomically |
| `apply_conditional_format` | auto | Apply `CellIsRule` or `FormulaRule` conditional formatting; formula injection guard active |
| `snapshot_workbook` | auto | Copy a workbook to a timestamped snapshot file (`shutil.copy2`); dest validated against `EXCEL_ALLOWLIST_ROOTS`; path traversal guard on `filename` |

### Formatting Tools (4)

All require `EXCEL_ENABLE_WRITE=true` + `confirm=True`. `apply_style` auto-saves on success; `apply_number_format`, `apply_alignment`, and `add_border` do **NOT** auto-save — call `save()` to persist.

| Tool | Auto-save? | Description |
|---|---|---|
| `apply_number_format` | no | Apply Excel format codes (e.g. `#,##0.00`, `YYYY-MM-DD`, `+0.0%;-0.0%`) to a cell or range |
| `apply_style` | **yes** | Set bold, italic, font size, font color, and/or fill; pass `address` (preferred) or `ranges` list |
| `apply_alignment` | no | Set horizontal (`left`/`center`/`right`/`fill`/`justify`), vertical alignment, and/or `wrap_text` |
| `add_border` | no | Apply a border to all edges; style: `thin`/`medium`/`thick`/`dashed`/`dotted`/`double`; optional hex color |

### COM Tools — Sprint O + Phase 2 (5)

All require `EXCEL_ENABLE_COM=true`. `recalculate_workbook`, `export_as_pdf`, `hydrate_template`, `pivot_table(operation='create'|'refresh')`, and `export_range_as_csv` (write path) also require `confirm=True`; `pivot_table(operation='read')` and `export_range_as_csv` (console-only path) do not. **Windows-only** — requires `pywin32` (`pip install pywin32`). A licensed Excel installation must be present.

> **Security:**
> - `AutomationSecurity` set to `msoAutomationSecurityForceDisable` (3) before any workbook opens — macro execution blocked at COM layer.
> - `hydrate_template` validates all `cell_writes` strings against a formula-injection guard before COM entry (OWASP A03).
> - PDF output paths validated against `EXCEL_EXPORT_ROOTS` (separate from `EXCEL_ALLOWLIST_ROOTS`) before write.
> - `pivot_table` create/refresh: only `xlDatabase` local sources accepted; external ODBC/OData/web sources blocked (SSRF guard).
> - `export_range_as_csv` CSV output paths validated against `EXCEL_CSV_EXPORT_ROOTS` before write; cell values sanitised against CSV injection (OWASP A03 — values starting with `=`, `+`, `-`, `@` are prefixed with `'`).
> - `_ensure_com_gate()` now includes a `sys.platform` guard — raises `NotAllowedError` on non-Windows platforms before attempting any COM call.

| Tool | Save mode | EXCEL_ENABLE_WRITE? | Description |
|---|---|---|---|
| `recalculate_workbook` | in-place | **yes** | Force-recalculate via COM; `full_calc=True` rebuilds full dependency tree; saves in-place. Returns `{ok, path, calc_mode, elapsed_ms}` |
| `export_as_pdf` | read-only | no | Export sheet (`scope='sheet'`) or workbook (`scope='workbook'`) to PDF; both `path` and `output_path` must be absolute, `scope` must be exactly `sheet` or `workbook`, and `output_path` must be an absolute `.pdf` path in `EXCEL_EXPORT_ROOTS`; once validation succeeds, any stale target is cleared immediately before export so evidence always refers to fresh bytes |
| `hydrate_template` | SaveAs | **yes** | Write `cell_writes` ({ref, value} list) to template then SaveAs; `recalculate=True` default. Returns `{ok, output_path, writes_applied, recalculated}`; calling with `cell_writes=[]` performs a pure recalc/save-as flow |
| `pivot_table` | varies | varies | Create (`operation='create'`), read (`operation='read'`), or refresh (`operation='refresh'`) pivot tables. Only `xlDatabase` local sources. `create` requires `EXCEL_ENABLE_WRITE=true` + `confirm=True`; `read` is COM-gated but read-only; `refresh` requires `confirm=True` |
| `export_range_as_csv` | varies | **yes** (write path) | Export a cell range to CSV. `live=False` (default) uses openpyxl — no COM required; `live=True` uses Excel COM for accurate cached values. Console-only mode (no `output_path`) requires only `EXCEL_ENABLE_COM=true` when `live=True`; writing to disk additionally requires `EXCEL_ENABLE_WRITE=true` + `output_path` inside `EXCEL_CSV_EXPORT_ROOTS` + `confirm=True`. CSV injection guard applied to all cell values via `_sanitize_csv_value`. |

### Active-Workbook Bridge Tools — Sprint W4 (2)

Both tools require `EXCEL_ENABLE_COM=true`. Read-only — no write gate, no `confirm` parameter. Results are filtered to `EXCEL_ALLOWLIST_ROOTS`; workbooks outside the allowlist are silently excluded. **Windows-only** — requires `pywin32`. These tools attach to the user's *running* Excel instance via `GetActiveObject`; they do not launch a new hidden Excel process.

| Tool | Key response fields | Description |
|---|---|---|
| `get_active_workbook_context` | `status`, `workbook.name`, `workbook.path`, `workbook.sheets[]` | Return name, path, and sheet list for the workbook currently active in Excel. Status values: `ok` (data returned), `no_active_workbook` (Excel has no open file), `outside_allowlist` (active workbook is outside `EXCEL_ALLOWLIST_ROOTS`), `com_error` (workbook metadata could not be read). Requires `EXCEL_ENABLE_COM=true`. |
| `get_all_open_workbooks` | `status`, `workbooks[].name`, `workbooks[].path`, `workbooks[].sheets[]`, `count` | List all workbooks open in the running Excel instance. Filters to `EXCEL_ALLOWLIST_ROOTS` only; workbooks outside the allowlist are silently excluded — their paths and names are not exposed, not even as a count or redacted entry. Returns `{status: ok, workbooks: [...], count: N}`. Requires `EXCEL_ENABLE_COM=true`. |

**Session-mode writes — `cell` and `range_io`:**
Setting `session_mode=True` on `cell` (write) or `range_io` (write) routes the write directly to the active Excel workbook via COM, bypassing openpyxl file I/O. Requires `EXCEL_ENABLE_COM=true` + `EXCEL_ENABLE_WRITE=true` + `confirm=True`. Guards: formula values (`=`, `+`, `-`, `@` prefixes) are rejected before COM entry; `xl_app.Ready` is verified before access; the workbook `ReadOnly` flag is checked before the write. Ragged `values` rows in `range_io` are padded to `n_cols` with `None` for COM SAFEARRAY compatibility. The Excel instance is never terminated — `Quit()` is never called on the attached session.

### Chart Tool — Sprint O (1)

| Tool | Operations | Description |
|---|---|---|
| `chart` | `list` · `create` · `delete` · `update` | Manage charts: `operation='list'` (no confirm needed), `'create'` (bar/line/scatter/pie), `'delete'` (by 0-based index), `'update'` (replace data range). `'create'` accepts `data_address`, `target_address` (placement cell), `anchor`, `width_px`, `height_px`. |

### Render Verification

After build iterations are complete, use `export_as_pdf` with `scope='sheet'` (one sheet
at a time) or `scope='workbook'` (full workbook) to render through a live Excel COM
session and visually confirm that print areas, conditional formatting, charts, and formula
values appear correctly. See [docs/render-check-iterate.md](docs/render-check-iterate.md)
for the full iteration loop and export directory convention.

### Render-Check-Iterate Tools (3)

All require `EXCEL_ENABLE_COM=true`. Auto-save on success. `confirm=True` required.

| Tool | Description |
|---|---|
| `review_workbook_render` | Run structural quality checks (print area, empty sheets, stale formulas) before PDF export. Returns `passed`, `findings[]`, `checks_run` |
| `produce_export_evidence_bundle` | Aggregate `export_as_pdf` results + render review findings into a versioned JSON evidence bundle for issue tracker attachment |
| `export_changed_sheets_only` | Hash sheet content; re-export only changed sheets to a timestamped directory; returns `exported_sheets[]`, `new_hashes{}` |

### ACP (Artifact Context Packet) Tool (1)

Read-only. No write gate or `confirm` required.

| Tool | Description |
|---|---|
| `get_workbook_context` | Return a structured Artifact Context Packet (ACP) for a workbook at `index`, `focused`, or `deep` disclosure level. Never exposes cell values. Use before styling or export operations. |

### VBA Tools — COM (2)

Both require `EXCEL_ENABLE_COM=true`. `run_macro` additionally requires `EXCEL_ENABLE_MACROS=true` + `EXCEL_ENABLE_WRITE=true` + `confirm=True`. **Windows-only** — requires `pywin32`. Only `.xlsm` files (macro-enabled workbooks) are accepted. `list_macros` requires Trust Center access to the VBA project object model.

| Tool | EXCEL_ENABLE_MACROS? | Description |
|---|---|---|
| `list_macros` | no | List all VBA components in a workbook via the VBA project object model; returns component names and types. `.xlsm` only. |
| `run_macro` | **yes** | Execute a named VBA macro in an `.xlsm` workbook via Excel COM; `macro_name` validated against `^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*)?$`; single-quote injection guard on workbook filename. Requires `EXCEL_ENABLE_WRITE=true` + `confirm=True`. |

---

## MCP Resources

Nine read-only resources expose workbook structure and workbook-level inventories without requiring a tool call.
Most follow the `excelmcp://workbook/{path}/...` URI scheme; `excelmcp://meta/contract` is a path-free schema resource.

> **Path encoding:** `{path}` must be URL-encoded because Windows paths contain
> backslashes and colons. Use `urllib.parse.quote(path, safe='')` in Python or the
> equivalent in your client before constructing the URI.
> Example: `C:\Temp\file.xlsx` → `C%3A%5CTemp%5Cfile.xlsx`

| Resource URI template | Description | Key response fields |
|---|---|---|
| `excelmcp://workbook/{path}/sheets` | List all worksheets in the workbook with used-range dimensions | `sheets[].name`, `sheets[].min_row`, `sheets[].max_row`, `sheets[].min_col`, `sheets[].max_col`, `sheets[].empty` |
| `excelmcp://workbook/{path}/sheet/{name}` | Return structural info (dims, headers, tables) for a specific sheet | `sheet`, `min_row`, `max_row`, `min_col`, `max_col`, `headers[]`, `tables[]` |
| `excelmcp://workbook/{path}/charts` | Return embedded chart inventory for the workbook | `charts[].sheet`, `charts[].index`, `charts[].type`, `charts[].anchor`, `charts[].data_address` |
| `excelmcp://workbook/{path}/tables` | Return all Excel tables with names, refs, and column names | `tables[].sheet`, `tables[].name`, `tables[].ref`, `tables[].columns[]` |
| `excelmcp://workbook/{path}/named_ranges` | Return all named ranges defined in the workbook | `named_ranges[].name`, `named_ranges[].refers_to`, `named_ranges[].scope` |
| `excelmcp://workbook/{path}/validations` | Return workbook data-validation inventory | `validations[].sheet`, `validations[].sqref`, `validations[].type`, `validations[].formula1`, `validations[].formula2` |
| `excelmcp://workbook/{path}/metadata` | Return canonical structural `WorkbookMetadata` for the workbook (10 fields; rewired in Metadata Umbrella Sprint from `_MCP_META` to `build_workbook_metadata()`) | `path`, `workbook_name`, `sheet_names`, `active_sheet`, `sheet_count`, `has_formulas`, `last_modified`, `size_bytes`, `format`, `contract_version`, `policy` |
| `excelmcp://workbook/{path}/meta` | Legacy alias for `/metadata`; identical response shape via the same `build_workbook_metadata()` call | same as `/metadata` |
| `excelmcp://meta/contract` | `WorkbookMetadata` contract schema — field names, types, and descriptions. No path parameter. | `contract_version`, `fields[]`, `canonical_field_order[]` |

---

## MCP Prompts

Thirty-five built-in prompts provide ready-to-use orchestration instructions covering 100% of the tool surface.
Pipe raw JSON output from a read tool directly into the matching prompt parameter.

### Phase 2.0 — Analysis and Design (8 prompts)

| Prompt | Parameters | Use case |
|---|---|---|
| `summarize_sheet` | `sheet_json: str` — raw JSON from `read_sheet_all` or the `sheet/{name}` resource | Produce a concise business summary: metrics, trends, data quality issues, actionable insights |
| `find_anomalies` | `range_json: str` — raw JSON from `read_range` | Flag statistical outliers, missing values, impossible values, and duplicate rows in a range |
| `write_formula` | `description: str`, `column_layout: str` | Generate an Excel formula for the described calculation given the column layout |
| `compare_ranges` | `range_a: str`, `range_b: str` — two JSON range strings | Summarise values added, removed, and changed between two ranges; group by significance |
| `extract_structured_table` | `raw_range: str` — raw JSON from `read_range` | Convert a messy Excel range to a clean JSON table with inferred column types; output: `{"headers": [...], "rows": [[...]]}` |
| `design_workbook` | `business_context: str` | Plan a new workbook structure from a business description: suggest sheet names, columns, data types, and examples ready to pass to `create_workbook` |
| `suggest_formatting` | `sheet_json: str`, `report_purpose: str` — purpose as plain text | Given sheet data and report purpose, suggest a minimal professional formatting sequence (which tools to call, in which order, with what parameters) |
| `prepare_for_print` | `sheet_description: str`, `page_size: str = "A4"` | Given a sheet description and target page size, suggest a professional print configuration: `set_print_area` range, `freeze_panes` cell, and `set_column_width` values so the data fits neatly on the page |

### Phase 2.0 — Orchestration (5 prompts)

| Prompt | Parameters | Use case |
|---|---|---|
| `audit_workbook` | `path: str` | Discovery prompt: orchestrate the core workbook resources to understand a workbook before editing — sheets, tables, charts, named ranges; the current resource surface totals seven read-only resources including validations and workbook metadata |
| `chart_from_data` | `sheet_json: str`, `insight: str` | Guide for creating a chart to visualise a specific business insight; selects chart type from time/category/part-to-whole/correlation and calls `chart create` |
| `build_data_validation` | `field_description: str` | Generate a complete `data_validation` add tool call JSON spec (address, type, operator, formula1/2, error/prompt messages) for the described field |
| `pivot_workflow` | `source_description: str` | Step-by-step guide for creating a pivot table via COM; includes a SUMIFS/COUNTIFS fallback for environments where COM is unavailable |
| `bulk_format_report` | `sheet_json: str`, `brand_colors: str = ""` | Orchestrate 7-step bulk styling for a report sheet using the consulting navy/grey palette (or custom hex colors) |

### Phase 2.1 — Tool Surface Coverage (12 prompts)

| Prompt | Parameters | Use case |
|---|---|---|
| `manage_sheets` | `workbook_description: str` | Audit and reorganise all sheets: add, rename, delete, copy, set tab color/visibility; orchestrates `sheet`, `sheet_properties`, `bulk_copy_sheet`, `save` |
| `reshape_table` | `table_description: str`, `change_description: str` | Insert/delete rows and columns, set row heights and column widths, hide columns; orchestrates `rows`, `cols`, `set_row_heights`, `set_column_widths`, `set_column_visibility` |
| `style_dashboard_range` | `range_description: str`, `purpose: str` | Apply polished dashboard formatting in the correct order (fill → align → number format → border → banding); orchestrates `apply_style`, `apply_alignment`, `apply_number_format`, `add_border` |
| `named_range_workflow` | `intent: str` | List, create, and verify named ranges; orchestrates `named_range` (list/write/read) with best-practice naming conventions |
| `annotate_workbook` | `items_json: str` — JSON array of `{cell, comment, url}` objects | Add cell comments and hyperlinks in bulk; orchestrates `bulk_add_comments`, `bulk_add_hyperlinks`, `get_hyperlinks`, `remove_hyperlink` |
| `format_header_section` | `header_description: str` | Merge title row, style header row, and apply auto-filter; orchestrates `merge`, `apply_style`, `apply_alignment`, `auto_filter` |
| `import_and_style_csv` | `csv_path: str`, `target_sheet: str` | Import a CSV, auto-fit columns, convert to an Excel Table, freeze the header row; orchestrates `import_csv_to_sheet`, `set_column_widths`, `create_table`, `freeze_panes`, `save` |
| `export_workbook_report` | `path: str`, `sheet_names: str` — comma-separated | Finalise print areas and export each sheet plus the full workbook to PDF; orchestrates `set_print_area`, `freeze_panes`, `export_as_pdf`, `save` |
| `lock_template` | `path: str`, `editable_ranges: str` — comma-separated | Protect all sheets and workbook structure while keeping specified ranges editable; orchestrates `protect`; includes password guidance |
| `grow_table` | `path: str`, `sheet_name: str`, `table_name: str` | Discover table schema, append rows in correct column order, verify table auto-expansion; orchestrates `list_tables`, `read_table`, `append_rows`, `get_used_range` |
| `highlight_exceptions` | `range_address: str`, `rule_description: str` | Apply CellIs or formula-based conditional formatting with traffic-light colours; orchestrates `apply_conditional_format` |
| `validate_and_recalculate` | `path: str` | Validate formula syntax, force-recalculate via COM, inspect used ranges, and verify KPI cached values; orchestrates `validate_formula_syntax`, `recalculate_workbook`, `get_used_range`, `evaluate_formula` |

### Phase 2.2+ — Extended Coverage (7 prompts)

| Prompt | Parameters | Use case |
|---|---|---|
| `read_write_and_inspect_range` | `path: str`, `sheet_name: str`, `address: str` | Read, write, copy (values+formats), inspect cell metadata, and find-replace in a range; orchestrates `range_io`, `copy_range`, `get_cell_metadata`, `find_replace` |
| `annotate_and_configure_sheet` | `path: str`, `sheet_name: str` | Manage cell comments and hyperlinks, and configure sheet display (gridlines, headers, zoom, tab color); orchestrates `cell_comment`, `hyperlink`, `sheet_properties` |
| `style_and_write_range` | `range_description: str`, `style_intent: str` | Write data with inline formatting, apply style to non-contiguous ranges, or format-paint from a template; orchestrates `write_range_with_format`, `apply_style`, `copy_range_format` |
| `duplicate_sheet_template` | `path: str`, `source_sheet: str`, `new_names: str` — comma-separated | Copy a template sheet to multiple named project sheets in one call; orchestrates `bulk_copy_sheet` |
| `discover_server_and_manage_named_ranges` | `path: str` | Discover server capabilities and feature flags, then list/create/verify named ranges; orchestrates `capabilities`, `named_range` |
| `export_or_hydrate` | `intent: str` | Export a sheet or workbook to PDF, or fill a template with project data and save as a new file; orchestrates `export_as_pdf`, `hydrate_template` (both require `EXCEL_ENABLE_COM=true`) |
| `configure_workbook_calculation` | `path: str` | Set `auto`/`manual`/`autoNoTable` calculation mode and the `full_calc_on_load` flag; orchestrates `set_workbook_calc_mode` |

### Metadata Umbrella Sprint — Governance Prompts (2 prompts)

| Prompt | Parameters | Use case |
|---|---|---|
| `workbook_metadata_workflow` | `path: str` — absolute workbook path | 4-step metadata-first workflow: load `/metadata` resource → review 10 canonical fields → act on `strict`/`lenient` policy → confirm before subsequent tool calls |
| `metadata_contract_reference` | *(none)* | Returns the `WorkbookMetadata` contract table (all 10 canonical fields with types and descriptions) plus the active `EXCEL_METADATA_POLICY` note |

---

### Removed Parameter Aliases

Seven backward-compatible parameter aliases (`freeze_at`, `src_range`, `dst_range`, `table_range`, `anchor_cell`, `end_cell`, `range_address`) were removed in the Contract Stabilization sprint. Only the canonical parameter names (`address`, `src_address`, `dst_address`, `start_address`, `end_address`) are accepted. See the tool descriptions above for current parameter names.

---

### Metadata Governance (audit_workbook update)

The `audit_workbook` prompt now references the updated `/metadata` resource which returns structural `WorkbookMetadata` (10 canonical fields) rather than the `_MCP_META` hidden-sheet payload. Use `excelmcp://meta/contract` to inspect the field contract before building automation that consumes metadata responses.

---

## UX Patterns and Recipes

### Recipe 1 — Read a Budget Spreadsheet

```
1. capabilities()
   → confirm version, phase, backend, available tools, and live prompts payload

2. sheet(path, operation="list")
   → {"sheets": ["Summary", "Q1", "Q2", "Q3", "Q4"], "count": 5}

3. get_used_range(path, sheet="Summary")
   → {"min_row": 1, "max_row": 42, "min_col": 1, "max_col": 8, "empty": false}

4. range_io(path, sheet="Summary", operation="read", address="A1:H42")
   → 2-D array of all budget data

5. named_range(path, operation="list")
   → ["TotalRevenue", "COGS", "GrossMargin"] with their cell references
```

### Recipe 2 — Update Cells

```
1. cell(path, sheet="Q1", operation="write", address="C10", value=125000, confirm=True)
   → auto-saved immediately

2. range_io(path, sheet="Q1", operation="write", address="A2",
       values=[["Jan", 10000], ["Feb", 12000], ["Mar", 11500]], confirm=True)
   → auto-saved immediately

3. save(path, confirm=True)
   → safe no-op if nothing pending in the write cache
```

### Recipe 3 — Import Data from CSV

```
1. import_csv_to_sheet(path, sheet="RawData",
       csv_path="C:/data/export.csv",
       overwrite=True, confirm=True)
   → auto-saved immediately; existing data replaced

2. read_sheet_all(path, sheet="RawData")
   → verify the imported rows

3. list_tables(path)
   → check if any Excel tables are defined over the imported range
```

### Recipe 4 — Sheet Management

```
1. sheet(path, operation="add", sheet_name="Q5 Forecast", position=5, confirm=True)
   (auto-saved immediately)

2. bulk_copy_sheet(path, source_sheet="Q4", new_names=["Q5 Template"], confirm=True)
   (auto-saved immediately)

3. sheet(path, operation="rename", sheet_name="Sheet1", new_name="Index", confirm=True)
   (auto-saved immediately)

4. sheet(path, operation="delete", sheet_name="Temp", confirm=True)
   (auto-saved immediately)

5. save(path, confirm=True)   # optional defensive no-op if nothing is pending
```

### Recipe 5 — Format a Summary Table

```
1. apply_style(path, sheet="Summary", address="A1:H1",
       bold=True, bg_color="1F3864", font_color="FFFFFF", confirm=True)
   (auto-saved)

2. apply_alignment(path, sheet="Summary", address="A1:H1",
       horizontal="center", vertical="center", confirm=True)

3. apply_number_format(path, sheet="Summary", address="C2:H20",
       format_code="#,##0.00", confirm=True)

4. set_column_widths(path, sheet="Summary", columns=["A"], width=12, confirm=True)
   (auto-saved)

5. add_border(path, sheet="Summary", address="A1:H20",
       border_style="thin", color="AAAAAA", confirm=True)

6. save(path, confirm=True)
   → persists apply_alignment, apply_number_format, add_border changes
```

### Recipe 6 — Create and Delete a Chart

```
1. chart(path, sheet="Revenue", operation="create", chart_type="bar",
       data_address="A1:B13",
       title="Monthly Revenue 2026",
       anchor="D1", confirm=True)
   → auto-saved; chart appears at column D, row 1

2. chart(path, sheet="Revenue", operation="list")
   → {"charts": [{"index": 0, "title": "Monthly Revenue 2026",
                  "type": "BarChart", "anchor": "3,0"}], "count": 1}

3. chart(path, sheet="Revenue", operation="delete", chart_index=0, confirm=True)
   → chart removed; auto-saved
```

### Recipe 7 — Read Workbook Structure via MCP Resources

```
1. GET excelmcp://workbook/C%3A%5CTemp%5Cmyfile.xlsx/sheets
   → all sheets with their used-range dimensions (min/max row and col)

2. GET excelmcp://workbook/C%3A%5CTemp%5Cmyfile.xlsx/sheet/Sales
   → sheet dims + first-row headers + list of Excel tables

3. GET excelmcp://workbook/C%3A%5CTemp%5Cmyfile.xlsx/named_ranges
   → [{"name": "TotalRevenue", "refers_to": "Sales!$B$2:$B$13", "scope": "workbook"}, ...]
```

### Recipe 8 — Analyse a Sheet with the `summarize_sheet` Prompt

```
1. read_sheet_all(path, sheet="Sales")
   → raw 2-D values JSON

2. Use prompt: summarize_sheet(sheet_json=<values from step 1>)
   → model produces: key metrics, trends, data quality issues, actionable insights
```

---

## What Is NOT Available (Known Limitations)

| Feature | Status | Notes |
|---|---|---|
| Creating a new `.xlsx` from scratch | ✅ Available | `create_workbook(path, sheets=["Sheet1"], overwrite=False, confirm=True)` |
| Formula recalculation | ✅ Available (COM, Sprint G) | `recalculate_workbook` — opens workbook in Excel COM, recalculates (dirty cells or full tree), saves in-place. Requires `EXCEL_ENABLE_COM=true`, `EXCEL_ENABLE_WRITE=true`, Windows-only. |
| `set_workbook_calc_mode` | ✅ Available (Phase 2.6) | Set auto/manual calculation + fullCalcOnLoad flag |
| Creating charts (basic types) | ✅ Available | `chart(operation='create'|'list'|'delete'|'update')` — see Chart Tool section |
| Advanced chart editing (series, axes) | ❌ Not available | Requires COM (Phase 2 win32com) |
| Pivot tables (create / refresh / read data) | ✅ Available (Sprint H COM) | `pivot_table(operation='create'|'refresh'|'read')` — requires `EXCEL_ENABLE_COM=true`, Windows-only. |
| Named range definitions (write) | ✅ Available | `named_range(path, operation='write', name=..., sheet=..., address=..., confirm=True)` — Sprint O |
| Conditional formatting (write) | ✅ Available (Phase 2.8) | `apply_conditional_format` — CellIsRule / FormulaRule with injection guard; `copy_range_format` — format-painter across sheets |
| Cell borders | ✅ Available | `add_border` — thin/medium/thick/dashed/dotted/double with optional hex color |
| Gradient fills, pattern fills | ❌ Not available | Basic solid fills via `apply_style`; gradient/pattern fills require COM (Phase 2) |
| Writing merged cells | ✅ Available | `merge(path, sheet, operation='merge'|'unmerge', address=..., confirm=True)` — Sprint O |
| Row/column freeze (freeze panes) | ✅ Available | `freeze_panes(path, sheet, address=..., confirm=True)` — Phase 2.2. Pass `address=None` to unfreeze. |
| AutoFilter on a data range | ✅ Available | `auto_filter(path, sheet, address, confirm=True)` — Phase 2.2. Pass `address=None` explicitly to clear. |
| Insert rows / columns | ✅ Available | `rows(operation='insert'|'delete')` / `cols(operation='insert'|'delete')` — Sprint O |
| Worksheet protection | ✅ Available | `protect(path, scope='sheet', sheet=..., password=..., confirm=True)` — Sprint O |
| Print area definition | ✅ Available | `set_print_area(path, sheet, address, confirm=True)` — Phase 2.3 |
| Cell comments (add / delete / read) | ✅ Available | `bulk_add_comments` (write) — Sprint D; `cell_comment(operation='get'|'delete')` — Sprint O |
| Sheet tab colour / visibility state | ✅ Available | `sheet_properties(path, sheet, operation='get'|'set', ...)` — Sprint O |
| Data validation rules | ✅ Available | `bulk_add_data_validation` (write) — Sprint D; `data_validation(operation='get'|'add')` — Sprint O |
| Workbook structure protection | ✅ Available | `protect(path, scope='workbook', password=..., confirm=True)` — Sprint O |
| Password-protected workbooks | ❌ Not available | openpyxl cannot open encrypted files |
| Multiple workbooks open simultaneously | ❌ Not available | Server holds one workbook per process |
| Macros / VBA | ✅ Available (COM) | `list_macros` (enumerate VBA components), `run_macro` (execute a named macro); `.xlsm` only; requires `EXCEL_ENABLE_COM=true`; `run_macro` additionally requires `EXCEL_ENABLE_MACROS=true` + `EXCEL_ENABLE_WRITE=true` + `confirm=True` |
| Live COM / Excel.Application interaction | ✅ Available (Sprint O) | 4 COM tools: `recalculate_workbook`, `export_as_pdf`, `hydrate_template`, `pivot_table`; plus live formula evaluation via `evaluate_formula(live=True)`. Requires `EXCEL_ENABLE_COM=true`, Windows-only (`pywin32`). |
| Insert / delete individual columns by letter | ✅ Available | `cols(path, sheet, operation='insert'|'delete', col=..., confirm=True)` — Sprint O |
| `set_column_visibility` | ✅ Available (Phase 2.6) | Hide/show columns by letter |
| Streaming large files | ❌ Not available | Full workbook is loaded into memory; very large files may be slow |

---

## Security & Governance

- All file paths are validated against `EXCEL_ALLOWLIST_ROOTS` before any I/O. Traversal
  above an allowed root is blocked with `"Path not in allowlist"`.
- Write tools require `EXCEL_ENABLE_WRITE=true` AND (for destructive operations)
  `confirm=True`. The two-factor pattern prevents accidental mutation from misconfigured clients.
- `stdout` carries MCP JSON-RPC traffic only. All server logs go to `stderr`.
- No network calls. All operations are local file I/O only.
- **Formula injection guard (openpyxl tools)**: `write_range_with_format` and `import_csv_to_sheet` validate every cell value against the injection guard — values starting with `=`, `+`, `-`, `@` are rejected (OWASP A03); plain negative numbers are exempt. `hyperlink` (`add` / `bulk_add` operations) validates `display_text` against the same guard.
- **File size pre-check**: workbook size is checked before load to prevent unexpectedly large files from consuming server memory.
- **UNC path guard:** `_check_path()` and `_check_csv_path()` both reject paths starting with `\\` (backslash UNC, e.g. `\\server\share`) or `//` (forward-slash UNC, e.g. `//server/share`) before `Path.resolve()` is called — preventing NTLM relay and path-normalisation bypass attacks.
- **COM tools** (`EXCEL_ENABLE_COM=true` required): `AutomationSecurity` is set to `msoAutomationSecurityForceDisable` (3) on the Excel Application object before any workbook is opened — macro execution is blocked at the COM layer. `hydrate_template` validates all `cell_writes` string values against a formula-injection guard before entering the COM context (strings starting with `=`, `+`, `-`, `@` are rejected — OWASP A03). PDF export output paths are validated against `EXCEL_EXPORT_ROOTS` before any write.
- **ACP adapter injection hardening (OWASP A03):** `_acp_adapter.py` sanitizes all user-controlled string fields embedded in ACP annotations — `sheet_names`, `table_names`, and `lineage_summary["workbook_name"]` — via `_sanitize_text_field` imported from `mcpshared`. The shared helper strips formula-prefix characters (`=`, `+`, `-`, `@`, `\t`, `\r`, `\n`), enforces a 255-character max length, and rejects NUL bytes.
- **UNC path guard (W2-005):** `_check_path` and `_check_csv_path` in `_core.py` reject paths starting with `\\` — prevents NTLM credential relay via UNC network shares.
- **CSV import null-byte guard (R-002):** `_check_csv_path` in `_core.py` rejects paths containing `\x00` before allowlist evaluation — prevents null-byte injection in CSV import operations.

---

## Performance Notes

### PERF-OPT (v0.2.9) — LRU-Cached Style Factories

Internal optimization: `apply_cell_style` and `apply_alignment` now use
`@functools.lru_cache`-backed factory functions for `Font`, `Alignment`, and
`PatternFill` construction. Identical style parameter sets reuse the same
Python object across cells rather than allocating a new instance per cell.

**Measured speedups** (versus pre-optimization baseline):
| Operation | 1k cells | 5k cells | 10k cells |
|-----------|----------|----------|-----------|
| `apply_cell_style` | ~1.2x | ~1.2x | ~1.7x |
| `apply_alignment` | ~1.8x | ~2.1x | ~1.8x |

No API surface changes in PERF-OPT itself. Current README inventory is 64 tools. Historical totals: v0.3.0 raised total to 66; v0.4.0 raised total to 72; v0.5.0 raised total to 73; v0.6.0 raised total to 78; v0.7.0 raised total to 82; Sprint N consolidation reduced total to 76; Sprint O consolidation reduced total to 50; Sprint R + Metadata Umbrella raised total to 54; Phase 2 CSV Export raised total to 55; Active-Workbook Bridge Sprint raised total to 57; Governed Write + W4 additions raised total to 59; Snapshot + VBA sprint raised total to 64.

**Governance overhead** (amortized, µs/call): `_check_write` ~0.4 µs,
`_check_confirm` ~0.1 µs, `_check_path` varies by allowlist length.

---

## Server Module Layout

The Refactor Sprint (2026-03-11) decomposed 4 monolithic files into 17 focused modules. `server.py` is the MCP entry point and imports all modules at startup via side-effect imports. Sprint W4 added `server_com_session.py`; Snapshot + VBA sprint added `server_snapshot.py` and `server_com_vba.py` (total: 64).

### Server Entry Points

| Module | Lines | Role |
|--------|-------|------|
| `server.py` | ~120L | FastMCP instance, resource registrations, imports prompt modules, `main()` |
| `server_com.py` | ~31L | Thin re-exporter for COM tools (imports `server_com_export`, `server_com_pivot`, `server_com_session`) |
| `server_com_export.py` | ~409L | `export_as_pdf`, `export_range_as_csv` tools |
| `server_com_pivot.py` | ~425L | `recalculate_workbook`, `hydrate_template`, `pivot_table` tools |
| `server_com_session.py` | ~120L | `get_active_workbook_context`, `get_all_open_workbooks` — active-workbook bridge tools |
| `server_com_vba.py` | — | `list_macros`, `run_macro` — VBA COM tools |
| `server_snapshot.py` | — | `snapshot_workbook`, `diff_workbooks` — snapshot/diff tools |
| `server_prompts_analysis.py` | ~47L | Analysis prompt definitions |
| `server_prompts_format.py` | ~319L | Formatting prompt definitions |
| `server_prompts_workflow.py` | ~426L | Workflow prompt definitions |
| `server_prompts_mgmt.py` | ~242L | Workbook management prompt definitions |

### Tool-Group Server Modules (Sprint P — unchanged)

| Module | Tools | Role |
|--------|-------|------|
| `server_io.py` | 18 | I/O tools: `capabilities`, `create_workbook`, `sheet`, `get_used_range`, `cell`, `range_io`, `read_sheet_all`, `named_range`, `get_cell_metadata`, `evaluate_formula`, `list_tables`, `read_table`, `append_rows`, `find_replace`, `import_csv_to_sheet`, `save`, `workbook_metadata`, `consolidate_ranges` |
| `server_format.py` | 8 | Formatting tools: `apply_number_format`, `apply_style`, `apply_alignment`, `add_border`, `apply_format_to_sheet_list`, `copy_range_format`, `write_range_with_format`, `apply_conditional_format` |
| `server_ops.py` | 16 | Operations tools: `merge`, `freeze_panes`, `auto_filter`, `rows`, `cols`, `protect`, `set_print_area`, `cell_comment`, `sheet_properties`, `data_validation`, `validate_formula_syntax`, `copy_range`, `create_table`, `set_column_visibility`, `set_workbook_calc_mode`, `hyperlink` |
| `server_batch.py` | 6 | Batch tools: `bulk_copy_sheet`, `set_row_heights`, `set_column_widths`, `bulk_add_comments`, `fill_formula_range`, `bulk_range_write` |
| `server_chart.py` | 1 | Chart tool (`chart`) — registers with the shared `mcp` instance via side-effect import |
| `server_review.py` | 3 | Render-review tools (`review_workbook_render`, `produce_export_evidence_bundle`, `export_changed_sheets_only`) |
| `server_acp.py` | 1 | ACP tool: `get_workbook_context` (Artifact Context Packet — index/focused/deep disclosure; never exposes cell values) |
| `server_com_session.py` | 2 | Active-workbook bridge: `get_active_workbook_context`, `get_all_open_workbooks` — read-only COM session tools |
| `server_snapshot.py` | 2 | Snapshot tools: `snapshot_workbook`, `diff_workbooks` |
| `server_com_vba.py` | 2 | VBA COM tools: `list_macros`, `run_macro` |
| `_server_instance.py` | — | Holds the single shared `mcp = FastMCP(...)` instance; imported by all `server_*.py` modules to guarantee one registry |

### Implementation Modules

| Module | Lines | Role |
|--------|-------|------|
| `_format_cell.py` | ~251L | Cell style, font, fill, alignment, border helpers |
| `_format_dimensions.py` | ~304L | Column/row width/height |
| `_format_bulk.py` | ~402L | `apply_style_to_ranges`, `apply_format_to_sheet_list`, `copy_range_format` |
| `_format.py` | ~8L | Thin re-export router |
| `_data_cell.py` | ~127L | Cell metadata, formula evaluation |
| `_data_tables.py` | ~214L | Tables + named ranges |
| `_data_annotations.py` | ~392L | Comments + hyperlinks |
| `_data_write.py` | ~188L | `fill_formula_range`, `write_range_with_format` |
| `_data.py` | ~9L | Thin re-export router |

All implementation modules (`_core.py`, `_ooxml.py`, `_com.py`, `_advanced.py`, `_sheet.py`, etc.) contain no FastMCP imports — pure logic, callable by both the MCP layer and tests.

---

## Changelog

See [`../CHANGELOG.md`](../CHANGELOG.md) for the full release history.

### v0.3 — capabilities_v2 + server decomposition ([5cd6dad](https://github.com/dosev-ai/mcp-office/commit/5cd6dad))

`capabilities()` now returns a machine-readable `capabilities_v2` object (TOOL_REGISTRY 64 entries,
17 dispatcher vocabularies, per-tool gate metadata, tool modes). `format()` dispatcher added.
Server layer decomposed into focused modules. `session_mode` stubs removed. See
[`../shared/design/capabilities_schema_v1.md`](../shared/design/capabilities_schema_v1.md) for schema spec.

---

### Governed Write — Active-Workbook Sprint (2026-03-13)

**Governed COM write for `cell` and `range_io`; 35 new tests (D-01..D-12, E-01..E-11, F/G-series); tool count unchanged at 57**

- `cell` (write) and `range_io` (write) extended with `session_mode: bool = False` parameter (`server_io.py`, `_io.py`).
- `session_mode=True` routes writes via COM to the active Excel workbook; `session_mode=False` (default) preserves existing openpyxl path — no behaviour change for existing callers.
- Requires `EXCEL_ENABLE_COM=true` + `EXCEL_ENABLE_WRITE=true` + `confirm=True` when `session_mode=True`.
- `_active_wb.py` extended: `_write_cell_via_com()`, `_write_range_via_com()`, `_is_valid_a1()` added.
- Guards: formula values (`=`, `+`, `-`, `@` prefixes) rejected; `xl_app.Ready` checked before access; `ReadOnly` flag checked before write; `Quit()` never called on the attached session.
- Ragged `values` rows padded to `n_cols` with `None` for COM SAFEARRAY safety (`range_io` write path).
- 35 new unit tests (`test_governed_write.py`): D-01..D-12, E-01..E-10, E-11, F/G-series.
- Total tool count: **57** (unchanged)

### Active-Workbook Bridge Sprint — W4 (2026-03-13)

**2 new COM session tools; 1161 tests passing; tool count 55 → 57**

- New tool `get_active_workbook_context` (`server_com_session.py`): returns `{status, workbook: {name, path, sheets[]}}` for the workbook active in the running Excel session. Requires `EXCEL_ENABLE_COM=true`. Workbooks outside `EXCEL_ALLOWLIST_ROOTS` return `status: outside_allowlist`. No write gate.
- New tool `get_all_open_workbooks` (`server_com_session.py`): lists all allowlisted workbooks open in Excel; out-of-scope workbooks are silently excluded. Returns `{status, workbooks[], count}`. No write gate.
- New module `_active_wb.py`: COM helper — `_com_active_app()` context manager (attaches to live Excel session via `GetActiveObject`), `_is_in_allowlist()`, `get_active_workbook_com()`, and `_check_file_not_open_in_excel()` (LC-2 conflict guard).
- `_io.py` and `_sheet.py` updated: `_check_file_not_open_in_excel()` added to all write paths (LC-2 guard — blocks openpyxl write when the target file is currently open in Excel).
- `runtime_config.py` updated: `get_session_mode()` added — reads `EXCEL_SESSION_MODE` env var; emits `UserWarning` when session mode is enabled without `EXCEL_ENABLE_COM=true`.
- New env var: `EXCEL_SESSION_MODE` (default `false`) — enables session-mode routing for the active-workbook bridge.
- Total tool count: **57**

### Refactor Sprint (2026-03-11)

**Monolith decomposition — no tools added or removed; 1069 tests passing; 4 project tracker tech-debt items closed**

- Decomposed monolithic `server.py` (was 1048L → ~120L) into 4 `server_prompts_*.py` modules (`server_prompts_analysis.py`, `server_prompts_format.py`, `server_prompts_workflow.py`, `server_prompts_mgmt.py`)
- Decomposed `server_com.py` (was 859L → ~31L re-export router) into `server_com_export.py` (~409L) + `server_com_pivot.py` (~425L)
- Decomposed `_format.py` (was 898L → ~8L re-export router) into `_format_cell.py` (~251L) + `_format_dimensions.py` (~304L) + `_format_bulk.py` (~402L)
- Decomposed `_data.py` (was 868L → ~9L re-export router) into `_data_cell.py` (~127L) + `_data_tables.py` (~214L) + `_data_annotations.py` (~392L) + `_data_write.py` (~188L)
- All 4 project tracker tech-debt items closed
- Total tool count: **55** (unchanged)

### Phase 2 — CSV Export + COM Hardening (2026-03-11)

**1 new tool (`export_range_as_csv`); 2 new security helpers in `_com.py`; tool count 54 → 55**

- `export_range_as_csv` (`server_com.py`): dual-path CSV export — `live=False` (openpyxl, no COM required) or `live=True` (Excel COM, requires `EXCEL_ENABLE_COM=true`). Writing to `output_path` requires `EXCEL_ENABLE_WRITE=true` + path inside `EXCEL_CSV_EXPORT_ROOTS` + `confirm=True`. Returns `{ok, sheet, address, rows_exported, output_path}` or `{ok, csv}` for console-only mode.
- `_ensure_com_gate()` hardened: `sys.platform` guard added — raises `NotAllowedError` immediately on non-Windows platforms before any `win32com` import is attempted.
- New helper `_check_csv_export_path` (`_com.py`): validates CSV `output_path` against `EXCEL_CSV_EXPORT_ROOTS`; mirrors the `_check_export_path` pattern used for PDF outputs.
- New helper `_sanitize_csv_value` (`_com.py`): CSV injection guard (OWASP A03) — cell string values starting with `=`, `+`, `-`, `@` are prefixed with `'` before serialisation.
- New env var: `EXCEL_CSV_EXPORT_ROOTS` (no default) — comma-separated absolute folder paths permitted as CSV output destinations.
- Total tool count: **55**

### Metadata Umbrella Sprint (2026-03-11)

**43 new tests (15+9+15+4); 143 total regression pass; tool count 54 unchanged; 2 new prompts; 2 new resources**

- New module `_metadata_contract.py`: `WorkbookMetadata` TypedDict SSOT, `CONTRACT_VERSION="1.0"`, 10 `CANONICAL_FIELDS`, `build_workbook_metadata()` — validates path via allowlist, stats filesystem, loads workbook read-only, scans for formulas
- New module `runtime_config.py`: `EXCEL_METADATA_POLICY` env var (`"strict"` default / `"lenient"`);
  `load_runtime_config()`, `get_effective_policy()`, `is_strict()`, `get_contract_version()`
- New module `_lineage.py`: `FIELD_ALIASES` dict (8 legacy→canonical mappings); `normalize_metadata()` / `map_lineage_fields()` — normalises dicts that use legacy key names to the canonical `WorkbookMetadata` shape
- New module `server_metadata_prompts.py`: `workbook_metadata_workflow` prompt (4-step metadata-first flow), `metadata_contract_reference` prompt (contract table), `excelmcp://meta/contract` resource (schema JSON, no path parameter)
- `server_io.py` updated: `/metadata` and `/meta` resources rewired to `build_workbook_metadata()`; `get_workbook_metadata()` shared impl injects `policy` from `EXCEL_METADATA_POLICY`; `capabilities()` now returns `metadata_contract_version` and `metadata_policy` fields
- `server.py` updated: side-effect import `excelmcp.server_metadata_prompts` added
- New test files: `test_metadata_contract.py` (15), `test_runtime_config.py` (9), `test_lineage.py` (15), `test_capabilities_drift.py` (4)

### DA Security Hardening — `probe_excel_reopen` (2026-03-06)

**`_da_com.py` security hardening — no tools added or removed in that sprint; current README inventory is 51 tools**

- `probe_excel_reopen`: now raises `NotAllowedError` when `EXCEL_ENABLE_COM != "true"` — was previously silently bypassed
- `probe_excel_reopen`: non-temp paths denied by default when `EXCEL_ALLOWLIST_ROOTS` is unset (explicit deny-by-default)
- `probe_excel_reopen`: `AutomationSecurity = 3` (`msoAutomationSecurityForceDisable`) set before any workbook open — macro execution blocked at COM layer
- Return dict no longer includes `temp_path` key
- New test file: `test_unit_da_com.py` (3 unit tests covering COM gate, allowlist gate, and security flags)

### Phase 5 QA — Security Hardening Bug Sprint (2026-03-05)

**Commit `b694ac6` on `staging` — no tools added or removed in that sprint; current README inventory is 51 tools; 815 → 846 tests**

- `write_range_with_format` + `import_csv_to_sheet`: formula injection guard applied to all cell values before write; plain negative numbers exempt in CSV context (OWASP A03)
- `hyperlink` (`add` / `bulk_add`): `display_text` parameter now validated against the injection guard
- File size pre-check added before opening workbooks
- `protect` (sheet scope): password hashing corrected — openpyxl hash algorithm now applied correctly
- `find_replace`: empty `find_text` and `None` cell values handled without raising
- New test file: `test_security_guards.py` (30 tests covering all injection-guard and file-size scenarios)
- ruff lint clean

### Sprint P — Server Decomposition (2026-03-04)

**Structural refactor: `server.py` (1363 lines) decomposed into 4 tool-group modules. Current README inventory is 51 total tools after the later addition of `workbook_metadata`.**

- `server_io.py` (new, 17 tools): `capabilities`, `create_workbook`, `sheet`, `get_used_range`, `cell`, `range_io`, `read_sheet_all`, `named_range`, `get_cell_metadata`, `evaluate_formula`, `list_tables`, `read_table`, `append_rows`, `find_replace`, `import_csv_to_sheet`, `save`, `workbook_metadata`
- `server_format.py` (new, 8 tools): `apply_number_format`, `apply_style`, `apply_alignment`, `add_border`, `apply_format_to_sheet_list`, `copy_range_format`, `write_range_with_format`, `apply_conditional_format`
- `server_ops.py` (new, 16 tools): `merge`, `freeze_panes`, `auto_filter`, `rows`, `cols`, `protect`, `set_print_area`, `cell_comment`, `sheet_properties`, `data_validation`, `validate_formula_syntax`, `copy_range`, `create_table`, `set_column_visibility`, `set_workbook_calc_mode`, `hyperlink`
- `server_batch.py` (new, 5 tools): `bulk_copy_sheet`, `set_row_heights`, `set_column_widths`, `bulk_add_comments`, `fill_formula_range`
- At Sprint P refactor time, `server.py` slimmed from 1363 lines into a lean entry point containing resources, prompts, side-effect imports, and `main()` only; the decomposition enforces ≤500-line threshold per tool module
- Test file structural splits (no logic changes): `test_sprint_n_b.py` (from `test_sprint_n.py` — CommitB+C tests), `test_unit_com_p2.py` (from `test_unit_com.py` — ExportWorkbook+Hydrate+OpenTemplate tests), `test_unit_com_h2.py` (from `test_unit_com_h.py` — CreatePivot+Refresh+ReadPivot tests)
- Current total tool count: **51**

### Sprint O — Tool Consolidation (2026-03-04)

**76 → 50 tools: 26 single-operation tools removed; superseded by unified dispatch tools**

| Removed tools | Unified dispatch replacement |
|---|---|
| `list_sheets`, `add_sheet`, `rename_sheet`, `delete_sheet` | `sheet(operation='list'\|'add'\|'rename'\|'delete')` |
| `read_cell`, `write_cell` | `cell(operation='read'\|'write')` |
| `read_range`, `write_range` | `range_io(operation='read'\|'write')` |
| `get_named_ranges`, `read_named_range`, `write_named_range` | `named_range(operation='list'\|'read'\|'write')` |
| `list_merged_cells`, `merge_cells`, `unmerge_cells` | `merge(operation='list'\|'merge'\|'unmerge')` |
| `get_cell_comment`, `delete_cell_comment` | `cell_comment(operation='get'\|'delete')` |
| `get_sheet_properties`, `set_sheet_properties` | `sheet_properties(operation='get'\|'set')` |
| `get_data_validation_info`, `bulk_add_data_validation` | `data_validation(operation='get'\|'add')` |
| `get_hyperlinks`, `add_hyperlink`, `bulk_add_hyperlinks`, `remove_hyperlink` | `hyperlink(operation='get'\|'add'\|'remove')` |
| `apply_cell_style`, `apply_style_to_ranges` | `apply_style(ranges=[...])` |
| `insert_rows`, `delete_rows` | `rows(operation='insert'\|'delete')` |
| `insert_cols`, `delete_cols` | `cols(operation='insert'\|'delete')` |
| `protect_sheet`, `protect_workbook` | `protect(scope='sheet'\|'workbook')` |
| `create_chart`, `list_charts`, `delete_chart`, `update_chart_data` | `chart(operation='create'\|'list'\|'delete'\|'update')` |
| `export_sheet_as_pdf`, `export_workbook_as_pdf` | `export_as_pdf(scope='sheet'\|'workbook')` |
| `open_template_and_recalculate` | `hydrate_template(cell_writes=[])` |
| `evaluate_formula_live` | `evaluate_formula(live=True)` |
| `create_pivot_table`, `refresh_pivot_tables`, `read_pivot_table` | `pivot_table(operation='create'\|'read'\|'refresh')` |

Tool count: 76 → 50 (server.py=45, server_com.py=4, server_chart.py=1). All unified dispatch tools accept the same payloads as their single-operation predecessors — no calling-pattern migration needed for existing valid inputs.

### Sprint N — Tool Consolidation (2026-03-04)

**82 → 76 tools: 6 single-operation tools removed; superseded by bulk counterparts**

| Removed tool | Superseded by |
|---|---|
| `set_column_width` | `set_column_widths` |
| `set_row_height` | `set_row_heights` |
| `copy_sheet` | `bulk_copy_sheet` |
| `add_hyperlink` | `bulk_add_hyperlinks` |
| `add_cell_comment` | `bulk_add_comments` |
| `add_data_validation` | `bulk_add_data_validation` |

All bulk replacements handle single-item inputs — no calling-pattern change needed.

### v0.6.0 — 2026-03-02

**Sprint G — 78-tool release: 5 COM Phase 2 tools (Windows-only, `pywin32`)**

- New infrastructure module: `_com.py` — shared COM helpers (`_com_excel_app`, `_open_workbook`, `_close_workbook`, `_ensure_com_gate`, `_check_export_path`, `_guard_cell_value`); `EXCEL_ENABLE_COM` gate; `AutomationSecurity=3` macro-execution block
- New server module: `server_com.py` — 5 COM tools registered via side-effect import at startup (mirrors `server_chart.py` pattern)
- New env vars: `EXCEL_ENABLE_COM` (COM activation gate), `EXCEL_EXPORT_ROOTS` (PDF output allowlist)
- `recalculate_workbook` (`server_com.py`): force-recalculate workbook via COM; `full_calc=False` (dirty cells) or `full_calc=True` (`CalculateFull`); saves in-place; returns `{ok, path, calc_mode, elapsed_ms}`
- `export_sheet_as_pdf` (`server_com.py`): export single sheet to PDF read-only; `quality` 0/1; output validated against `EXCEL_EXPORT_ROOTS`; returns `{ok, sheet, output_path, quality}`
- `export_workbook_as_pdf` (`server_com.py`): export entire workbook to PDF read-only; all visible sheets in print order; returns `{ok, output_path, quality, source}`
- `hydrate_template` (`server_com.py`): write `cell_writes` ({ref, value} list) to template then `SaveAs` new file; formula-injection guard on all string values before COM entry; `recalculate=True` default; returns `{ok, output_path, writes_applied, recalculated}`
- `open_template_and_recalculate` (`server_com.py`): open template, `CalculateFull()`, `SaveAs` new file; for volatile function workbooks (TODAY, NOW, RAND, OFFSET); returns `{ok, output_path, template, recalculated: true}`
- New test files: `test_unit_com.py` (31 unit tests, mock-based, no Excel required), `test_integration_com.py` (integration, `pytest -m integration`)
- **E2E gap-fix** (commits `2e7e178`, `4f75ff5`): ruff lint fixes; `recalculated: true` added to `open_template_and_recalculate` return; `test_unit_com.py` expanded 22 → 31 tests; total suite **602 passing**
- Total tool count: **78** (was 73)

### v0.5.1 — 2026-03-02

**Sprint F — Internal quality hardening: modular server architecture (no new tools)**

- `_ooxml.py`: v-before-f regex bugfix — corrected attribute ordering in `<f>` element OOXML output (internal, no user-visible change)
- `_format.py`: gate-order hardening for `apply_cell_style` + `apply_alignment` — write-enable check now always precedes confirm check (internal, no API change)
- New files: `_server_instance.py` (shared `mcp` FastMCP instance) + `server_chart.py` (chart tools module); `server.py` registers chart tools via side-effect import at startup
- New test: `test_server_chart_module.py` — verifies modular import contract
- Total tool count: **73** (unchanged)

### v0.5.0 — 2026-03-02

**Sprint E — 73-tool release: `update_chart_data` + `create_chart` size/anchor extensions (Phase 2.9)**

- `update_chart_data` (chart write, new): replace the data range of an existing chart by 0-based `chart_index` (see `list_charts`); `data_range` e.g. `'A1:B5'` — must contain `':'`; requires `EXCEL_ENABLE_WRITE=true` + `confirm=True`; returns `{ok, sheet, chart_index, data_range}`
- `create_chart` (chart write, extended): new optional params `anchor` (overrides `target_cell`, backward-compatible alias), `width_px`, `height_px` (chart size in pixels, 1–5000); fully backward-compatible with existing callers
- Total tool count: **73** (was 72)
- New test files: `test_charts_sprint_e.py`, `test_da_ooxml_strip.py` (split from `test_da_finalize.py`)
- Commit: `2a1115b` on `staging`

### v0.4.0 — 2026-02-27

**Sprint D — 72-tool release: 7 bulk-operation / formatting tools (Phase 2.8)**

- Sprint D tools added (7): `fill_formula_range`, `write_range_with_format`, `bulk_add_comments`, `bulk_add_data_validation`, `apply_format_to_sheet_list` (v2 extension), `copy_range_format`, `apply_conditional_format`
- `fill_formula_range` (_data.py): fill a formula across a range using `openpyxl.formula.translate.Translator` for correct cell-reference relativisation; returns `{ok, sheet, anchor_cell, cells_filled}`
- `write_range_with_format` (_data.py): write 2-D values + apply style/number_format/alignment atomically; returns `{ok, sheet, start_cell, rows_written}`
- `bulk_add_comments` (_data.py): add multiple cell comments in one call with fail-fast validation; returns `{ok, sheet, comments_added}`
- `bulk_add_data_validation` (_advanced.py): apply multiple validation rules at once; `validation_type` restricted to type allowlist to prevent injection; returns `{ok, sheet, rules_applied}`
- `apply_format_to_sheet_list` v2 (_format.py): extends Sprint C tool with `column_widths` and `row_heights` dict params; fully backward-compatible
- `copy_range_format` (_format.py): format-painter — copies cell formatting (not values) from src_range to dst_range, cross-sheet supported; returns `{ok, src_sheet, dst_sheet, cells_copied}`
- `apply_conditional_format` (_advanced.py): applies CellIsRule or FormulaRule conditional formatting; formula strings validated against injection guard before write; returns `{ok, sheet, range_address, rule_type}`

### v0.3.0 — 2026-02-26

**Sprint C — 66-tool release: 6 bulk-operation tools + `add_hyperlink` security backport**

- Sprint C tools added (6): `bulk_copy_sheet`, `set_row_heights`, `apply_style_to_ranges`, `set_column_widths`, `bulk_add_hyperlinks`, `apply_format_to_sheet_list`
- `bulk_copy_sheet` (_sheet.py): copy a source sheet to multiple named copies in one call; validates unique names, ≤31 chars, no invalid chars; returns `{ok, source_sheet, sheets_created, names}`; charts/images/print settings NOT copied (openpyxl limitation)
- `set_row_heights` (_format.py): uniform mode (`rows`/`row_range` + `height`) or per-row mode (`row_specs`); returns `{ok, sheet, rows_set}`
- `apply_style_to_ranges` (_format.py): apply font/fill to non-contiguous ranges in one call; `EXCEL_MAX_RANGE_CELLS` enforced across all ranges; returns `{ok, sheet, ranges_processed, cells_styled}`
- `set_column_widths` (_format.py): explicit list or `col_range`; manual `width` or `autofit=True` (content-length approximation); returns `{ok, sheet, columns_set, widths}`
- `bulk_add_hyperlinks` (_data.py): add multiple hyperlinks atomically; URL scheme allowlist `{https, http, mailto}`; `file://` opt-in via `EXCEL_ALLOW_FILE_HYPERLINKS=true`; `javascript:`/`vbscript:`/`data:` always blocked; returns `{ok, sheet, links_added}`
- `apply_format_to_sheet_list` (_format.py): apply style/number_format/alignment to the same address across multiple sheets; `sheets=None` auto-selects all non-`_`-prefixed sheets; returns `{ok, sheets_processed, cells_formatted}`
- **Security backport**: `add_hyperlink` now validates URL scheme — rejects `javascript:`, `vbscript:`, `data:` (OWASP A03/A05); `file://` requires `EXCEL_ALLOW_FILE_HYPERLINKS=true`
- New env var: `EXCEL_ALLOW_FILE_HYPERLINKS` (default `false`) — opt-in for `file://` hyperlink scheme

### v0.2.9 — 2026-02-26
- **Engine**: `apply_cell_style` / `apply_alignment` style factories wrapped
  with `@functools.lru_cache` — identical style param sets reuse cached objects,
  reducing per-cell Python allocation overhead.
- **Benchmarks**: cProfile analysis added to benchmark suite; governance gate
  overhead quantified; `profile_report.txt` committed as evidence.
- No new tools. No API changes.

### v0.2.8 — 2026-02-26

**Phase 2.7 — 60-tool release: hyperlink management + `zoom_scale` extension for `set_sheet_properties`**

- Phase 2.7 tools added (3): `get_hyperlinks`, `add_hyperlink`, `remove_hyperlink`
- `get_hyperlinks` (read): return all hyperlinks in a worksheet; returns `{"sheet": ..., "count": ..., "hyperlinks": [{address, url, display_text, tooltip}]}`
- `add_hyperlink` (write): add or update a hyperlink on a cell; optional `display_text` (replaces cell value) and `tooltip`; requires `EXCEL_ENABLE_WRITE=true` + `confirm=True`
- `remove_hyperlink` (write): remove a hyperlink from a cell; requires `EXCEL_ENABLE_WRITE=true` + `confirm=True`
- `set_sheet_properties` extended: `zoom_scale` parameter added (integer 10–400) to set sheet zoom level
- Commit: `4b03caf` on `staging`

### v0.2.7 — 2026-02-26

**Docs sync: document Phase 2.6 tools `set_column_visibility` + `set_workbook_calc_mode`; version bump**

- README tool table updated to reflect all 57 tools delivered in Phase 2.6
- Added `set_column_visibility` to Write / Structural Tools: hide or show one or more columns by letter list; requires `EXCEL_ENABLE_WRITE=true` + `confirm=True`
- Added `set_workbook_calc_mode` to Write / Structural Tools: set `auto`/`manual`/`autoNoTable` calc mode and `fullCalcOnLoad` flag; requires `EXCEL_ENABLE_WRITE=true` + `confirm=True`
- No code changes; docs-only release

### v0.2.6 — 2026-02-26

**Phase 2.6 — 57-tool release: read data validations, read cell comment, optional path for formula validation; set_column_visibility, set_workbook_calc_mode**

- Phase 2.6 tools added (2): `get_data_validation_info`, `get_cell_comment`
- `get_data_validation_info` (read): return all data validation rules configured on a sheet; returns `{"sheet": ..., "count": ..., "validations": [...]}`
- `get_cell_comment` (read): return the comment on a cell, including text and author; returns `{"address": ..., "has_comment": bool, "text": ..., "author": ...}`
- Ergonomic fix: `validate_formula_syntax` `path` parameter is now optional (defaults to `""`)
- 265 pytest tests green (unit + smoke)

### v0.2.5 — 2026-02-26

**Phase 2.5 — 53-tool release: formula validation, range copy, merged cell listing, native table creation**

- Phase 2.5 tools added (4): `validate_formula_syntax`, `copy_range`, `list_merged_cells`, `create_table`
- `validate_formula_syntax` (read): check formula string syntax without writing to the workbook; returns `{"valid": bool, "error": str | null}`
- `copy_range` (write): copy values and formats from a source range to a destination range within the same sheet; requires `EXCEL_ENABLE_WRITE=true` + `confirm=True`
- `list_merged_cells` (read): list all merged cell ranges in a worksheet; returns `{"sheet": ..., "merged_ranges": [...]}`
- `create_table` (write): create a native Excel table (ListObject) over a range; applies TableStyleMedium9 with row stripes; enforces name uniqueness, min 2 rows, valid Excel identifier for `table_name`; requires `EXCEL_ENABLE_WRITE=true` + `confirm=True`
- 243 pytest tests green (unit + smoke)

### v0.2.4 — 2026-02-25

**Phase 2.4 — 49-tool release: cell comments, sheet properties, data validation, workbook protection**

- Phase 2.4 tools added (6): `add_cell_comment`, `delete_cell_comment`, `get_sheet_properties`, `set_sheet_properties`, `add_data_validation`, `protect_workbook`
- All tools implemented via openpyxl (file-based, no COM); full `confirm=True` + write-gate discipline
- 197 pytest unit tests + 20 smoke tests green

### v0.2.3 — 2026-02-24

**Phase 2.3 — 43-tool, 8-prompt release: insert rows/cols, worksheet protection, print area**

- Phase 2.3 structural tools added (4): `insert_rows`, `insert_cols`, `protect_sheet`, `set_print_area`
- Range address validation (N3): `write_named_range` now rejects garbage input with a clear `ValidationError`
- Prompt added: `prepare_for_print` — generates a professional print configuration from a sheet description
- 186 pytest tests green (unit + smoke)

### v0.2.0 — 2026-02-24

**Phase 2.0 — 33-tool release: formatting + chart tools; `list_charts` bug fix**

- Phase 2.0 formatting tools added (6): `apply_number_format`, `apply_cell_style`,
  `apply_alignment`, `set_column_width`, `set_row_height`, `add_border`
- Phase 2.0 chart tools added (3): `create_chart`, `list_charts`, `delete_chart`
- **Bug fix** (`list_charts`): openpyxl `Title` objects are now correctly serialised
  to plain strings via `title.tx.rich.p[0].r[0].t`; added `index` field; simplified
  anchor format to `"col,row"`
- Showcase file built end-to-end exercising all 33 tools
  (`C:\Temp\excelmcp_showcase.xlsx` — UserStories, SalesData, Products sheets)
- Issue-tracker bug entries created under Excel MCP epic
   (internal tracking references redacted in this public repository)

### v0.1.5 — 2026-02-23

**Phase 1.5 — 24-tool release + E2E UX sprint (20/20 PASS)**

- Full write toolset added (12 tools): `write_cell`, `write_range`, `append_rows`,
  `add_sheet`, `save`, `find_replace`, `rename_sheet`, `delete_sheet`, `copy_sheet`,
  `delete_rows`, `delete_cols`, `import_csv_to_sheet`
- E2E UX sprint — 20 user stories validated, all PASS:
  - **GAP-03:** `read_sheet_all` response wrapped in `{"sheet": ..., "values": [[...]]}` dict
  - **CODE-GAP-01:** workbook cache key changed from `path` string to `(path, sheet)` tuple
    to isolate per-sheet state correctly
  - **GAP-02:** `_flush_if_dirty` helper added for consistent auto-save in write tools
  - **GAP-07:** `has_header` dead-code parameter removed from `import_csv_to_sheet`
- 77 pytest tests green (unit + smoke)
- Commits: `c2ac564`, `f7c49b8` on `staging`

### v0.1.0 — 2026-02-23

**Phase 1 initial release — 12 read tools**

- `capabilities`, `list_sheets`, `get_used_range`, `read_cell`, `read_range`,
  `read_sheet_all`, `get_named_ranges`, `read_named_range`, `get_cell_metadata`,
  `evaluate_formula`, `list_tables`, `read_table`
- openpyxl backend, FastMCP stdio transport
- Path allowlist enforcement, `EXCEL_MAX_RANGE_CELLS` guard
- Unit and smoke test suite
- Commit: `d94d2cd` on `staging`

---

## Backlog

### Phase 2 — COM / win32com (Windows-only)

- [x] Live formula recalculation via `Excel.Application`
- [x] Pivot table create/read/refresh
- [x] PDF export via `export_as_pdf`
- [ ] Advanced chart editing (series data, axis labels, chart type change)

### Phase 1 / 2.0 Extensions (openpyxl)

- [ ] Create a new `.xlsx` from a blank template
- [ ] Conditional formatting read
- [ ] Streaming read for very large files (openpyxl read-only mode)
- [ ] Gradient and pattern fill write support
- [ ] Merge / unmerge cells
