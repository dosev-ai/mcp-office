# wordmcp — Word MCP Server

`wordmcp` is a Python MCP stdio server for reading and editing `.docx` files using
[python-docx](https://python-docx.readthedocs.io/). It is part of the
mcp-office suite and operates entirely at the file level — no
Word application or COM layer is required (Phase 1). Write operations are gated behind
`WORD_ENABLE_WRITE=true` (environment variable) and a `confirm=True` parameter for
destructive calls, preventing accidental mutations.

---

## Status

> Phase 1 + Phase 2 COM — v0.4.0 · **51 tools** · 456 tests (non-integration) · document assembly, review/evidence, security hardening, and full dispatcher surface

---

## Quick Start

```powershell
# Install (editable, with dev deps)
pip install -e "./wordmcp[dev]"

# Run all tests
pytest wordmcp/tests/ -v

# Run by marker
pytest wordmcp/tests/ -m unit -v
pytest wordmcp/tests/ -m smoke -v
pytest wordmcp/tests/ -m security -v
pytest wordmcp/tests/ -m "not integration" -v
```

### VS Code mcp.json Registration

Add the following entry to
`%APPDATA%\Code - Insiders\User\mcp.json` (or `Code\User\mcp.json` for stable VS Code):

```json
"word-wordmcp": {
    "type": "stdio",
    "command": "python",
    "args": ["-m", "wordmcp.server"],
    "env": {
        "PYTHONPATH": "C:\\path\\to\\mcp-office\\wordmcp\\src",
        "WORD_ALLOWLIST_ROOTS": "C:\\Users\\yourname\\Documents,C:\\Temp",
        "WORD_ENABLE_WRITE": "true"
    }
}
```

---

## Environment Variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `WORD_ALLOWLIST_ROOTS` | **yes** | — | Comma-separated absolute folder paths the server may access. Any path outside these roots is blocked with a `"Path not in allowlist"` error. **Container default:** The published Docker image sets `WORD_ALLOWLIST_ROOTS=/tmp` by default. Mount your documents into `/tmp` or override this environment variable to point to your volume mount path. |
| `WORD_ENABLE_WRITE` | no | unset (read-only) | Must be `true` to enable any mutating tool. If unset or any other value, all write calls return an error even when `confirm=True` is passed. |
| `WORD_ENABLE_COM` | no | unset (COM disabled) | Must be `true` to enable COM tools (`manage_tracked_changes`, `export_document`). Requires `pywin32` and Microsoft Word on Windows. |
| `WORD_MAX_TEXT_CHARS` | no | `50000` | Maximum characters returned by `export_as_text`. If output is truncated, `truncated=True` is set in the response and a notice is appended to the text field. |
| `WORD_MAX_FILE_MB` | no | `50` | Maximum `.docx` file size in MB; files exceeding this limit are rejected before loading into memory. |
| `WORD_MAX_CACHE_DOCS` | no | `20` | Maximum documents held in the open-document cache (LRU eviction). |

---

## Tools Reference

### Read Tools (8)

No write-gate required. Safe to call without `WORD_ENABLE_WRITE`.

| Tool | Description |
|---|---|
| `capabilities` | Return metadata about this MCP server phase, backend, and available tools. |
| `read_document` | Read a Word document, routing to the appropriate handler via `scope`. Scope values: `full` (default — summary counts), `metadata` (core properties), `headings` (flat heading list), `outline` (hierarchical tree), `section` (one section by heading index; requires `section_index`), `paragraphs` (flat paragraph list), `context` (token-budget overview). Replaces and deprecates `get_document_metadata`, `list_headings`, `get_document_outline`, `read_section`, `list_paragraphs`, `get_document_context`. |
| `get_document_metadata` | Return full document metadata: title, author, subject, keywords, dates, revision. Times are as stored in the .docx file (typically UTC, no timezone suffix). **Deprecated** — use `read_document(scope='metadata')`. |
| `get_document_context` | Return a summary-only context packet for a .docx file. Payload is read-only and limited to a metadata subset, aggregate counts, and heading text; it excludes paragraph bodies, run text, table cell contents, comments, and tracked changes. **Deprecated** — use `read_document(scope='context')`. |
| `list_paragraphs` | Return summary of all body paragraphs: index, style, text_preview, outline_level. Table-cell paragraphs are excluded. **Deprecated** — use `read_document(scope='paragraphs')`. |
| `read_paragraph` | Return full detail for one body paragraph including all run properties. |
| `list_tables` | Return summary of all tables in the document: index, rows, cols, style. **Deprecated** — use `table(operation='list')`. |
| `read_table` | Return full cell data for one table as a 2D row-major array. Merged cells are repeated (python-docx behaviour). **Deprecated** — use `table(operation='read')`. |

### Write Tools (7)

All require `WORD_ENABLE_WRITE=true` **and** `confirm=True`.

> **Atomic save:** In Phase 1 every write tool saves the file immediately after mutation.
> `save` is provided for forward-compatibility with Phase 2 batched workflows — calling
> it after any Phase 1 write is a harmless double-save.

| Tool | Description |
|---|---|
| `add_paragraph` | Append a paragraph to the end of the document. `style`: optional Word paragraph style name (e.g. `'Normal'`, `'Body Text'`). Defaults to `'Normal'`. |
| `add_heading` | Append a heading paragraph. `level` 0 = `'Title'`, 1–9 = `'Heading N'`. Level 10+ raises an error. |
| `add_page_break` | Insert a page break at the end of the document body. |
| `add_table` | Append a table to the end of the document. `data`: optional 2-D list of strings (rows × cols). `style`: optional Word table style name (default `'Normal Table'`). |
| `insert_image` | Insert an inline image at the end of the document body. `image_path` must be under `WORD_ALLOWLIST_ROOTS`. Supported formats: PNG, JPEG, GIF, BMP, TIFF. **SVG not supported.** `width_inches`: optional resize (aspect ratio preserved). |
| `find_replace` | Find and replace text across body paragraphs and table cells. Supports cross-run matching within a paragraph. Optional `paragraph_index` limits replacement to one body paragraph; optional 1-based `occurrence` replaces only the Nth document occurrence. Case-sensitive. |
| `save` | Explicitly persist the document to disk. Requires `WORD_ENABLE_WRITE=true` and `confirm=True`. |

### Sprint C — Navigation and Discovery (5)

`create_document` requires `WORD_ENABLE_WRITE=true` and `confirm=True`. The remaining four tools are read-only.

| Tool | Description |
|---|---|
| `create_document` | Create a new empty Word document at the given path. Requires `WORD_ENABLE_WRITE=true` and `confirm=True`. Raises ToolError if the file already exists. |
| `list_headings` | List all headings in the document with their outline level (1–9), paragraph index, and style name. Returns an empty list if the document has no headings. **Deprecated** — use `read_document(scope='headings')`. |
| `read_section` | Read the content of a section identified by its heading paragraph index. Returns the heading and all body paragraphs until the next same-or-higher-level heading. Use `list_headings()` to discover heading indices. **Deprecated** — use `read_document(scope='section', section_index=N)`. |
| `search_text` | Search for text in document paragraphs and table cells. Returns matching paragraph indices, context (first 200 chars), and match count per paragraph. `max_results`: cap on total matches (1–1000, default 100). **Deprecated** — use `document(operation='search_text')`. |
| `list_styles` | List all styles available in the document. `style_type`: optional filter — `'paragraph'` \| `'character'` \| `'table'` \| `'numbering'`. Returns name, type, and whether the style is built-in. **Deprecated** — use `style(operation='list')`. |

### Sprint C — Edit Operations (6)

All require `WORD_ENABLE_WRITE=true` **and** `confirm=True`.

| Tool | Description |
|---|---|
| `apply_style` | Apply a named style to a paragraph by index. Use `list_styles()` to discover available styles. Raises ToolError for unknown style or out-of-range index. **Deprecated** — use `style(operation='apply')`. |
| `update_paragraph` | Replace the text content of a paragraph identified by index. Preserves paragraph style — only the runs (text) are replaced. **Deprecated** — use `paragraph(operation='update')`. |
| `delete_paragraph` | Delete a paragraph from the document by index. **Warning:** all subsequent paragraph indices shift down by 1 after deletion. **Deprecated** — use `paragraph(operation='delete')`. |
| `update_table_cell` | Set the text content of a specific table cell. Use `list_tables()` to discover table count; use `read_table()` to see row/col structure. **Deprecated** — use `table(operation='update_cell')`. |
| `set_document_properties` | Set document core properties (metadata): title, author, subject, keywords, category. At least one property must be provided. **Deprecated** — use `content(operation='set_properties')`. |
| `set_paragraph_format` | Apply spacing and line-spacing to a paragraph's ParagraphFormat. `paragraph_index` (required), `space_before` (pt), `space_after` (pt), `line_spacing` (multiplier; 1.0=single, 1.5=one-and-a-half, 2.0=double). Optional `table_cell` dict (`row`, `col`) targets a paragraph inside a table cell. **Deprecated** — use `paragraph(operation='set_format')`. |

### Sprint C — Content (4)

`get_headers_footers` is read-only. `add_list`, `insert_paragraph`, and `bulk_add_paragraphs` require `WORD_ENABLE_WRITE=true` and `confirm=True`.

> **Removed:** `export_to_markdown` is no longer a standalone tool. Use `export_document(format="md")` instead.

| Tool | Description |
|---|---|
| `add_list` | Append a bulleted or numbered list to the document. `list_type`: `'bullet'` (default) \| `'number'`. `items`: list of strings, max 500. `level`: indent level 0–8 (default 0). |
| `insert_paragraph` | Insert a new paragraph at a specific position in the document body. `paragraph_index=0` inserts before the first paragraph; `paragraph_index=len(paragraphs)` appends to the end. `style`: optional Word paragraph style name. |
| `bulk_add_paragraphs` | Add multiple paragraphs in a single operation (one save). `paragraphs`: list of `{"text": str, "style": optional str}` dicts. Max 500. More efficient than repeated `add_paragraph` calls. |
| `get_headers_footers` | Return the header and footer text for each document section. Includes whether each is linked to the previous section. |

### Sprint C — oxml (3)

`manage_comments('list')` is read-only. `manage_comments('add'/'resolve'/'delete')`, `add_hyperlink`, and `add_footnote` require `WORD_ENABLE_WRITE=true` and `confirm=True`.

| Tool | Description |
|---|---|
| `manage_comments` | Manage document comments (oxml implementation). `operation`: `'list'` (read-only) \| `'add'` \| `'resolve'` \| `'delete'` (all write ops require write gate). `paragraph_index` anchors the comment to that paragraph. |
| `add_hyperlink` | Add a clickable hyperlink to a paragraph. `url` must begin with `http://` or `https://` (`javascript:` and `file://` schemes are blocked). `text`: display text (defaults to url if empty). |
| `add_footnote` | Add a footnote reference to a paragraph. Requires the document to have an existing footnotes part (open in Word and add one footnote manually first if absent). |

### Review and Evidence Tools (4)

Two tools are read-only. `write_review_findings` requires `WORD_ENABLE_WRITE=true` and `confirm=True`. All four tools reuse `WORD_ALLOWLIST_ROOTS` — no new environment variables.

| Tool | Description |
|---|---|
| `get_document_outline` | Extract heading hierarchy, paragraph count, and table count from a document. Read-only. **Deprecated** — use `read_document(scope='outline')`. |
| `review_document` | Run up to 4 structural checks against the document: `style_consistency`, `heading_hierarchy`, `word_count`, `table_structure`. Returns a findings list per check. Read-only. Accepts optional `checks` list to run a subset. |
| `write_review_findings` | Append a review findings bundle to a `.jsonl` evidence file. Requires `WORD_ENABLE_WRITE=true` and `confirm=True`. Output path must be within `WORD_ALLOWLIST_ROOTS`. |
| `export_review_evidence` | Package a completed findings file into an evidence bundle suitable for audit trail / sign-off workflow. Read-only. |

### Export — 1 tool

`export_document` is the unified export surface. Format routing:

| `format` | Backend | Gate | Notes |
|---|---|---|---|
| `"txt"` | python-docx (`export_as_text`) | read-only | Output capped at `WORD_MAX_TEXT_CHARS`. `output_path` not used. |
| `"md"` | python-docx (`export_to_markdown`) | read-only | Converts headings, bullets, tables, bold/italic. `output_path` not used. |
| `"pdf"` | Word COM (`document_com.export_document`) | `WORD_ENABLE_COM=true` + `WORD_ENABLE_WRITE=true` + `confirm=True` | `output_path` required. |
| `"html"` | Word COM (`document_com.export_document`) | `WORD_ENABLE_COM=true` + `WORD_ENABLE_WRITE=true` + `confirm=True` | `output_path` required. |

| Tool | Description |
|---|---|
| `export_document` | Export document to a specified format. `format`: `Literal["txt", "md", "pdf", "html"]` (default `"txt"`). `output_path` required only for `pdf`/`html`. Windows-only for COM formats. **Deprecated** — use the `export` dispatcher tool. |

### Bulk Tools

All require `WORD_ENABLE_WRITE=true` and `confirm=True`. These live in `_server_handlers_docx_advanced.py`.

| Tool | Description |
|---|---|
| `bulk_update_paragraphs` | Update multiple paragraphs by index in a single save. `updates`: list of `{"paragraph_index": int, "new_text": str}`. Max 500 updates per call. Per-item errors collected; partial success supported. **Deprecated** — use `paragraph(operation='bulk_update')`. |
| `bulk_update_table_cells` | Update multiple table cells in a single save. `updates`: list of `{"table_index": int, "row": int, "col": int, "new_text": str}`. Max 200 updates per call. Per-item errors collected; partial success supported. **Deprecated** — use `table(operation='bulk_update_cells')`. |

### Document Assembly (3 tools)

`bulk_find_replace` and `manage_hyperlinks` require `WORD_ENABLE_WRITE=true` and `confirm=True`. `verify_no_placeholders` is read-only.

| Tool | Description |
|---|---|
| `bulk_find_replace` | Replace multiple `{{TOKEN}}` placeholders in a Word document in a single pass. Requires `WORD_ENABLE_WRITE=true` and `confirm=True`. Returns replaced count, per-token results, and any tokens not found in the document. |
| `manage_hyperlinks` | Update hyperlink URLs and labels for named display-text entries in a Word document. Requires `WORD_ENABLE_WRITE=true` and `confirm=True`. Only `http`, `https`, and `mailto` URL schemes are accepted. Returns update status per display-text key. |
| `verify_no_placeholders` | Verify that no unreplaced `{{TOKEN}}` placeholders remain in a Word document. Read-only. Returns pass/fail status and list of any residual tokens found in body paragraphs and table cells. Headers and footers are not scanned. |

### Dispatcher Tools (8 tools — Dispatcher Sprint)

These unified dispatcher tools were added to reduce surface area. Each routes to the appropriate individual tool via an `operation` or `scope` parameter. The individual tools they replace remain registered for backward compatibility (noted as **Deprecated** in each section above).

#### `paragraph` — Paragraph operations dispatcher

| `operation` | Gate | Delegates to |
|---|---|---|
| `list` | read-only | `list_paragraphs` |
| `read` | read-only | `read_paragraph` |
| `add` | write-gated | `add_paragraph` |
| `add_heading` | write-gated | `add_heading` |
| `insert` | write-gated | `insert_paragraph` |
| `update` | write-gated | `update_paragraph` |
| `delete` | write-gated | `delete_paragraph` |
| `set_format` | write-gated | `set_paragraph_format` |
| `bulk_add` | write-gated | `bulk_add_paragraphs` |
| `bulk_update` | write-gated | `bulk_update_paragraphs` |

#### `table` — Table operations dispatcher

| `operation` | Gate | Delegates to |
|---|---|---|
| `list` | read-only | `list_tables` |
| `read` | read-only | `read_table` |
| `add` | write-gated | `add_table` |
| `update_cell` | write-gated | `update_table_cell` |
| `bulk_update_cells` | write-gated | `bulk_update_table_cells` |

#### `style` — Style operations dispatcher

| `operation` | Gate | Delegates to |
|---|---|---|
| `list` | read-only | `list_styles` |
| `apply` | write-gated | `apply_style` |

#### `content` — Content operations dispatcher (all write-gated)

| `operation` | Delegates to |
|---|---|
| `add_list` | `add_list` |
| `add_page_break` | `add_page_break` |
| `insert_image` | `insert_image` |
| `add_hyperlink` | `add_hyperlink` |
| `add_footnote` | `add_footnote` |
| `find_replace` | `find_replace` |
| `set_properties` | `set_document_properties` |

#### `document` — Document utility dispatcher

| `operation` | Gate | Delegates to |
|---|---|---|
| `save` | write-gated | `save` |
| `search_text` | read-only | `search_text` |
| `get_headers_footers` | read-only | `get_headers_footers` |

#### `review` — Review operations dispatcher

| `operation` | Gate | Delegates to |
|---|---|---|
| `review` | read-only | `review_document` |
| `write_findings` | write-gated | `write_review_findings` |
| `export_evidence` | read-only | `export_review_evidence` |
| `manage_comments` | mixed | `manage_comments` |
| `manage_tracked_changes` | mixed (COM required) | `manage_tracked_changes` |

#### `export` — Export dispatcher (all write-gated, requires `WORD_ENABLE_WRITE=true` and `confirm=True`)

| `scope` | Backend | Output requirement |
|---|---|---|
| `pdf` | Word COM (requires `WORD_ENABLE_COM=true`) | `output_path` must end in `.pdf` |
| `txt` | python-docx | `output_path` must end in `.txt` |
| `markdown` | python-docx | `output_path` must end in `.md` |

#### `read_document` (scope dispatcher — see Read Tools section above)

Routes to seven different read handlers via `scope`. See the Read Tools section for the full scope reference.

---

### Document Assembly (3 tools)

`bulk_find_replace` and `manage_hyperlinks` require `WORD_ENABLE_WRITE=true` and `confirm=True`. `verify_no_placeholders` is read-only.

| Tool | Description |
|---|---|
| `bulk_find_replace` | Replace multiple `{{TOKEN}}` placeholders in a Word document in a single pass. Requires `WORD_ENABLE_WRITE=true` and `confirm=True`. Returns replaced count, per-token results, and any tokens not found in the document. |
| `manage_hyperlinks` | Update hyperlink URLs and labels for named display-text entries in a Word document. Requires `WORD_ENABLE_WRITE=true` and `confirm=True`. Only `http`, `https`, and `mailto` URL schemes are accepted. Returns update status per display-text key. |
| `verify_no_placeholders` | Verify that no unreplaced `{{TOKEN}}` placeholders remain in a Word document. Read-only. Returns pass/fail status and list of any residual tokens found in body paragraphs and table cells. Headers and footers are not scanned. |

### COM Tools — Phase 2 (1 tool)

Requires `WORD_ENABLE_COM=true`. Write operations also require `WORD_ENABLE_WRITE=true` and `confirm=True`.
Windows-only — requires Microsoft Word desktop and `pywin32`. See **COM Tools (Phase 2)** below
for install instructions.

| Tool | Description |
|---|---|
| `manage_tracked_changes` | Manage document tracked changes via COM. `operation`: `'list'` \| `'accept_all'` \| `'reject_all'`. `'list'` is read-only; others require `WORD_ENABLE_WRITE=true` and `confirm=True`. |

> **Removed:** `export_as_pdf`, `accept_all_track_changes`, `reject_all_track_changes`, `list_tracked_changes` (deprecated backward-compat wrappers) have been removed. Use `export_document` and `manage_tracked_changes` directly.

---

## COM Tools (Phase 2)

Word COM tools automate the Microsoft Word desktop application via `win32com` to perform
operations beyond what `python-docx` supports at the file level.

**Install:**

```powershell
pip install wordmcp[com]
```

**Requirements:**
- Microsoft Word desktop app on Windows
- `pywin32 >= 306`
- `WORD_ENABLE_COM=true` environment variable

---

## UX Patterns and Recipes

### Recipe 1 — Inspect a Document

```
1. read_document(path)
   → paragraph_count, table_count, image_count, word_count, title, author

2. list_paragraphs(path)
   → all body paragraphs with index, style, text_preview, outline_level

3. read_paragraph(path, paragraph_index=3)
   → full run detail: bold, italic, font_name, font_size, color_hex, text

4. export_as_text(path)
   → full plain-text dump (body order preserved, capped at WORD_MAX_TEXT_CHARS)
```

### Recipe 2 — Append Content

```
1. add_heading(path, text="Executive Summary", level=1, confirm=True)
   → {"paragraph_index": 42, "style": "Heading 1", "text": "Executive Summary"}

2. add_paragraph(path, text="This report covers Q1 results.",
       style="Body Text", confirm=True)
   → {"paragraph_index": 43, "style": "Body Text", "text": "..."}

3. add_table(path, rows=3, cols=4,
       data=[["Item","Q1","Q2","Q3"],["A","1","2","3"],["B","4","5","6"]],
       confirm=True)
   → {"table_index": 2, "rows": 3, "cols": 4}

4. add_page_break(path, confirm=True)
   → {"ok": true}
```

### Recipe 3 — Find and Replace

```
1. find_replace(path, find_text="DRAFT", replace_text="FINAL",
       paragraph_index=12, occurrence=2, confirm=True)
   → {"replacements_made": 1, "find_len": 5, "replace_len": 5}
```

### Recipe 4 — Insert an Image

```
1. insert_image(path, image_path="C:\\Temp\\chart.png",
       width_inches=5.0, confirm=True)
   → {"ok": true, "width_inches": 5.0}
```

---

## Test Architecture

Current pytest inventory in `wordmcp/tests/`:

- **456** non-integration tests
- Markers in active use: `unit`, `smoke`, `security`, `integration`

---

## Architecture

- **`src/wordmcp/document_docx.py`** — public python-docx facade; re-exports the Phase 1 backend implemented under the private `src/wordmcp/_docx/` package; no FastMCP imports
- **`src/wordmcp/document_com.py`** — backward-compatibility shim for the COM backend (Phase 2); re-exports the implementation from the private `src/wordmcp/_com/` sub-package (`_base.py`, `_read.py`, `_write.py`, `_export.py`); Windows-only; no FastMCP imports
- **`src/wordmcp/_docx/context.py`** — summary-only context extraction for metadata subset, aggregate counts, and heading text
- **`src/wordmcp/server.py`** — thin FastMCP bootstrap/composition entrypoint; configures runtime, binds tool callables, and registers handler modules; no business logic
- **`src/wordmcp/_server_handlers_docx_core.py`** — core docx read/write handlers (including `find_replace`)
- **`src/wordmcp/_server_handlers_context.py`** — read-only context tool registration for `get_document_context`
- **`src/wordmcp/_server_handlers_docx_edit.py`** / **`src/wordmcp/_server_handlers_docx_advanced.py`** — decomposed docx edit/review/content/navigation handlers
- **`src/wordmcp/_server_handlers_com.py`** — COM tool handlers; **`src/wordmcp/_server_handlers_meta.py`**, **`src/wordmcp/_server_runtime.py`** — capabilities/meta registration and backend dispatch/error translation; **`src/wordmcp/_manifest_enums.py`** — tool-category enums (operation sets, scope constants); **`src/wordmcp/_manifest_registry.py`** — `TOOL_REGISTRY` and write-gate metadata; **`src/wordmcp/_server_manifest.py`** — thin re-export shim (backwards-compatible public surface for callers importing from this module)

Layer separation is strictly enforced. Public MCP tool callables now live in the handler modules, while `server.py` remains the thin bootstrap that composes and registers them.

---

## Security

- **Allowlist:** All file paths are validated against `WORD_ALLOWLIST_ROOTS` before any read or write. Path traversal outside configured roots is blocked.
- **Null-byte guard:** `_check_path()` and `_check_image_path()` reject any path containing a null byte (`\0`) to prevent null-byte injection attacks.
- **Write gate:** All mutating tools require `WORD_ENABLE_WRITE=true` (env) **and** `confirm=True` (parameter). Omitting either returns an error before touching any file.
- **DoS bounds (`add_table`):** `rows` is capped at 500 and `cols` at 100. Requests exceeding these limits are rejected before any file I/O.
- **Atomic save:** Each write tool in Phase 1 saves to disk immediately — no partial-write state is possible.
- **COM gate:** COM tools check `WORD_ENABLE_COM=true` before dispatch. If `pywin32` is not installed, tools raise `ToolError` with an install hint rather than crashing.
- **COM resource safety:** Every COM `Document` is closed in a `finally` block to prevent orphaned Word processes on error.
- **UNC path rejection (W2-005):** `_check_path()`, `_check_image_path()`, and `_check_evidence_path()` in `_docx/core.py` all reject paths starting with `\\` (backslash UNC, e.g. `\\server\share`) or `//` (forward-slash UNC, e.g. `//server/share`) before `Path.resolve()` is called — preventing path-traversal via UNC normalisation. `document_com.py` retains an additional `\\`-prefix check as a defence-in-depth layer for COM calls.
- **FileError message sanitisation (W2-004):** `_server_runtime.py` translates `FileError` to `ToolError("Operation failed (FileError)")` — filesystem paths cannot leak through tool error responses.
- **Phase 1 tools COM-free:** python-docx tools work on Windows and Linux without a Word installation.
- **stdout discipline:** All logging goes to `stderr`. `stdout` carries MCP JSON-RPC traffic only.
- **Comment sanitization:** `manage_comments` author and text fields are stripped of XML control characters before lxml storage (prevents `ValueError` on write).
- **Bounded document cache:** The open-document cache is bounded (LRU, max `WORD_MAX_CACHE_DOCS` entries) — unbounded growth is not possible.
- **File size limit:** `.docx` files exceeding `WORD_MAX_FILE_MB` MB are rejected by `_check_path` before loading into memory.
- **COM error hygiene:** HRESULT tuple details are logged to `stderr` only; `ToolError` messages returned to the client are sanitized.
- **URL guard:** `add_hyperlink` rejects URLs containing null bytes or control characters in addition to scheme validation.

---

## Known Constraints (Phase 1)

- `find_replace` supports cross-run matching within a paragraph, but `paragraph_index` scoping applies to body paragraphs only.
- SVG images are unsupported (python-docx 0.8.11 limitation).
- No pagination for `list_paragraphs` or `read_table` — very large documents may produce oversized responses.
- Phase 1 auto-saves after every mutation; batched/transactional edits are planned for Phase 2.
- `find_replace` now returns `find_len`/`replace_len` (character counts) instead of echoing the raw search/replace strings — callers that relied on the old string fields must update.

---

## Changelog

### 2026-06-07 — Dispatcher Tools (doc sync)

Eight new dispatcher tools registered, bringing the total to **51**:

- `paragraph` — unified dispatcher for all paragraph operations (list, read, add, add_heading, insert, update, delete, set_format, bulk_add, bulk_update)
- `table` — unified dispatcher for all table operations (list, read, add, update_cell, bulk_update_cells)
- `style` — unified dispatcher for style operations (list, apply)
- `content` — unified dispatcher for content write operations (add_list, add_page_break, insert_image, add_hyperlink, add_footnote, find_replace, set_properties)
- `document` — unified dispatcher for document utility operations (save, search_text, get_headers_footers)
- `review` — unified dispatcher for review/evidence operations (review, write_findings, export_evidence, manage_comments, manage_tracked_changes)
- `export` — unified export dispatcher with write-gated scopes (pdf, txt, markdown)
- `read_document` — promoted from single-scope to unified read dispatcher with 7 scopes (full, metadata, headings, outline, section, paragraphs, context)

Also registered: `set_paragraph_format` (standalone, deprecated in favour of `paragraph(operation='set_format')`).

All individual tools replaced by dispatchers remain registered for backward compatibility.
Total tool count: **43 → 51**.

### 2026-05-17 — Document Assembly

Three new tools delivered in `_server_handlers_assembly.py` / `_docx/assembly.py`:

- `bulk_find_replace` — write-gated; replaces multiple `{{TOKEN}}` placeholders in one pass (`WORD_ENABLE_WRITE=true` + `confirm=True`)
- `manage_hyperlinks` — write-gated; updates hyperlink URLs/labels by display-text matching (`WORD_ENABLE_WRITE=true` + `confirm=True`)
- `verify_no_placeholders` — read-only scan for residual `{{TOKEN}}` placeholders in body paragraphs and table cells

New module: `_docx/assembly.py` (backend logic), `_server_handlers_assembly.py` (MCP registration).
No new environment variables — `WORD_ALLOWLIST_ROOTS` and `WORD_ENABLE_WRITE` reused.
Total tool count: **40 → 43**.

### v0.4.0 — 2026-03-12

- **Unified export surface:** `export_document(format=Literal["txt","md","pdf","html"])` replaces four deprecated tools.
  - `format="txt"` routes to python-docx path (was `export_as_text`)
  - `format="md"` routes to python-docx path (was `export_to_markdown`)
  - `format="pdf"` / `format="html"` route to Word COM path
- **New:** `bulk_update_paragraphs` — update up to 500 paragraphs in a single save
- **New:** `bulk_update_table_cells` — update up to 200 table cells in a single save
- **`manage_comments`:** added `resolve` and `delete` operations
- **`manage_tracked_changes`:** per-revision ops (`accept_one`, `reject_one`) removed from MCP surface (Phase 3 scope)
- **Removed (4 deprecated COM wrappers):** `export_as_pdf`, `accept_all_track_changes`, `reject_all_track_changes`, `list_tracked_changes`
- **Removed (absorbed into `export_document`):** `export_as_text`, `export_to_markdown`
- **Tool count:** 43 → 39 (budget ≤40 ✔)
- Stage 3 template applied to all tool descriptions; peer review: PASS

### 2026-03-11 — Review/Evidence Tools (T-WR1..T-WR4)

Four new tools delivered and peer-reviewed:

- `get_document_outline` — read-only structural outline extractor
- `review_document` — 4-check structural review runner (`style_consistency`, `heading_hierarchy`, `word_count`, `table_structure`)
- `write_review_findings` — write-gated `.jsonl` evidence writer (`WORD_ENABLE_WRITE=true` + `confirm=True`)
- `export_review_evidence` — read-only evidence bundle packager for AW sign-off

New modules: `_docx/review_doc.py` (backend logic), `_server_handlers_review.py` (MCP registration).
New MCP resource: `word://prompts/word_review_check_iterate_v1`.
No new environment variables — `WORD_ALLOWLIST_ROOTS` reused.
Total tool count: **43**.

### v0.3.0 — 2026-03-05
- **Version bump:** 0.1.0 → 0.3.0
- 9 security bugs fixed (H-001 to H-009): LRU cache, file-size limit,
  XML control-char sanitization, URL null-byte guard, COM error hygiene,
  gate order consistency, Popen error wrapping, PII echo removal
- **`capabilities()` now exposes `"version"` key** — reflects installed package version
  (`wordmcp.__version__`); return shape: `{phase, backend, version, python_docx_version, tools, governance, com_tools}`
- New env vars: `WORD_MAX_FILE_MB` (default 50), `WORD_MAX_CACHE_DOCS` (default 20)
- New test file: `test_security_guards.py` (18 security tests)
- Historical test count at sprint close: 346
