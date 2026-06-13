# pptmcp — PowerPoint MCP Server v0.5.0

`pptmcp` is a Python MCP stdio server for reading and editing `.pptx` files using
[python-pptx](https://python-pptx.readthedocs.io/). It is part of the
[mcp-office](https://github.com/dosev-ai/mcp-office) suite. Phase 1 operates entirely at the file level —
no PowerPoint installation required. Phase 2 adds six Windows-only COM tools
(`export_slide`, `export_deck`, `run_slide_show`, `recalculate_charts`,
plus deprecated wrappers `export_slide_as_png` and `export_deck_as_pdf`)
via `pywin32`. Write operations are gated behind `PPT_ENABLE_WRITE=true` (environment
variable) and a `confirm=True` parameter for write and other mutation operations, preventing accidental
mutations. COM tools additionally require `PPT_ENABLE_COM=true`.

---

## Server Architecture

`server.py` (77 L) is a thin aggregator. On startup it calls four `register_X_tools(mcp)` functions, one per handler module:

| Module | Role | Tools |
|---|---|---|
| `_server_handlers_read.py` (227 L) | Read tools | 15 |
| `_server_handlers_write.py` | Write tools incl. slide-rebuild, table styling, and dispatch | 13 |
| `_server_handlers_review.py` (85 L) | Review / Output-Contract tools | 4 |
| `_server_handlers_com.py` (183 L) | COM tools (Windows-conditional) | COM-conditional |

Supporting modules: `presentation_pptx.py`, `content_pptx.py`,
`_pptx_caps.py` (capabilities + param schema), `presentation_com.py`,
`shapes_pptx.py` (re-export shim — implementation in `_shapes_helpers.py`, `_shapes_text.py`, `_shapes_geometry.py`, `_shapes_links.py`),
`contract_pptx.py`,
`review_pptx.py` (re-export shim — implementation in `_review_detectors.py`, `_review_export.py`),
`server_prompts.py`.

---

## Quick Start

```powershell
# Install local workspace dependency first (required — not on PyPI)
pip install -e "./shared"

# Install (editable, with dev deps — file-backend only)
pip install -e "./pptmcp[dev]"

# Install with Windows COM support (pywin32)
pip install -e "./pptmcp[com,dev]"

# Run tests (unit + smoke + security, no PowerPoint required)
pytest pptmcp/tests/test_unit.py pptmcp/tests/test_server_smoke.py pptmcp/tests/test_security.py -v

# Run integration tests (real .pptx file I/O)
pytest pptmcp/tests/test_integration.py -v -m integration
```

> **Current status:** v0.5.0 plus the 2026-05-08 table-authoring gap closure, with **48 always-registered tools** including the consolidated `slide`, `shape`, and `export` dispatch surfaces plus `set_table_style`, `manage_comments`, and `batch_set_text`. Historical UAT milestones, prior tool counts, and older test-count snapshots are preserved in the changelog sections below.

### VS Code mcp.json Registration

Add the following entry to
`%APPDATA%\Code - Insiders\User\mcp.json` (or `Code\User\mcp.json` for stable VS Code):

```json
"powerpoint-pptmcp": {
    "type": "stdio",
    "command": "python",
    "args": ["-m", "pptmcp.server"],
    "env": {
        "PYTHONPATH": "<path-to-pptmcp>/src",
        "PPT_ALLOWLIST_ROOTS": "C:\\Users\\yourname\\Documents,C:\\Temp",
        "PPT_ENABLE_WRITE": "true",
        "PPT_ENABLE_COM": "true"
    }
}
```

---

## Environment Variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `PPT_ALLOWLIST_ROOTS` | **yes** | — | Comma-separated absolute folder paths the server may access. Any path outside these roots is blocked with a `"Path not in allowlist"` error. Also used to validate COM export output paths (PNG/PDF). |
| `PPT_ENABLE_WRITE` | no | `false` | Must be `true` to enable any mutating tool. If unset or `false`, all write calls return an error even when `confirm=True` is passed. |
| `PPT_ENABLE_COM` | no | `false` | Set to `true` to enable the 6 Windows COM tools. `export_slide` and `export_deck` are always registered (return a graceful error when COM is unavailable); deprecated aliases and `run_slide_show`/`recalculate_charts` are COM-conditional. Requires `pywin32` (`pip install -e "./pptmcp[com,dev]"`). Has no effect on non-Windows platforms. |
| `PPT_COMPACT_TOOL_SURFACE` | no | `false` | Set to `true` to register the compact 21-tool agent-facing surface. This exposes dispatcher tools (`slide`, `shape`, `add_content`, `set_format`, `export`) plus read/review/QA tools while keeping the full 45-tool surface as the default for backward compatibility. |
| `PPT_MAX_FILE_MB` | no | `256` | Maximum `.pptx` file size in MB. Files larger than this limit are rejected at load time before any parsing. |
| `PPT_CACHE_MAX` | no | `20` | Maximum presentations held in the in-memory cache. Oldest entry is evicted when the limit is reached. |
| `PPT_LOAD_TIMEOUT_SECONDS` | no | `30` | Maximum seconds to wait when loading a `.pptx` file into the cache. |
| `PPT_MAX_COMMENTS` | no | `200` | Maximum number of comments returned by `manage_comments` `list` operation per request. |

---

## Tools Reference

### Read Tools (13)

No write-gate required. Safe to call without `PPT_ENABLE_WRITE`.

| Tool | Description | Key response fields |
|---|---|---|
| `capabilities` | Returns server phase, backend, per-tool parameter schema, and governance summary. `tools` and `com_tools` entries use dict format: `{"tool": str, "params": [{"name": str, "type": str, "required": bool}]}`. 48 always-registered tools; 4 platform-conditional COM-only tools in `com_tools`. | `phase`, `tools[].tool`, `tools[].params[]`, `com_tools[].tool`, `com_tools[].params[]`, `total_tools`, `governance` |
| `read_presentation` | Overview of all slides — count, titles, shape counts, notes flag | `slide_count`, `slides[].slide_index`, `slides[].title`, `slides[].shapes_count`, `slides[].has_notes` |
| `get_presentation_metadata` | File-level document properties | `title`, `author`, `subject`, `keywords`, `slide_count`, `created`, `modified` |
| `list_layouts` | Return all slide layouts in a presentation with index, name, and placeholder info. Use with `add_slide(layout_index=N)` to pick the right layout | `index`, `name`, `placeholder_count`, `placeholder_types` |
| `list_slides` | Ordered list of every slide with layout name | `result[].slide_index`, `result[].title` (`null` when not set), `result[].layout`, `result[].has_notes` |
| `read_slide` | Full detail of one slide — shapes, text, notes | `shapes[].shape_id`, `shapes[].name`, `shapes[].shape_type`, `shapes[].has_text`, `shapes[].text`, `shapes[].placeholder_idx`, `notes_text` |
| `list_shapes` | All shapes on a slide with geometry | `result[].shape_id`, `result[].name`, `result[].left`, `result[].top`, `result[].width`, `result[].height`, `result[].has_text_frame`, `result[].placeholder_idx`, `result[].bounds` (`left_in`, `top_in`, `width_in`, `height_in` in decimal inches; `null` if undetermined) |
| `get_shape` | Single shape by `shape_id` — includes text content | `shape_id`, `name`, `text`, geometry keys |
| `extract_tables` | All table data from one slide or the whole deck | `tables[].slide_index`, `tables[].rows[]` |
| `extract_images` | Metadata for every embedded image in the deck | `images[].slide_index`, `images[].shape_id`, `images[].name`, `images[].content_type`, `images[].width_px`, `images[].height_px` |
| `read_speaker_notes` | Speaker notes for one slide or all slides | `result[].slide_index`, `result[].notes_text` (`null` when slide has no notes) |
| `export_slide_as_text` | Plain-text dump of one or all slides. Table cells are included row-by-row as tab-delimited text. | `result[].slide_index`, `result[].texts[]` |
| `export` | **[Dispatch]** Unified export: `scope=slide_text` (text dump), `slide_images` (image metadata), `slide_png` (legacy PNG via COM), `deck_pdf` (PDF via COM), and `slide` (consolidated PNG slide render). `scope="slide"` supports `return_inline=True`, `response_mode`, `max_inline_bytes`, and allowlisted `options={"slide_index": 0, "inline_dpi": 96}` via the shared `mcpshared` inline artifact response helper. | list or dict per scope |
| `detect_overlapping_shapes` | Return all overlapping shape pairs on a slide with overlap area in square inches. Read-only — no write gate or `confirm` required. Shapes with undetermined bounds are skipped. | `slide_index`, `overlapping_pairs[].shape_a` (`shape_id`, `name`), `overlapping_pairs[].shape_b` (`shape_id`, `name`), `overlapping_pairs[].overlap_area_in2`, `pair_count` |

### Compact Agent Surface

Set `PPT_COMPACT_TOOL_SURFACE=true` to expose a dispatcher-first surface for agent deployments. Narrow legacy tools stay implemented and remain available in the default registration, but compact mode steers agents to `slide`, `shape`, `add_content`, `set_format`, and `export`.

`capabilities()` reports the compact contract under `compact_tool_surface` and recommended call groups under `tool_bundles` (`inspect`, `author`, `table_qa`, `render_check_iterate`).

### Inline Slide Render Response

Use the consolidated export dispatcher for visual feedback loops:

```python
export(
    scope="slide",
    path="deck.pptx",
    fmt="png",
    output_path="slide-001.png",
    return_inline=True,
    response_mode="auto",
    max_inline_bytes=3_000_000,
    options={"slide_index": 0, "inline_dpi": 96},
    confirm=True,
)
```

Inline payloads are opt-in. Oversize payloads do not fail a successful file export; the response includes `artifact.inline=false` with an explicit omission reason.

### Write Tools (12)

All require `PPT_ENABLE_WRITE=true`. Tools marked **confirm required** also require
`confirm=True` in the call arguments — omitting it returns an error without making any change.

> **Important:** Changes are **not auto-saved**. All write tools update in-memory state
> only. You **must** call `save` (with `confirm=True`) to persist changes to disk. Without
> saving, all changes are lost when the server process exits.

| Tool | confirm=True required? | Description | Key response fields |
|---|---|---|---|
| `edit_text_placeholder` | **yes** | Replace text in a placeholder by `placeholder_idx`, or in a text-containing shape by `shape_id` | `slide_index`, `placeholder_idx` or `shape_id`, `text` |
| `set_speaker_notes` | **yes** | Write speaker notes for a slide (held in memory until `save`) | `slide_index`, `notes_text` |
| `insert_image` | **yes** | Insert a local image file onto a slide (path must be in allowlist) | `slide_index`, `shape_name` |
| `add_slide` | **yes** | Append a new slide by layout index (0–10) with optional title. Optional `suppress_content_placeholder=True` removes empty non-title content placeholders from the new slide so authored content does not overlap them. | `slide_index`, `title`, `placeholders_removed` |
| `replace_slide_text` | **yes** | Find-and-replace text across all slides | `replacements_made` |
| `reorder_slides` | **yes** | Reorder slides by providing a complete new-index-order array | `new_order` |
| `delete_slide` | **yes** | Remove a slide by index | `deleted_index` |
| `delete_shape` | **yes** | Remove a shape from a slide by `shape_id` | `deleted` (bool), `shape_id`, `slide_index` |
| `save` | **yes** | Persist all in-memory changes to disk | `saved`, `path` |
| `create_presentation` | **yes** | Create a brand-new .pptx file at `path`. Optionally accepts `template_path` (existing PPTX used as template — validated via allowlist, template slides removed, title slide added). Optionally set a title slide. Writes directly to disk — no `save` call needed. Fails if file already exists | `path`, `slide_count`, `title` |
| `slide` | **yes** | **[Dispatch]** Unified slide mutation: `operation=add` (append slide), `delete` (by slide_index), `reorder` (full permutation via new_order), `copy` (within or across files). For `operation=add`, `suppress_content_placeholder` defaults to `True` and removes empty non-title content placeholders. Consolidates `add_slide`, `delete_slide`, `reorder_slides`, `copy_slide`. | `slide_index`, `new_order`, `placeholders_removed`, etc. per operation |
| `shape` | **yes** | **[Dispatch]** Unified shape mutation: `add_text_box`, `add_autoshape`, `add_table`, `delete` (by shape_id), `set_properties` (text/paragraph/fill formatting), `set_table_style` (table formatting). `set_properties` supports `target="fill"` with `fill_color_hex`, and fill colour is readable via `get_shape` / `list_shapes` as `fill_color_hex`. Consolidates common shape tools. | `shape_id`, `slide_index`, `fill_color_hex`, etc. per operation |

### Phase 2.1 Tools (2)

Two new tools added in v0.2.1. `extract_presentation_text` is read-only; `manage_hyperlinks`
is mixed-mode.

#### `extract_presentation_text` — READ-ONLY

No write gate required. Safe to call without `PPT_ENABLE_WRITE`.

| Tool | Description | Key response fields |
|---|---|---|
| `extract_presentation_text` | Full-deck structured text export for RAG indexing. Returns title, body text, notes, and per-shape text for every slide in one structured response. More structured than `export_slide_as_text`: title extracted separately, notes included, per-shape metadata. Table cells are included row-by-row in `body_text` and table shapes include `rows`. | `slide_count`, `slides[].slide_index`, `slides[].title`, `slides[].body_text`, `slides[].notes`, `slides[].shapes[].shape_id`, `slides[].shapes[].name`, `slides[].shapes[].text`, `slides[].shapes[].rows` |

#### `manage_hyperlinks` — MIXED READ/WRITE

`list` is read-only. `add` and `remove` require `PPT_ENABLE_WRITE=true` + `confirm=True`.

| operation | confirm=True required? | Description | Key response fields |
|---|---|---|---|
| `list` | no | List all hyperlinks on shapes in a slide | `links[].shape_id`, `links[].url`, `links[].tooltip` |
| `add` | **yes** (+ `PPT_ENABLE_WRITE=true`) | Add a hyperlink to a shape's text run | `ok: true`, `runs_updated: N` |
| `remove` | **yes** (+ `PPT_ENABLE_WRITE=true`) | Clear all hyperlinks from a shape | `ok: true`, `runs_cleared: N` |

> **Security:** `target_url` is validated against an allowlist of URL schemes: `https`, `http`,
> `mailto` only. Any other scheme (e.g. `javascript:`, `file:`, `ftp:`) is rejected before
> any mutation occurs.

#### `manage_comments` — MIXED READ/WRITE

`list` is read-only. `add` and `delete` require `PPT_ENABLE_WRITE=true` + `confirm=True`.

| Tool | Description |
|---|---|
| `manage_comments` | List, add, or delete comments on a PowerPoint slide (`operation`: `list` / `add` / `delete`). `list` is read-only; `add` and `delete` require `PPT_ENABLE_WRITE=true` and `confirm=True`. |

| operation | confirm=True required? | Description | Key response fields |
|---|---|---|---|
| `list` | no | List all comments on `slide_index` (or all slides if omitted) | `comments`, `total_found`, `truncated`, `max_comments` |
| `add` | **yes** (+ `PPT_ENABLE_WRITE=true`) | Add a comment. Requires `slide_index` and `text`. Optional `author` (default `"MCP"`), `slide_x_emu`, `slide_y_emu` | `status`, `comment_id`, `author`, `slide_index` |
| `delete` | **yes** (+ `PPT_ENABLE_WRITE=true`) | Delete comment by `comment_id`. Requires `slide_index` and `comment_id` | `status`, `comment_id`, `slide_index` |

#### `batch_set_text` — WRITE-GATED

Requires `PPT_ENABLE_WRITE=true` + `confirm=True`.

| Tool | Description | Parameters | Key response fields |
|---|---|---|---|
| `batch_set_text` | Update text on multiple shapes in a single slide in one open/save cycle. Each entry in `updates` must be `{"shape_id": int, "text": str}`. Shapes that are missing or have no text frame are collected in `skipped` rather than aborting the batch. Requires `PPT_ENABLE_WRITE=true` and `confirm=True`. | `path` (str), `slide_index` (int), `updates` (list[dict]), `confirm` (bool) | `updated`, `path`, `slide_index`, `results[].shape_id`, `results[].status`, `results[].previous_text`, `results[].new_text`, `skipped[].shape_id`, `skipped[].reason`, `error_count` |

### Phase 1.5 Write Tools + Table Styling (8)

All require `PPT_ENABLE_WRITE=true` + `confirm=True`.

| Tool | confirm=True required? | Description | Key parameters |
|---|---|---|---|
| `add_textbox` | **yes** | Add a text box to a slide. Dimensions in inches. Text is word-wrapped. Response includes `overflow_risk` (`low`/`medium`/`high`), `overflow_detail` (heuristic estimate), `overlap_warning` (`null` or `{shape_a, shape_b, overlap_area_in2}` if the new shape overlaps an existing one) | `slide_index`, `left`, `top`, `width`, `height`, `text` |
| `add_shape` | **yes** | Add an autoshape to a slide. `shape_type`: `RECTANGLE`, `OVAL`, `ROUNDED_RECTANGLE`, etc. Dimensions in inches. Text is word-wrapped. Response includes `overflow_risk` (`low`/`medium`/`high`), `overflow_detail` (heuristic estimate), `overlap_warning` (`null` or `{shape_a, shape_b, overlap_area_in2}`) | `slide_index`, `shape_type`, `left`, `top`, `width`, `height`, `text` |
| `add_table_to_slide` | **yes** | Add a table to a slide. `data` is an optional 2-D list of strings. Dimensions in inches. Optional compact-table controls: `font_size_pt`, `header_font_size_pt`, `header_bold`, `row_height_pt`, and `suppress_content_placeholder` (default `True`, removes only empty overlapping non-title content placeholders). | `slide_index`, `rows`, `cols`, `left`, `top`, `width`, `height`, `data`, `font_size_pt`, `header_font_size_pt`, `header_bold`, `row_height_pt`, `suppress_content_placeholder` |
| `set_table_style` | **yes** | Apply table formatting to an existing table shape. Supports body/header font size, bold, font colour, fills, header formatting, and row heights. | `slide_index`, `shape_id`, `font_size_pt`, `bold`, `color_hex`, `header_font_size_pt`, `header_bold`, `header_color_hex`, `fill_color_hex`, `header_fill_color_hex`, `row_height_pt`, `header_row_height_pt` |
| `set_text_format` | **yes** | Apply text formatting to all runs in a shape's text frame. `color_hex`: 6-char hex e.g. `FF0000`. All format params are optional | `slide_index`, `shape_id`, `bold`, `italic`, `font_size_pt`, `font_name`, `color_hex` |
| `set_paragraph_format` | **yes** | Set paragraph alignment and spacing for one paragraph in a shape. `alignment`: `CENTER`, `LEFT`, `RIGHT`, `JUSTIFY`, `DISTRIBUTE` | `slide_index`, `shape_id`, `paragraph_index`, `alignment`, `line_spacing`, `space_before_pt`, `space_after_pt` |
| `copy_slide` | **yes** | Copy a slide within or across presentations. `source_path == target_path` copies within the same file. `target_slide_index` optional (default: append at end) | `source_path`, `source_slide_index`, `target_path`, `target_slide_index` |
| `add_hyperlink` | **yes** | Apply a URL hyperlink (https://, http://, mailto: only) to a text run in a shape. | `slide_index`, `shape_id`, `run_index`, `url`, `display_text` |

### Phase 2 COM Tools (6) — Windows only

Require **`PPT_ENABLE_COM=true`** (env var), `pywin32` installed, and PowerPoint desktop
application present. COM tools that write output also require `PPT_ENABLE_WRITE=true` +
`confirm=True`. Output paths are validated against `PPT_ALLOWLIST_ROOTS`.

| Tool | confirm=True required? | Description |
|---|---|---|
| `export_slide` | **yes** | Export a single slide as PNG, PDF, or SVG. `fmt`: `'png'`, `'pdf'`, or `'svg'`; `slide_number` is 1-based. `dpi` clamped 72–600 for raster formats. Requires `PPT_ENABLE_COM=true`, Windows only. |
| `export_deck` | **yes** | Export an entire presentation as PDF, PPTX, or a folder of PNG images. `fmt`: `'pdf'`, `'pptx'`, or `'images'`. For `'images'`, `output_path` is a directory. Requires `PPT_ENABLE_COM=true`, Windows only. |
| `run_slide_show` | **yes** (+ `PPT_ENABLE_WRITE=true`) | Open the presentation in slide-show mode in PowerPoint (requires PPT_ENABLE_COM=true, PPT_ENABLE_WRITE=true, Windows only). |
| `recalculate_charts` | **yes** | Activate embedded charts for recalculation only. Linked and external refresh is skipped for safety. Requires `PPT_ENABLE_COM=true`, `PPT_ENABLE_WRITE=true`, and `confirm=True` (Windows only). |
| `export_slide_as_png` | **yes** | **(Deprecated — use `export_slide(fmt='png')` instead. Removed in v0.6.0.)** Export a single slide as a PNG image file. Optional `width_px` / `height_px` parameters (0–7680 px each; Phase C). |
| `export_deck_as_pdf` | **yes** | **(Deprecated — use `export_deck(fmt='pdf')` instead. Removed in v0.6.0.)** Export the entire presentation as a PDF file. Response includes `page_count` key (Phase C). |

### Phase A Output Contract Tools (3)

Tools for Output Contract authoring and validation. `validate_contract` and `check_presentation_against_contract` are read-only. `declare_slide_contract` requires `PPT_ENABLE_WRITE=true` + `confirm=True`.

| Tool | confirm=True required? | Description | Key response fields |
|---|---|---|---|
| `validate_contract` | no | Validate a JSON Output Contract file against the pptmcp v1.0 schema. Enforces required fields, `contract_version` enumeration (`"1.0"` only), type constraints (no booleans in numeric fields), and a 512 KB size limit. Path must be in `PPT_ALLOWLIST_ROOTS`. | `valid` (bool), `contract_version` (str or null), `slide_count` (int), `errors` (list[str]) |
| `check_presentation_against_contract` | no | Check a `.pptx` file against its Output Contract specification. Returns per-slide findings for criterion violations (slide count, shape count bounds, slide dimensions). | `passed` (bool), `slide_count` (int), `findings[].slide_index`, `findings[].criterion`, `findings[].detail` |
| `declare_slide_contract` | **yes** (+ `PPT_ENABLE_WRITE=true`) | Persist a slide-export contract beside the PPTX (RCI-US-01). Writes `contract.json` in the same directory as the PPTX. The spec is validated against the v1.0 schema before the file is committed; failures leave no partial file on disk. Path must be in `PPT_ALLOWLIST_ROOTS`. | `contract_path`, `checksum_sha256`, `contract_version` |

> **Error handling:** `validate_contract` raises `ToolError` for I/O failures (file not found, path not in allowlist, bad extension, file > 512 KB, malformed JSON, missing `contract_version`, boolean in a numeric field). Schema mismatches return `valid=False` with a populated `errors` list.

### Phase C Export Enhancement Tools

`export_slides_to_stamped_dir` requires `PPT_ENABLE_COM=true` + `PPT_ENABLE_WRITE=true` + `confirm=True`. Output paths validated against `PPT_ALLOWLIST_ROOTS`.

| Tool | confirm=True required? | Description | Key response fields |
|---|---|---|---|
| `export_slides_to_stamped_dir` | **yes** | Export all slides to a timestamped run directory as PNG images. Files named `slide-000.png`, `slide-001.png`, … (0-based). Optional `width_px` / `height_px` (0–7680 px each, validated). Requires `PPT_ENABLE_COM=true` + `PPT_ENABLE_WRITE=true`. Path validated against `PPT_ALLOWLIST_ROOTS`. | `run_dir`, `exported_slides`, `slide_count` |

**Phase C enhancements to existing Phase 2 COM tools:**

| Tool | Enhancement |
|---|---|
| `export_slide_as_png` | Added optional `width_px` / `height_px` parameters (0–7680 px each, validated). |
| `export_deck_as_pdf` | Response dict now includes `page_count` key. |

### Phase D Review Gate Tools (1)

Write-gated tool for slide export quality review. Requires `PPT_ENABLE_WRITE=true` + `confirm=True` (the gate covers writing the `review_result.json` audit artefact beside the PNG).

| Tool | confirm=True required? | Description | Key response fields |
|---|---|---|---|
| `review_slide_export` | **yes** (+ `PPT_ENABLE_WRITE=true`) | Review a slide export against **6 structural criteria**: `TEXT_OVERFLOW` (chars/sq.in heuristic), `SHAPE_OUTSIDE_SLIDE` (bounding box vs slide dimensions), `TABLE_CONTENT_CLIPPING` (table row/column count vs visible cell area), `UNFINISHED_PLACEHOLDER` (placeholder text left unfilled), `LOW_FONT_SIZE` (font size below readability threshold), and `SPARSE_SLIDE` (slide has too little content). `png_path` must exist and be in allowlist. `contract_path` is optional. Writes `review_result.json` beside the PNG as a durable audit artefact. Severity values: `high`/`medium`/`low`. **`passed=True` reflects structural checks only — not delivery-readiness.** Advisory visual QA findings are returned separately in `visual_qa_findings` and do not affect `passed`. | `passed` (bool — structural checks only), `findings[].criterion`, `findings[].result`, `findings[].detail`, `findings[].severity`, `checks_run` (list[str] — 6 checks), `findings_count` (int), `visual_qa_findings[].severity`, `visual_qa_findings[].check`, `visual_qa_findings[].detail`, `visual_readiness` (always `"not_assessed"`), `confidence_boundary` (always `"structural_heuristics_only"`), `review_result_path` (str) |

### Phase E Iterate + Evidence Bundle Tools (2)

`export_changed_slides_only` requires `PPT_ENABLE_COM=true` + `PPT_ENABLE_WRITE=true` + `confirm=True`. `produce_evidence_bundle` is read-only.

| Tool | confirm=True required? | Description | Key response fields |
|---|---|---|---|
| `produce_evidence_bundle` | no | Aggregate `review_slide_export` results into a machine-verifiable evidence bundle. `review_results`: list of dicts from `review_slide_export()`. Read-only — no write gate required. Path must be in `PPT_ALLOWLIST_ROOTS`. | `bundle_version`, `overall_passed` (bool), `slides[]` (per-slide summary), `total_findings` (int) |
| `export_changed_slides_only` | **yes** | Re-export only slides whose content changed since a `previous_hashes` snapshot. `previous_hashes`: dict `{str(slide_index): md5_hex_32_chars}`. Requires `PPT_ENABLE_COM=true` + `PPT_ENABLE_WRITE=true`. Previous-hash values validated as 32-char MD5 hex before use. | `run_dir`, `exported_slides` (list of PNG paths), `unchanged_slides` (list of indices), `new_hashes` (dict) |

### ACP Tool (1)

Read-only. No write gate required.

| Tool | Description | Key response fields |
|---|---|---|
| `get_presentation_context` | Return an Artifact Context Packet (ACP) for a PowerPoint presentation. Progressive disclosure via `level`: `"index"` (identity + summary only), `"focused"` (+ `slide_count`, `slide_titles`), `"deep"` (+ `png_paths`, `review_findings`, optional `annotations`). Path must be in `PPT_ALLOWLIST_ROOTS`. Annotations are validated before storage. | `artifact_id`, `summary`, `slide_count`, `slide_titles`, `png_paths`, `review_findings`, `annotations` |

### FR-826 Slide Rebuild Tools (3)

All require `PPT_ENABLE_WRITE=true` + `confirm=True`. Implemented in `_server_handlers_write.py`.

| Tool | confirm=True required? | Description | Key response fields |
|---|---|---|---|
| `clear_slide_content` | **yes** | Remove all shapes from a slide. | `slide_index`, `shapes_removed` (int) |
| `apply_slide_layout` | **yes** | Swap the layout of a slide via OPC relationship. Optional `remove_placeholders` (bool, default `False`): if `True`, removes any layout placeholders not already present on the slide. | `slide_index`, `layout_index` |
| `remove_empty_placeholders` | **yes** | Remove placeholders whose text is empty or contains only default placeholder text. | `slide_index`, `placeholders_removed` (int) |

---

## Workflow Prompt

### `ppt_render_check_iterate_v1`

> **Description:** Canonical agent workflow: Output Contract → Build → Export PNG/PDF → Self-Review → Iterate → Evidence Bundle. Use this prompt whenever building or editing a PowerPoint deck.

Accessible via MCP `prompts/list` and `prompts/get`. The prompt guides an agent through six mandatory steps:

| Step | Name | Summary |
|---|---|---|
| 1 | **Output Contract** | Produce `output_contract.json` before writing any slides. Defines every slide, shape, layout, export plan, acceptance criteria, and evidence paths. |
| 2 | **Build** | Implement all slides using pptmcp authoring tools (`create_presentation` → `add_slide` → `add_textbox` / `add_shape` / `add_table_to_slide` / `insert_image`). Save after every meaningful change. |
| 3 | **Export PNG Renders** | Export a PNG for every changed slide via `export_slide(fmt='png', dpi=150, ...)`. Requires `PPT_ENABLE_COM=true`. Naming convention: `slide-001.png`, `slide-002.png`, ... |
| 4 | **Milestone PDF Export** | Export the full deck via `export_deck(fmt='pdf', ...)` for end-to-end review. Requires `PPT_ENABLE_COM=true`. |
| 5 | **Self-Review Gate** | Check each PNG against `output_contract.json`: text overflow, shapes outside slide, layout drift, font/style, table sizing. Fix → re-export → repeat until all PASS (max 3 iterations). |
| 6 | **Evidence Bundle** | Produce `evidence.json` recording contract path, slide exports, PDF path, per-slide review results, final verdict, tool-call count, and iteration count. Attach to the Action Worker record. |

**Required environment gates for the full workflow:**
- `PPT_ENABLE_WRITE=true` — all build and export steps
- `PPT_ENABLE_COM=true` — Steps 3 and 4 (PNG/PDF export)
- `PPT_ALLOWLIST_ROOTS=<output_dir>` — your working directory
- `confirm=True` on every write/export call

The prompt body is also callable directly via the FastMCP `.fn()` pattern for unit testing.

### `capabilities()` — workflow discovery keys

The `capabilities` tool now returns two additional keys that help agents discover the recommended workflow:

| Key | Value |
|---|---|
| `recommended_workflow` | `"Output Contract → Build → Export (PNG/PDF) → Self-Review → Iterate → Evidence Bundle"` |
| `prompt_name` | `"ppt_render_check_iterate_v1"` |

---

## MCP Resources

| Resource URI | MIME type | Description |
|---|---|---|
| `resource://contract-schema` | `application/json` | JSON Schema v1.0 for pptmcp Output Contract files. Returns the static schema dict serialised as JSON. No I/O performed — schema is a module-level constant in `contract_pptx.py`. |
| `ppt://prompts/ppt_render_check_iterate_v1` | `text/markdown` | Canonical Render-Check-Iterate workflow prompt as a resource endpoint. Returns the six-step workflow (Declare Contract → Build → Export PNG → Review → Iterate → Evidence Bundle) as Markdown text. Registered via `@mcp.resource("ppt://prompts/ppt_render_check_iterate_v1")` in `server.py` (Phase E). |

---

## UX Patterns and Recipes

### Recipe 1 — Add a New Slide and Populate It

```
1. add_slide(path, layout_index=1, title="Q2 Summary", confirm=True)
   → {"slide_index": 4, "title": "Q2 Summary"}

2. list_shapes(path, slide_index=4)
   → shows placeholder_idx=0 (title) and placeholder_idx=1 (content body)

3. edit_text_placeholder(path, slide_index=4, placeholder_idx=1,
       text="• Revenue up 12%\n• NPS 62")
   (requires confirm=True)

4. set_speaker_notes(path, slide_index=4,
       notes_text="Presenter context here")
   (requires confirm=True)

5. save(path, confirm=True)
   → {"saved": true, "path": "..."}
```

### Recipe 2 — Inspect and Edit an Existing Presentation

```
1. read_presentation(path)
   → slide count, titles, shapes_count per slide

2. list_slides(path)
   → see layout names, which slides have notes, null-title slides

3. read_slide(path, slide_index=2)
   → all shapes with placeholder_idx, text content, and notes_text

4. edit_text_placeholder(path, slide_index=2, placeholder_idx=1,
       text="Updated content")

5. replace_slide_text(path, find="Draft", replace="Final", confirm=True)
   → {"replacements_made": 3}

6. save(path, confirm=True)
```

### Recipe 3 — Extract All Content for Analysis

```
1. export_slide_as_text(path)
   → all text by slide (omit slide_index to get all slides)

2. extract_tables(path)
   → all table data across the deck as row/cell arrays

3. extract_images(path)
   → image inventory: content_type, dimensions (metadata only, no binary)

4. read_speaker_notes(path)
   → all speaker notes; notes_text is null for slides with no notes
```

### Recipe 4 — Reorder and Clean Up a Deck

```
1. list_slides(path)
   → get current slide count and order

2. reorder_slides(path, new_order=[0, 3, 1, 2], confirm=True)
   → slides are rearranged; old indices are invalidated

3. delete_slide(path, slide_index=3, confirm=True)
   → removes the (now last) slide

4. save(path, confirm=True)
```

---

## Current Limitations and Not Yet Available Features

| Feature | Status | Notes |
|---|---|---|
| Creating a new `.pptx` from scratch | ✅ Delivered | `create_presentation` tool (Phase 1.5) |
| Slide templates / themes / master edits | ❌ Not available | python-pptx has partial support; not exposed |
| Copy / duplicate a slide | ✅ Delivered | `copy_slide` tool (Phase 1.5) |
| Animations & transitions | ❌ Not available | Not supported in python-pptx |
| Charts (create / edit data) | ❌ Not available | Read embedded table data via `extract_tables` only |
| SmartArt | ❌ Not available | python-pptx does not expose SmartArt |
| Embedded videos / audio | ❌ Not available | |
| Slide shows / presentation mode | ✅ Delivered | `run_slide_show` tool (Phase 2 COM, Windows only) |
| Export to PDF / image / SVG per slide | ✅ Delivered | `export_slide` (png/pdf/svg) + `export_deck` (pdf/pptx/images) (Phase 2 COM, Windows only) |
| Live COM recalculation | ✅ Delivered | `recalculate_charts` tool (Phase 2 COM, Windows only) |
| Multi-file / merge decks | ❌ Not available | Planned |
| Image extraction (binary / base64) | ❌ Not available | Metadata only via `extract_images` |
| Text formatting (bold, font size, colour) | ✅ Delivered | `set_text_format` + `set_paragraph_format` (Phase 1.5) |
| Slide layout preview / catalogue | ✅ Delivered | `list_layouts` tool (Phase 1.5) |

---

## Layout Index Reference

Layout indices depend on the slide master embedded in the `.pptx` file. Use `list_slides`
to see which layout names are already in use. Common indices for default Office themes:

| Index | Typical name |
|---|---|
| 0 | Title Slide |
| 1 | Title and Content |
| 2 | Title and Two Content |
| 3 | Title Only |
| 4 | Blank |
| 5–10 | Varies by theme |

---

## Security & Governance

- All file paths are validated against `PPT_ALLOWLIST_ROOTS` before any I/O. Traversal
  above an allowed root is blocked with `"Path not in allowlist"`.
- Write and other mutation operations require `PPT_ENABLE_WRITE=true` AND `confirm=True`.
  The two-factor pattern prevents accidental mutation from misconfigured clients.
- `stdout` carries MCP JSON-RPC traffic only. All server logs go to `stderr`.
- No network calls. All operations are local file I/O only.
- **File-size limit:** `.pptx` files exceeding `PPT_MAX_FILE_MB` (default 256 MB) are rejected
  at load time before any parsing occurs.
- **URL scheme validation:** `manage_hyperlinks` and `add_hyperlink` use `urlparse` to validate
  URL schemes; only `https`, `http`, and `mailto` are accepted.
- **HRESULT sanitisation:** COM error messages have HRESULT codes stripped before being forwarded
  to MCP clients — raw Win32 error details are never exposed.
- **COM error wrapping:** all COM-calling tool wrappers in `_server_handlers_com.py` convert `PPTMCPError`
  to `ToolError`, preventing raw exceptions from reaching the MCP protocol layer.
- **ACP adapter injection hardening (OWASP A03):** `_pptx_acp_adapter.py` sanitizes `slide_titles` and `filename_clean` via `_sanitize_text_field` (imported from `mcpshared`) before embedding them in ACP annotations; `artifact_id` and summary fields use `filename_clean` rather than raw user input. `png_paths` are validated against `PPT_ALLOWLIST_ROOTS` to prevent symlink escape.

---

## Backlog

### Phase 2 — COM / win32com (Windows-only) — DELIVERED in v0.4.0

- [x] ~~Export slide as PNG / JPG via PowerPoint COM~~ → `export_slide_as_png`
- [x] ~~Export full deck to PDF via COM~~ → `export_deck_as_pdf`
- [x] ~~Trigger slide show / presentation mode~~ → `run_slide_show`
- [x] ~~Recalculate embedded chart data via COM~~ → `recalculate_charts`
- [ ] Live theme / master slide edits

### Phase 1 Extensions

- [x] ~~Create new `.pptx` from a blank template~~ → delivered as `create_presentation` (Phase 1.5 Batch 1)
- [ ] Copy / duplicate a slide within a deck
- [ ] Merge two decks (append slides from source to target)
- [ ] Set text formatting (bold, font size, colour) via python-pptx
- [ ] Extract image binaries (base64) from embedded images
- [x] ~~Layout catalogue tool (list all layout names + indices in a file)~~ → delivered as `list_layouts` (Phase 1.5 Batch 1)
- [ ] `export_slide_as_text` variant with per-shape detail

### Quality & Governance

- [ ] Bump to `python-pptx>=1.0` when stable
- [ ] Add `test_server_smoke.py` coverage for all 9 write tools
- [ ] UAT script for full 23-scenario suite in CI (non-COM, fixture-based)
