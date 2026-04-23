# Excel MCP — Render-Check-Iterate Workflow

> Version: excelmcp Sprint P
> Applies to: `export_as_pdf(scope='sheet' | 'workbook')`

## Purpose

This workflow produces PDF render artefacts from Excel through COM so an agent or human
reviewer can inspect the rendered output. The truth boundary is strict:

- `export_as_pdf` generates PDFs and returns export evidence only.
- excelmcp does **not** perform automated visual scoring.
- Visual judgement remains a human review or external agent-review activity.

## Environment

Required:

- `EXCEL_ENABLE_COM=true`
- `EXCEL_ALLOWLIST_ROOTS` for the source workbook
- `EXCEL_EXPORT_ROOTS` for the PDF destination

Notes:

- `EXCEL_ENABLE_WRITE` is **not** required for `export_as_pdf`.
- The PDF output parent directory must already exist.
- `export_as_pdf` rejects a missing parent directory or a non-directory parent.

## Unified export tool

Use the single MCP entrypoint.

### Sheet export

`export_as_pdf(scope='sheet', path=..., sheet=..., output_path=..., quality=0, confirm=True)`

Returns:

- `ok`
- `sheet`
- `output_path`
- `quality`
- `sha256`
- `size_bytes`
- `exported_at`
- `elapsed_ms`

### Workbook export

`export_as_pdf(scope='workbook', path=..., output_path=..., quality=0, confirm=True)`

For workbook scope, callers must omit `sheet`. A stray `sheet` argument is an invalid request.

Returns:

- `ok`
- `output_path`
- `quality`
- `source`
- `sha256`
- `size_bytes`
- `exported_at`
- `elapsed_ms`

`ok=True` means the PDF was created, is a file, and is non-zero in size.

Invalid export requests are rejected before any existing PDF target is cleared.
A pre-existing target is preserved on validation failure, including a stray `sheet` argument on workbook scope.

## Recommended loop

1. Apply workbook edits.
2. If formula freshness is in doubt, run `recalculate_workbook(..., confirm=True)`.
3. Export the changed sheet with `export_as_pdf(scope='sheet', ...)`.
4. Review the PDF outside excelmcp.
5. Repeat until the rendered output is acceptable.
6. Produce a milestone export with `export_as_pdf(scope='workbook', ...)`.

## Output handling

- Prefer stable export folders such as `C:/Temp/exports/...`.
- Create directories before calling the export tool.
- Keep each iteration in a separate folder if traceability is required.

Example paths:

- `C:/Temp/exports/budget/run-001/summary.pdf`
- `C:/Temp/exports/budget/run-001/workbook-full.pdf`

## Evidence semantics

- `sha256` identifies the exact exported PDF bytes.
- `size_bytes` confirms the export is non-empty.
- `exported_at` records the UTC completion time.
- `elapsed_ms` records end-to-end export duration.

These fields are evidence for export completion only. They are not proof of visual
quality and they are not automated visual scoring.

## Example calls

Sheet:

`export_as_pdf(scope='sheet', path='C:/Temp/book.xlsx', sheet='Summary', output_path='C:/Temp/exports/book/run-001/summary.pdf', quality=0, confirm=True)`

Workbook:

`export_as_pdf(scope='workbook', path='C:/Temp/book.xlsx', output_path='C:/Temp/exports/book/run-001/workbook-full.pdf', quality=0, confirm=True)`

## Limitations

- COM and desktop Excel are required.
- Missing output directories are rejected.
- excelmcp does not parse the generated PDF.
- excelmcp performs no automated visual scoring; visual acceptance remains human review.

## Automated Review Gate

Before or after export, call `review_workbook_render` to flag structural issues that predict render problems:

| Criterion | Detection | Severity |
|-----------|-----------|----------|
| `PRINT_AREA_UNSET` | `ws.print_area` is None or empty — workbook will print the full used range, possibly bleeding data | medium |
| `SHEET_EMPTY` | Sheet has no data (max_row ≤ 1) — will produce a blank PDF page | low |
| `FORMULA_STALE` | Workbook has formulas and `recalculated=False` — exported values may be stale cache | medium |
| `PRINT_AREA_EXCEEDS_PAGE` | Print area column count > 10 — may produce wide multi-page output | low |
| `FORMULA_SCAN_FAILED` | Formula scan could not complete — formula staleness is unknown | low |

**Usage:**

```python
result = review_workbook_render(path, recalculated=True)
if not result["passed"]:
    # Inspect result["findings"] and fix before exporting
    for f in result["findings"]:
        if f["severity"] in ("high", "medium"):
            print(f["criterion"], f["detail"])
```

`passed` is `True` when no `medium` or `high` severity findings exist. Low-severity findings appear in `findings` but do not block `passed`.

## Evidence Bundle

After reviewing and exporting, close the render loop by calling `produce_export_evidence_bundle`:

```python
export_result = export_as_pdf(path, scope="workbook", confirm=True)
review_result = review_workbook_render(path, recalculated=True)
bundle = produce_export_evidence_bundle(
    path=path,
    export_results=[export_result],
    review_results=[review_result],
)
# bundle["overall_passed"] is True when all reviews passed and at least one export succeeded
# bundle["generated_at"] provides the ISO-8601 closure timestamp
```

The bundle aggregates both the export evidence (sha256, size_bytes, exported_at) and the review findings into a single structured dict suitable for storing as actionable evidence.

### Iteration with Changed-Sheets-Only Re-export

After editing sheets between review passes, use `export_changed_sheets_only` to re-export only mutated sheets instead of the full workbook:

```python
# First run — export all
first_run = export_changed_sheets_only(
    path=path,
    base_dir="/path/to/exports",
    previous_hashes={},   # empty on first run → exports everything
    confirm=True,
)
# Store hashes for next iteration
my_hashes = first_run["new_hashes"]

# After further edits, re-export only what changed
second_run = export_changed_sheets_only(
    path=path,
    base_dir="/path/to/exports",
    previous_hashes=my_hashes,
    confirm=True,
)
print("Re-exported:", second_run["exported_sheets"])
print("Unchanged:", second_run["unchanged_sheets"])
```

Requires `EXCEL_ENABLE_WRITE=true` and `EXCEL_ENABLE_COM=true`. Without COM, returns hash data only (`success: false`).
