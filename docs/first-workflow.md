# Your First Excel MCP Workflow

The flagship workflow: **structured input → write data → style it → validate → export**.

This is the full end-to-end path — real commands, real expected output. Before you start, make sure you’ve completed the [Quick Start](quickstart.md) and confirmed that `capabilities()` works in your MCP client.

---

## What you need

- excelmcp installed and configured in your MCP client
- `EXCEL_ENABLE_WRITE=true` in your MCP config
- A folder in your `EXCEL_ALLOWLIST_ROOTS` — for this walkthrough we’ll call it your allowed folder
- Microsoft Excel installed (required for `apply_style` and `export_as_pdf`)

Create a blank workbook called `budget.xlsx` in your allowed folder before starting.

---

## Step 1 — Write structured data with `range_io`

Ask your MCP client:

```
Use range_io to write the following data to budget.xlsx,
sheet Sheet1, starting at cell A1, confirm=True:
[
  ["Category", "Amount", "Status"],
  ["Kitchen", 8500, "In progress"],
  ["Bathroom", 4200, "Complete"],
  ["Living room", 3100, "Planned"]
]
```

**Expected response:**
```json
{ "cells_written": 12 }
```

Two things are enforced automatically:
- The file path must be inside `EXCEL_ALLOWLIST_ROOTS`. Anything outside is rejected.
- `confirm=True` is required. Without it, the write is refused. This is the gate that keeps automation from running away from you.
- Any cell value starting with `=`, `+`, `-`, or `@` is rejected as a potential formula injection.

---

## Step 2 — Style the header row with `apply_style`

```
Use apply_style on A1:C1 in budget.xlsx:
bold=True, background_color="#1F3864", font_color="#FFFFFF"
```

**Expected response:**
```json
{ "cells_styled": 3, "engine": "com" }
```

If Excel has the file open, excelmcp uses the COM engine (live update, no file close needed). If Excel is not open, it uses openpyxl. The `engine` field in the response tells you which path ran.

---

## Step 3 — Validate the workbook contract

```
Run validate_contract on budget.xlsx
```

**Expected response (pass):**
```json
{ "status": "valid", "checks_passed": 4, "checks_failed": 0 }
```

`validate_contract` checks the workbook against the schema embedded in its `_MCP_META` sheet. For a fresh workbook without a registered schema, it validates structural integrity: sheet existence, header row presence, and data type consistency.

If it fails, you get a structured error:
```json
{
  "status": "invalid",
  "violations": [
    { "sheet": "Sheet1", "cell": "B3", "reason": "expected number, got string" }
  ]
}
```

This is what makes the workflow feel like infrastructure rather than scripting — the validation step is deterministic and auditable.

---

## Step 4 — Export as PDF

```
Export budget.xlsx to PDF
```

**Expected response:**
```json
{ "output_path": "C:\\...\\budget.pdf", "pages": 1 }
```

This triggers Excel’s native PDF export via COM. The output lands in the same directory as the source file. Excel must be installed for this step.

---

## The full workflow in one prompt

If you want to run the whole thing in one go, you can chain the steps:

```
Using excelmcp:
1. Write this data to budget.xlsx, Sheet1, starting at A1, confirm=True:
   [["Category","Amount","Status"],["Kitchen",8500,"In progress"],["Bathroom",4200,"Complete"],["Living room",3100,"Planned"]]
2. Apply bold + dark blue background + white font to A1:C1
3. Run validate_contract on budget.xlsx
4. Export budget.xlsx to PDF
Report the result of each step.
```

---

## What this is and isn’t

**This is:**
- Deterministic tool execution. Every step is explicit, auditable, and repeatable.
- A governed automation layer. Write gates, allowlists, and validation are not optional.
- Developer-first. You control what runs, when, and against which files.

**This is not:**
- An in-Excel Copilot replacement. excelmcp doesn’t operate inside the Excel UI.
- An autonomous agent. Nothing runs without explicit tool calls.
- A broad-suite Office automation layer yet — Excel is the lead package. Others are coming.

---

## Next steps

- Explore the full tool list: ask `capabilities()` in your MCP client
- Try `bulk_range_write` to write multiple ranges in one call
- Try `read_sheet_metadata` to inspect a workbook’s contract and schema
- Open a [First Run Report](https://github.com/dosev-ai/mcp-office/issues/new?template=first_run_report.yml) if anything was unclear or didn’t work
- Ask questions in [GitHub Discussions](https://github.com/dosev-ai/mcp-office/discussions)
