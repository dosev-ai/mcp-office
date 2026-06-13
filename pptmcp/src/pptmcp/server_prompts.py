"""
Prompt definitions for pptmcp. Registered via register_prompts(mcp).
Separated from server.py to keep server.py under 800 lines (WARN-04).
"""
from __future__ import annotations

import json

_PROMPT_DESCRIPTION_V1 = (
    "Canonical agent workflow: Output Contract \u2192 Build \u2192 Export PNG/PDF"
    " \u2192 Self-Review \u2192 Iterate \u2192 Evidence Bundle. "
    "Use this prompt whenever building or editing a PowerPoint deck."
)


def ppt_render_check_iterate_v1() -> str:
    """Return the canonical Render-Check-Iterate workflow prompt for the agent."""
    output_contract_schema = json.dumps(
        {
            "kind": "pptx",
            "template": "<theme_or_style_anchor>",
            "slides": [
                {
                    "slide_number": 1,
                    "title": "<slide_title>",
                    "layout": "<layout_name>",
                    "shapes": [
                        {
                            "name": "<shape_name>",
                            "type": "TEXTBOX | PICTURE | TABLE | ...",
                            "text": "<content_or_null>",
                        }
                    ],
                }
            ],
            "export_plan": {
                "png_slides": "changed",
                "pdf_milestone": True,
            },
            "acceptance": {
                "overflow": True,
                "outside_slide": True,
                "min_dpi": 150,
            },
            "evidence_paths": {
                "png_dir": "<output_dir>/run-<timestamp>/png",
                "pdf_path": "<output_dir>/run-<timestamp>/deck.pdf",
            },
        },
        indent=2,
    )

    return f"""# PowerPoint Render-Check-Iterate Workflow v1

## Required Environment Gates
Before starting, ensure:
- PPT_ENABLE_WRITE=true  (all build/export steps)
- PPT_ALLOWLIST_ROOTS=<output_dir>  (your working directory)
- confirm=True on every write/export call

---

## STEP 1 \u2014 Output Contract
Produce `output_contract.json` BEFORE writing any slides.
This contract is the acceptance specification for the entire build.

Contract schema:
```json
{output_contract_schema}
```

Store the contract at `<output_dir>/output_contract.json`.
Use it in Step 5 (Self-Review) to verify every requirement is met.

---

## STEP 2 \u2014 Build
Implement all slides using pptmcp authoring tools:
- `create_presentation` \u2192 `add_slide` \u2192 `add_textbox` / `add_shape` / `add_table_to_slide` / `insert_image`
- `set_text_format` / `set_paragraph_format` for typography
- For every text shape: word_wrap=True (prevents text overflow)
- `save` (+ confirm=True) after every meaningful change

Follow the slide inventory in `output_contract.json` exactly.
Do not add slides not listed in the contract.

---

## STEP 3 \u2014 Export PNG Renders  [requires PPT_ENABLE_COM=true]
Export a PNG for every slide (or changed slides only on re-runs):

```
export(
    scope="slide",
    path="<output_dir>/<deck>.pptx",
    fmt="png",
    output_path="<output_dir>/run-<timestamp>/png/slide-<NNN>.png",
    return_inline=True,
    response_mode="auto",
    options={{"slide_index": <0-based>, "inline_dpi": 150}},
    confirm=True,
)
```

Naming convention: `slide-001.png`, `slide-002.png`, ... (zero-padded to 3 digits).
If PPT_ENABLE_COM is not set, document the gap in the evidence bundle and proceed to Step 5
using file-based heuristics only.

---

## STEP 4 \u2014 Milestone PDF Export  [requires PPT_ENABLE_COM=true]
For end-to-end milestone checks:

```
export(
    scope="deck_pdf",
    path="<output_dir>/<deck>.pptx",
    output_path="<output_dir>/run-<timestamp>/deck.pdf",
    confirm=True,
)
```

---

## STEP 5 \u2014 Self-Review Gate (iterate until PASS)
For each exported PNG, review against `output_contract.json`:

Checklist per slide:
- [ ] Text overflow: any text clipped or overflowing shape boundary?
- [ ] Shapes outside slide area: any shape partially or fully outside slide bounds?
- [ ] Layout drift: shape positions match contract geometry?
- [ ] Font/style: text formatted per contract (size, bold, colour)?
- [ ] Table sizing: rows/columns match contract table spec?

If ANY check FAILS:
1. Apply fixes using pptmcp tools (`edit_text_placeholder`, `set_text_format`, `add_textbox`, etc.)
2. Save (confirm=True)
3. Re-export changed slides only with `export(scope="slide", fmt="png", ...)` for each fixed slide
4. Re-run checklist for fixed slides
5. Repeat loop until ALL checks PASS or a maximum of 3 iterations is reached

Do not declare PASS until every slide's checklist is fully green.

---

## STEP 6 \u2014 Evidence Bundle
Produce `evidence.json` at `<output_dir>/evidence.json`:

```json
{{
  "contract_path": "<output_dir>/output_contract.json",
  "slide_exports": [
    {{"slide_index": 0, "png_path": "<output_dir>/run-<timestamp>/png/slide-001.png"}}
  ],
  "pdf_path": "<output_dir>/run-<timestamp>/deck.pdf",
  "review_results": [
    {{
      "slide_index": 0,
      "checks": [
        {{"criterion": "text_overflow", "result": "PASS", "detail": ""}},
        {{"criterion": "outside_slide", "result": "PASS", "detail": ""}}
      ],
      "verdict": "PASS"
    }}
  ],
  "final_verdict": "PASS",
  "tool_calls_used": <integer count>,
  "iterations": <number of Step 5 loops run>
}}
```

Attach the evidence.json path to the Action Worker record for this task.
"""


def register_prompts(mcp) -> None:  # type: ignore[type-arg]
    """Register all pptmcp prompts on the provided FastMCP instance."""
    mcp.prompt(
        name="ppt_render_check_iterate_v1",
        description=_PROMPT_DESCRIPTION_V1,
    )(ppt_render_check_iterate_v1)
