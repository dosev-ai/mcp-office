from pathlib import Path
import re


DOC_PATH = Path(__file__).resolve().parent.parent / "docs" / "render-check-iterate.md"
README_PATH = Path(__file__).resolve().parent.parent / "README.md"
SERVER_PATH = Path(__file__).resolve().parent.parent / "src" / "excelmcp" / "server.py"


def test_render_check_iterate_uses_unified_export_as_pdf_name() -> None:
    text = DOC_PATH.read_text(encoding="utf-8")

    assert "export_as_pdf" in text
    assert "export_sheet_as_pdf" not in text
    assert "export_workbook_as_pdf" not in text


def test_render_check_iterate_preserves_truth_boundary() -> None:
    text = DOC_PATH.read_text(encoding="utf-8").lower()

    assert "no automated visual scoring" in text
    assert "human review" in text or "manual review" in text


def test_render_check_iterate_documents_output_parent_directory_requirement() -> None:
    text = DOC_PATH.read_text(encoding="utf-8")

    assert "The PDF output parent directory must already exist." in text
    assert "rejects a missing parent directory or a non-directory parent" in text


def test_render_check_iterate_scopes_workbook_export_contract_fields() -> None:
    text = DOC_PATH.read_text(encoding="utf-8")

    sheet_section = text.split("### Sheet export", 1)[1].split("### Workbook export", 1)[0]
    workbook_section = text.split("### Workbook export", 1)[1].split("## Recommended loop", 1)[0]

    assert "`source`" not in sheet_section
    assert "`source`" in workbook_section
    assert "`elapsed_ms`" in workbook_section
    assert "must omit `sheet`" in workbook_section


def test_render_check_iterate_documents_invalid_request_preservation_contract() -> None:
    text = DOC_PATH.read_text(encoding="utf-8")

    assert "Invalid export requests are rejected before any existing PDF target is cleared." in text
    assert "A pre-existing target is preserved on validation failure, including a stray `sheet` argument on workbook scope." in text


def test_readme_uses_correct_sprint_o_export_replacement_mapping() -> None:
    text = README_PATH.read_text(encoding="utf-8")
    current_state = text.split("## Changelog", 1)[0]

    assert "| `export_sheet_as_pdf`, `export_workbook_as_pdf` | `export_as_pdf(scope='sheet'\\|'workbook')` |" in text
    assert "| `open_template_and_recalculate` | `hydrate_template(cell_writes=[])` |" in text
    assert "| `evaluate_formula_live` | `evaluate_formula(live=True)` |" in text
    assert "| `create_pivot_table`, `refresh_pivot_tables`, `read_pivot_table` | `pivot_table(operation='create'\\|'read'\\|'refresh')` |" in text
    assert "All require `EXCEL_ENABLE_COM=true` + `confirm=True`." not in current_state
    assert "`pivot_table(operation='create'|'refresh')`, and `export_range_as_csv` (write path) also require `confirm=True`; `pivot_table(operation='read')` and `export_range_as_csv` (console-only path) do not" in current_state
    assert "`evaluate_formula(live=True)`" in current_state
    for legacy_name in (
        "export_sheet_as_pdf",
        "export_workbook_as_pdf",
        "open_template_and_recalculate",
        "evaluate_formula_live",
        "create_pivot_table",
        "refresh_pivot_tables",
        "read_pivot_table",
    ):
        assert not re.search(rf"\b{re.escape(legacy_name)}\b", current_state), (
            f"README current-state still references removed tool {legacy_name}"
        )
    assert "no COM needed" not in text
    assert "- [ ] Live formula recalculation via `Excel.Application`" not in text
    assert "- [ ] Pivot table read, refresh, and field modification" not in text
    assert "- [ ] Named range write support" not in text
    assert "- [ ] `save_as` to a different path / format (PDF export)" not in text


def test_readme_export_contract_mentions_fresh_vs_preserved_targets() -> None:
    text = README_PATH.read_text(encoding="utf-8")
    current_state = text.split("## Changelog", 1)[0]

    assert "any stale target is cleared immediately before export so evidence always refers to fresh bytes" in current_state
    assert "calling with `cell_writes=[]` performs a pure recalc/save-as flow" in current_state


def test_server_export_prompts_use_unified_export_payload_contract() -> None:
    server_text = SERVER_PATH.read_text(encoding="utf-8")
    # After the Phase 3 prompt split, export prompt bodies live in server_prompts_workflow.py
    workflow_path = SERVER_PATH.parent / "server_prompts_workflow.py"
    prompt_text = workflow_path.read_text(encoding="utf-8")

    # server.py must NOT contain deprecated shim names (W5 ADJUSTMENT)
    assert "export_sheet_as_pdf" not in server_text
    assert "export_workbook_as_pdf" not in server_text
    # Evidence contract fields must be documented in the workflow prompts module
    assert "sha256" in prompt_text
    assert "size_bytes" in prompt_text
    assert "exported_at" in prompt_text
    assert "absolute output path ending in .pdf" in prompt_text
    assert "absolute source workbook path" in prompt_text
    assert 'scope: "sheet" or "workbook"' in prompt_text or 'scope="sheet"' in prompt_text or "scope" in prompt_text
