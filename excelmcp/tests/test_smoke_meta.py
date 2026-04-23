"""Smoke tests for the excelmcp meta and registry layer."""
import pytest
from fastmcp.exceptions import ToolError

from ._smoke_helpers import _get_capabilities_snapshot, _get_capability_prompt_names, _get_prompt_names, _get_tool_names, _run_tool


@pytest.mark.smoke
def test_import_clean(capsys):
    """Importing the server module must not write anything to stdout."""
    import excelmcp.server  # noqa: F401

    captured = capsys.readouterr()
    assert captured.out == "", f"Unexpected stdout on import: {captured.out!r}"


@pytest.mark.smoke
def test_mcp_object_exists():
    from excelmcp import server

    assert server.mcp is not None


@pytest.mark.smoke
def test_tools_registered():
    from excelmcp import server

    mcp = server.mcp
    tool_names = _get_tool_names(mcp)
    caps = _get_capabilities_snapshot(mcp)

    expected = {
        # Phase 1
        "capabilities",
        "sheet",
        "get_used_range",
        "cell",
        "range_io",
        "append_rows",
        "save",
        "workbook_metadata",
        # Phase 1.5
        "read_sheet_all",
        "find_replace",
        "named_range",
        "rows",
        "cols",
        "import_csv_to_sheet",
        "get_cell_metadata",
        "evaluate_formula",
        "list_tables",
        "read_table",
        # Phase 2.0
        "create_workbook",
        "apply_number_format",
        "apply_style",
        "apply_alignment",
        "add_border",
        "chart",
        # Phase 2.2
        "merge",
        "freeze_panes",
        "auto_filter",
        # Phase 2.3
        "rows",
        "cols",
        "protect",
        "set_print_area",
        # Phase 2.4
        "cell_comment",
        "sheet_properties",
        # Phase 2.5
        "validate_formula_syntax",
        "copy_range",
        "create_table",
        # Phase 2.6
        "data_validation",
        "cell_comment",
        "set_column_visibility",
        "set_workbook_calc_mode",
        # Phase 2.7
        "hyperlink",
        # Phase 2.7 Sprint C — bulk operations + security backport
        "bulk_copy_sheet",
        "set_row_heights",
        "apply_style",
        "set_column_widths",
        "apply_format_to_sheet_list",
        # Sprint D Batch 1
        "bulk_add_comments",
        "fill_formula_range",
        # Sprint D Batch 2
        "copy_range_format",
        # Sprint D Batch 3
        "write_range_with_format",
        "apply_conditional_format",
        # Sprint G — COM Phase 2 tools
        "recalculate_workbook",
        "export_as_pdf",
        "hydrate_template",
        "pivot_table",
    }
    missing = expected - tool_names
    assert not missing, f"Missing registered tools: {missing}"
    assert "render_review" not in tool_names
    assert "render_review" not in caps["tools"]
    assert tool_names == caps["tools"], (
        "capabilities() tool inventory drifted from the live registry: "
        f"live_only={sorted(tool_names - caps['tools'])}, "
        f"caps_only={sorted(caps['tools'] - tool_names)}"
    )


@pytest.mark.smoke
def test_prompts_registered():
    from excelmcp import server

    mcp = server.mcp
    prompt_names = _get_prompt_names(mcp)
    capability_prompt_names = _get_capability_prompt_names(mcp)
    caps = _get_capabilities_snapshot(mcp)

    expected = {"summarize_sheet", "suggest_formatting", "prepare_for_print"}
    missing = expected - prompt_names
    assert not missing, f"Missing registered prompts: {missing}"
    assert caps["prompts"] == capability_prompt_names, (
        "capabilities() prompt inventory drifted from the live registry: "
        f"live_only={sorted(capability_prompt_names - caps['prompts'])}, "
        f"caps_only={sorted(caps['prompts'] - capability_prompt_names)}"
    )
    assert caps["prompt_count"] == len(capability_prompt_names), (
        "capabilities() prompt count drifted from its advertised prompt set: "
        f"prompt_count={caps['prompt_count']} live_count={len(capability_prompt_names)}"
    )


@pytest.mark.smoke
def test_capabilities_normalizes_prompt_lookup_failure(monkeypatch):
    from excelmcp import server
    from excelmcp import server_io

    async def _failing_list_prompts():
        raise RuntimeError("prompt registry unavailable")

    monkeypatch.setattr(server_io.mcp, "list_prompts", _failing_list_prompts)

    with pytest.raises(ToolError, match="prompt registry unavailable"):
        _run_tool(server.mcp, "capabilities", {})


@pytest.mark.smoke
def test_capabilities_bridges_live_prompt_names_into_manifest(monkeypatch):
    from excelmcp import server_io

    captured: dict[str, object] = {}

    def _fake_capabilities(*, prompt_names):
        captured["prompt_names"] = prompt_names
        return {
            "version": "test-version",
            "phase": "2.0",
            "backend": "openpyxl",
            "tools": ["capabilities"],
            "prompts": {
                "count": len(prompt_names),
                "names": list(prompt_names),
            },
            "governance": {},
        }

    monkeypatch.setattr(
        server_io,
        "_get_live_prompt_names",
        lambda mcp_server=server_io.mcp: ["suggest_formatting", "prepare_for_print"],
    )
    monkeypatch.setattr(server_io.wb, "capabilities", _fake_capabilities)

    result = server_io.capabilities()

    assert captured["prompt_names"] == ["suggest_formatting", "prepare_for_print"]
    assert result["prompts"] == {
        "count": 2,
        "names": ["suggest_formatting", "prepare_for_print"],
    }
    assert result["backend"] == "openpyxl"
