"""
I/O operations: capabilities, workbook creation, sheet/cell read/write, save.

Backend logic has been extracted into focused sub-modules:
  - _range_read.py  — list_sheets, get_used_range, read_cell, read_range
  - _range_write.py — create_workbook, write_cell, write_range,
                      bulk_write_cells, append_rows, save

All functions are re-exported from this module so that existing imports
(``from excelmcp._io import ...``) continue to work without change.
"""
from __future__ import annotations

import logging

from excelmcp import __version__
from excelmcp._metadata_contract import CONTRACT_VERSION
from excelmcp.runtime_config import get_effective_policy

# Re-export all extracted helpers so callers of ``excelmcp._io`` are unaffected.
from ._range_read import (  # noqa: F401
    get_used_range,
    list_sheets,
    read_cell,
    read_range,
)
from ._range_write import (  # noqa: F401
    append_rows,
    bulk_write_cells,
    create_workbook,
    save,
    write_cell,
    write_range,
)

logger = logging.getLogger(__name__)


def capabilities(prompt_names: list[str] | None = None) -> dict:
    """Return metadata about this MCP phase and its available tools.

    prompt_names: optional live MCP prompt inventory supplied by the server layer.
    When omitted, prompt metadata defaults to an empty inventory because prompt
    registration is owned by the MCP server, not the workbook I/O layer.
    """
    prompt_inventory = list(prompt_names or [])
    result = {
        "version": __version__,
        "phase": "2.0",
        "backend": "openpyxl",
        "tools": [
            # server_io.py (17)
            "capabilities",
            "create_workbook",
            "sheet",
            "get_used_range",
            "cell",
            "range_io",
            "read_sheet_all",
            "named_range",
            "get_cell_metadata",
            "evaluate_formula",
            "list_tables",
            "read_table",
            "append_rows",
            "find_replace",
            "import_csv_to_sheet",
            "save",
            "workbook_metadata",
            # server_format.py (9 = 1 dispatcher + 8 deprecated aliases)
            "format",
            "apply_number_format",
            "apply_style",
            "apply_alignment",
            "add_border",
            "apply_format_to_sheet_list",
            "copy_range_format",
            "write_range_with_format",
            "apply_conditional_format",
            # server_ops.py (16)
            "merge",
            "freeze_panes",
            "auto_filter",
            "rows",
            "cols",
            "protect",
            "set_print_area",
            "cell_comment",
            "sheet_properties",
            "data_validation",
            "validate_formula_syntax",
            "copy_range",
            "create_table",
            "set_column_visibility",
            "set_workbook_calc_mode",
            "hyperlink",
            # server_batch.py (6)
            "bulk_copy_sheet",
            "set_row_heights",
            "set_column_widths",
            "bulk_add_comments",
            "fill_formula_range",
            "bulk_range_write",
            # server_chart.py (1)
            "chart",
            # server_com.py (5 — COM-conditional, Windows + EXCEL_ENABLE_COM=true)
            "recalculate_workbook",
            "export_as_pdf",
            "hydrate_template",
            "pivot_table",
            "export_range_as_csv",
            # server_com_session.py (2 — COM-conditional, require EXCEL_ENABLE_COM=true)
            "get_active_workbook_context",
            "get_all_open_workbooks",
            # server_review.py (3)
            "review_workbook_render",
            "produce_export_evidence_bundle",
            "export_changed_sheets_only",
            # server_acp.py (1)
            "get_workbook_context",
            # server_snapshot.py (2 -- openpyxl-based, no COM required)
            "snapshot_workbook",
            "diff_workbooks",
            # server_io.py extension (1 -- openpyxl-based consolidation)
            "consolidate_ranges",
            # server_com_vba.py (2 -- COM-conditional, require EXCEL_ENABLE_COM=true)
            "list_macros",
            "run_macro",
        ],
        "prompts": {
            "count": len(prompt_inventory),
            "names": prompt_inventory,
        },
        "governance": {
            "allowlist_roots_env": "EXCEL_ALLOWLIST_ROOTS",
            "enable_write_env":    "EXCEL_ENABLE_WRITE",
            "enable_com_env":      "EXCEL_ENABLE_COM",
            "enable_macros_env":   "EXCEL_ENABLE_MACROS",
            "max_range_cells_env": "EXCEL_MAX_RANGE_CELLS",
            "max_range_cells_default": 5000,
        },
        "metadata_contract_version": CONTRACT_VERSION,
        "metadata_policy": get_effective_policy(),
    }
    from excelmcp._server_manifest import build_capabilities_v2  # noqa: PLC0415
    result["capabilities_v2"] = build_capabilities_v2(None)
    return result
