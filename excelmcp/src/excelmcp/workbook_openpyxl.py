# ruff: noqa: F401, F403, F405
"""
Excel MCP workbook operations ÔÇö public re-export facade.

All logic lives in the sub-modules below. Import from this module for
backward compatibility; server.py and tests need zero changes.
"""
from ._advanced import *
from ._chart import *
from ._core import *

# Private symbols needed by tests (not re-exported by wildcard imports)
from ._core import _wb_cache, _WbCache
from ._data import *
from ._format import *
from ._io import *
from ._metadata import *
from ._ooxml import patch_dynamic_array_metadata
from ._resource import *
from ._sheet import *

__all__ = [
    # Exceptions
    "ExcelMCPError",
    "ValidationError",
    "NotAllowedError",
    "OfficeCOMError",
    # IO
    "capabilities",
    "create_workbook",
    "list_sheets",
    "get_used_range",
    "read_cell",
    "read_range",
    "write_cell",
    "write_range",
    "append_rows",
    "save",
    "read_workbook_metadata",
    "write_workbook_metadata",
    # Sheet
    "read_sheet_all",
    "find_replace",
    "add_sheet",
    "rename_sheet",
    "delete_sheet",
    "delete_rows",
    "insert_rows",
    "delete_cols",
    "insert_cols",
    "import_csv_to_sheet",
    "get_sheet_properties",
    "set_sheet_properties",
    "protect_workbook",
    "set_workbook_calc_mode",
    # Data
    "get_cell_metadata",
    "evaluate_formula",
    "list_tables",
    "read_table",
    "get_named_ranges",
    "read_named_range",
    "write_named_range",
    "delete_cell_comment",
    # Format
    "apply_number_format",
    "apply_cell_style",
    "apply_alignment",
    "set_column_visibility",
    "add_border",
    # Chart
    "create_chart",
    "list_charts",
    "delete_chart",
    "update_chart_data",
    # Advanced
    "merge_cells",
    "unmerge_cells",
    "freeze_panes",
    "auto_filter",
    "protect_sheet",
    "set_print_area",
    # Phase 2.5
    "validate_formula_syntax",
    "list_merged_cells",
    "copy_range",
    "create_table",
    # Phase 2.6
    "get_data_validation_info",
    "get_cell_comment",
    # Phase 2.7
    "get_hyperlinks",
    "remove_hyperlink",
    # Phase 2.7 Sprint C ÔÇö bulk operations + security backport
    "bulk_copy_sheet",
    "set_row_heights",
    "apply_style_to_ranges",
    "set_column_widths",
    "bulk_add_hyperlinks",
    "apply_format_to_sheet_list",
    # Sprint D Batch 1 ÔÇö new bulk + formula fill
    "bulk_add_comments",
    "fill_formula_range",
    # Sprint D Batch 2 ÔÇö bulk DV, copy format
    "bulk_add_data_validation",
    "copy_range_format",
    # Sprint D Batch 3 ÔÇö write with format, conditional format
    "write_range_with_format",
    "apply_conditional_format",
    # Route A: bulk non-contiguous cell write
    "bulk_write_cells",
    # Resources
    "resolve_sheets_resource",
    "resolve_sheet_resource",
    "resolve_named_ranges_resource",
    "resolve_charts_resource",
    "resolve_tables_resource",
    "resolve_validations_resource",
    "resolve_workbook_metadata_resource",
]
