"""COM-backed MCP tools for excelmcp.

Thin dispatcher -- registers tools via side-effect imports of sub-modules.
Tools registered on shared mcp instance at import time.

Sub-modules:
    server_com_export  -- export_as_pdf, export_range_as_csv
    server_com_recalc  -- recalculate_workbook, hydrate_template
    server_com_pivot   -- pivot_table
"""
from __future__ import annotations

import excelmcp.server_com_export  # noqa: F401 -- side-effect: registers export tools
import excelmcp.server_com_pivot  # noqa: F401 -- side-effect: registers pivot tools
import excelmcp.server_com_recalc  # noqa: F401 -- side-effect: registers recalc tools
import excelmcp.server_com_session  # noqa: F401 -- side-effect: registers session tools
import excelmcp.server_com_vba  # noqa: F401 -- side-effect: registers VBA/macro tools
from excelmcp._core import (  # noqa: F401 -- re-exported for test assertions
    NotAllowedError,
    OfficeCOMError,
    ValidationError,
)
from excelmcp.server_com_export import (  # noqa: F401
    export_as_pdf,
    export_range_as_csv,
    export_sheet_as_pdf,
    export_workbook_as_pdf,
)
from excelmcp.server_com_pivot import (  # noqa: F401
    create_pivot_table,
    evaluate_formula_live,
    open_template_and_recalculate,
    pivot_table,
    read_pivot_table,
    refresh_pivot_tables,
)
from excelmcp.server_com_recalc import (  # noqa: F401
    hydrate_template,
    recalculate_workbook,
)
from excelmcp.server_com_vba import (  # noqa: F401
    list_macros,
    run_macro,
)
