"""
Formatting operations: number format, cell style, alignment, column/row dimensions, borders.

Public router - delegates to focused sub-modules.
"""
from __future__ import annotations

from ._format_bulk import *  # noqa: F401, F403
from ._format_cell import *  # noqa: F401, F403
from ._format_dimensions import *  # noqa: F401, F403
