"""
Data introspection: cell metadata, formula evaluation, tables, named ranges.

Public router - delegates to focused sub-modules.
"""
from __future__ import annotations

from ._data_annotations import *  # noqa: F401, F403
from ._data_cell import *  # noqa: F401, F403
from ._data_tables import *  # noqa: F401, F403
from ._data_write import *  # noqa: F401, F403
