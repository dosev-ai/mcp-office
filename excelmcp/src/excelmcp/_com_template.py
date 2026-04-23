"""Backward-compat re-export shim for excelmcp COM helpers.

Logic has been split into:
    _com_recalc.py  -- _evaluate_formula_live
    _com_pivot.py   -- _create_pivot_table, _refresh_pivot_tables, _read_pivot_table

This module re-exports all four names so that existing importers of
``excelmcp._com_template`` (and ``excelmcp._com`` which imports from here)
continue to work unchanged.

No MCP imports. No win32com imports.
"""
from __future__ import annotations

from excelmcp._com_pivot import (  # noqa: F401
    _create_pivot_table,
    _read_pivot_table,
    _refresh_pivot_tables,
)
from excelmcp._com_recalc import _evaluate_formula_live  # noqa: F401

__all__ = [
    "_evaluate_formula_live",
    "_create_pivot_table",
    "_refresh_pivot_tables",
    "_read_pivot_table",
]
