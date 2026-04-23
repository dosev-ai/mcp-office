"""Analysis prompts: sheet summarisation, anomaly detection, formula writing, range comparison.

Registers 5 @mcp.prompt() on the shared mcp instance at import time.
"""
from __future__ import annotations

import logging

from excelmcp._server_instance import mcp

logger = logging.getLogger(__name__)


@mcp.prompt()
def summarize_sheet(sheet_json: str) -> str:
    """Produce a concise business summary of an Excel sheet."""
    return (
        f"Given this Excel sheet data JSON:\n{sheet_json}\n\n"
        "Produce a concise business summary covering: key metrics, trends, "
        "data quality issues, and actionable insights."
    )


@mcp.prompt()
def find_anomalies(range_json: str) -> str:
    """Flag statistical outliers, missing values, and duplicates in a range."""
    return (
        f"Scan this numeric range data:\n{range_json}\n\n"
        "Flag statistical outliers, missing values, impossible values, and duplicates. "
        "Include the cell address for each anomaly found."
    )


@mcp.prompt()
def write_formula(description: str, column_layout: str) -> str:
    """Generate an Excel formula for the described calculation."""
    return (
        f"Column layout:\n{column_layout}\n\n"
        f"Write an Excel formula (or set of formulas) to accomplish:\n{description}\n\n"
        "Include the formula, where to enter it, and a brief explanation."
    )


@mcp.prompt()
def compare_ranges(range_a: str, range_b: str) -> str:
    """Summarise values added, removed, and changed between two ranges."""
    return (
        f"Compare these two Excel ranges:\nRange A:\n{range_a}\n\n"
        f"Range B:\n{range_b}\n\n"
        "Summarise values added, removed, and changed. Group by significance."
    )


@mcp.prompt()
def extract_structured_table(raw_range: str) -> str:
    """Convert a messy Excel range to a clean JSON table with inferred types."""
    return (
        f"Convert this messy Excel range to a clean JSON table:\n{raw_range}\n\n"
        'Output: {"headers": [...], "rows": [[...], ...]} with inferred types '
        "(string, number, date, boolean)."
    )
