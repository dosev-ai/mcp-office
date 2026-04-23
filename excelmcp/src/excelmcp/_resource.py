"""MCP resource resolution helpers (no FastMCP imports)."""
from __future__ import annotations

from ._core import (
    ExcelMCPError,
    _check_path,
    _load_wb,
)
from ._metadata import (
    resolve_workbook_metadata_resource as _resolve_workbook_metadata_resource,
)


def resolve_sheets_resource(path: str) -> dict:
    """Return list of all sheets with used range dimensions."""
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    results = []
    for ws in wb.worksheets:
        dims = ws.dimensions or ""
        results.append({
            "name": ws.title,
            "used_range": dims or "A1:A1",  # backward compat
            "dimensions": dims,
            "max_row": ws.max_row or 0,
            "max_column": ws.max_column or 0,
        })
    return {"sheets": results}


def resolve_sheet_resource(path: str, sheet_name: str) -> dict:
    """Return structural info about a specific sheet."""
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    if sheet_name not in wb.sheetnames:
        raise ExcelMCPError(f"Sheet not found: {sheet_name!r}")
    ws = wb[sheet_name]
    if ws.max_row is None or ws.max_column is None:
        headers = []
    else:
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    tables = []
    for t in ws.tables.values():
        try:
            cols_list = [col.name for col in t.table_columns]
        except AttributeError:
            cols_list = []
        tables.append({"name": t.name, "ref": t.ref, "columns": cols_list})
    return {
        "sheet": sheet_name,
        "dims": ws.dimensions,
        "max_row": ws.max_row or 0,
        "max_col": ws.max_column or 0,
        "first_row_headers": headers,
        "tables": tables,
    }


def resolve_named_ranges_resource(path: str) -> dict:
    """Return all named ranges defined in the workbook."""
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    results = []
    try:
        items = list(wb.defined_names.items())
    except AttributeError:
        try:
            items = [(defn.name, defn) for defn in wb.defined_names.definedName]
        except AttributeError:
            items = []
    for name, defn in items:
        try:
            destinations = list(defn.destinations)
            dest_list = [{"sheet": s, "ref": r} for s, r in destinations] if destinations else []
            results.append({"name": name, "destinations": dest_list})
        except Exception:
            results.append({"name": name, "destinations": []})
    return {"named_ranges": results}


def resolve_charts_resource(path: str) -> dict:
    """Return chart inventory for all sheets in a workbook."""
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    charts = []
    for ws in wb.worksheets:
        try:
            ws_charts = ws._charts
        except AttributeError:
            ws_charts = []
        for chart in (ws_charts or []):
            chart_type = type(chart).__name__
            anchor = str(chart.anchor) if hasattr(chart, "anchor") else ""
            data_address = None
            try:
                if chart.series:
                    data_address = chart.series[0].val.numRef.f
            except Exception:
                data_address = None
            charts.append({
                "sheet": ws.title,
                "chart_type": chart_type,
                "anchor": anchor,
                "data_address": data_address,
            })
    return {"charts": charts}


def resolve_tables_resource(path: str) -> dict:
    """Return table inventory with column headers for all sheets."""
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    tables = []
    for ws in wb.worksheets:
        try:
            for tbl in ws.tables.values():
                try:
                    cols_list = [col.name for col in tbl.table_columns]
                except AttributeError:
                    cols_list = []
                tables.append({
                    "sheet": ws.title,
                    "name": tbl.name,
                    "ref": tbl.ref,
                    "display_name": getattr(tbl, "displayName", tbl.name),
                    "columns": cols_list,
                })
        except AttributeError:
            pass
    return {"tables": tables}


def resolve_validations_resource(path: str) -> dict:
    """Return data validation inventory for all sheets."""
    resolved = _check_path(path)
    wb = _load_wb(resolved)
    validations = []
    for ws in wb.worksheets:
        try:
            for dv in ws.data_validations.dataValidation:
                try:
                    address = str(dv.sqref)
                except Exception:
                    address = ""
                validations.append({
                    "sheet": ws.title,
                    "address": address,
                    "type": dv.type or "any",
                    "formula1": dv.formula1,
                    "formula2": dv.formula2,
                    "show_dropdown": not dv.showDropDown,
                })
        except AttributeError:
            pass
    return {"validations": validations}


def resolve_workbook_metadata_resource(path: str) -> dict:
    """Return workbook-level metadata stored on the hidden _MCP_META sheet."""
    return _resolve_workbook_metadata_resource(path)
