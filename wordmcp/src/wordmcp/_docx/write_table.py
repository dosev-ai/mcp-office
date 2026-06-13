from __future__ import annotations

from wordmcp._docx import _facade


def add_table(
    path: str,
    rows: int,
    cols: int,
    data: list[list[str]] | None = None,
    style: str | None = None,
    confirm: bool = False,
) -> dict:
    facade = _facade()
    facade._check_write()
    facade._check_confirm(confirm)
    resolved = facade._check_path(path)

    if rows < 1:
        raise facade.ValidationError("rows must be >= 1")
    if cols < 1:
        raise facade.ValidationError("cols must be >= 1")
    if rows > 500 or cols > 100:
        raise facade.ValidationError(
            "Table too large: rows must be <= 500 and cols must be <= 100"
        )
    if data is not None:
        if len(data) != rows:
            raise facade.ValidationError(f"data has {len(data)} rows, expected {rows}")
        for row_index, row_data in enumerate(data):
            if len(row_data) != cols:
                raise facade.ValidationError(
                    f"data row {row_index} has {len(row_data)} cols, expected {cols}"
                )

    doc = facade._load_doc(resolved)
    if style is not None:
        try:
            doc.styles[style]
        except KeyError:
            raise facade.ValidationError(f"Unknown table style: {style!r}")

    table_count_before = len(doc.tables)
    table = doc.add_table(rows=rows, cols=cols, style=style)
    if data is not None:
        for row_index, row_data in enumerate(data):
            for col_index, value in enumerate(row_data):
                table.rows[row_index].cells[col_index].text = str(value)

    facade._atomic_save(doc, resolved)
    facade._evict_doc(resolved)
    facade._audit_log("add_table", resolved, extra={"rows": rows, "cols": cols})
    return {
        "table_index": table_count_before,
        "rows": rows,
        "cols": cols,
        "style": table.style.name,
    }


def update_table_cell(
    path: str,
    table_index: int,
    row: int,
    col: int,
    new_text: str,
    confirm: bool = False,
) -> dict:
    facade = _facade()
    facade._check_write()
    facade._check_confirm(confirm)
    resolved = facade._check_path(path)
    doc = facade._load_doc(resolved)
    tables = doc.tables
    if table_index < 0 or table_index >= len(tables):
        raise facade.NotFoundError(
            f"Table index {table_index} out of range (0-{len(tables)-1})"
        )
    table = tables[table_index]
    if row < 0 or row >= len(table.rows):
        raise facade.ValidationError(f"Row {row} out of range (0-{len(table.rows)-1})")
    if col < 0 or col >= len(table.columns):
        raise facade.ValidationError(f"Col {col} out of range (0-{len(table.columns)-1})")
    table.cell(row, col).text = new_text
    facade._atomic_save(doc, resolved)
    facade._evict_doc(resolved)
    facade._audit_log("update_table_cell", resolved)
    return {"table_index": table_index, "row": row, "col": col, "text": new_text}


def bulk_update_table_cells(path: str, updates: list, confirm: bool = False) -> dict:
    facade = _facade()
    facade._check_write()
    facade._check_confirm(confirm)
    if not updates:
        raise facade.ValidationError("updates must not be empty")
    if len(updates) > 200:
        raise facade.ValidationError("max 200 updates per call")

    resolved = facade._check_path(path)
    document = facade._load_doc(resolved)
    tables = document.tables
    results = []
    errors = []

    for i, item in enumerate(updates):
        try:
            table_index = item.get("table_index")
            if table_index is None:
                raise facade.ValidationError("table_index is required")
            table_index = int(table_index)
            if table_index < 0 or table_index >= len(tables):
                raise facade.NotFoundError(
                    f"table_index {table_index} out of range (0-{len(tables) - 1})"
                )
            row = int(item.get("row", 0))
            col = int(item.get("col", 0))
            table = tables[table_index]
            if row < 0 or row >= len(table.rows):
                raise facade.ValidationError(
                    f"row {row} out of range (0-{len(table.rows) - 1})"
                )
            if col < 0 or col >= len(table.columns):
                raise facade.ValidationError(
                    f"col {col} out of range (0-{len(table.columns) - 1})"
                )
            new_text = item.get("new_text", "")
            clean = facade._CTRL_RE.sub("", str(new_text)) if new_text else ""
            table.cell(row, col).text = clean
            results.append(
                {"index": i, "table_index": table_index, "row": row, "col": col, "status": "updated"}
            )
        except Exception as exc:  # noqa: BLE001
            errors.append({"index": i, "error": f"Operation failed ({type(exc).__name__})"})

    facade._atomic_save(document, resolved)
    facade._evict_doc(resolved)
    facade._audit_log("bulk_update_table_cells", resolved)
    return {
        "updated": len(results),
        "failed": len(errors),
        "results": results,
        "errors": errors,
    }
