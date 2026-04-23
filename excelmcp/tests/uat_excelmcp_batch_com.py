"""UAT User Stories: excelmcp server_batch (5 tools) + chart + COM-backed (4 tools).

Batch tools (always run):
    US-BCH-01  bulk_copy_sheet
    US-BCH-02  set_row_heights
    US-BCH-03  set_column_widths
    US-BCH-04  bulk_add_comments
    US-BCH-05  fill_formula_range
    US-CHT-01  chart

COM tools (skip if not Windows):
    US-COM-02  recalculate_workbook
    US-COM-03  export_as_pdf
    US-COM-04  hydrate_template
    US-COM-05  pivot_table

Run: python excel/tests/uat_excelmcp_batch_com.py

Set EXCELMCP_UAT_ENABLE_COM=true to opt into the live COM stories.
"""
import json
import subprocess
import sys
import os
import threading
from pathlib import Path

# ---------------------------------------------------------------------------
# Harness
# ---------------------------------------------------------------------------

PYTHON = sys.executable
REPO = Path(__file__).parent.parent.parent
SERVER_CMD = [PYTHON, "-m", "excelmcp.server"]
ENV = {
    **os.environ,
    "PYTHONPATH": str(REPO / "excel" / "src"),
    "EXCEL_ENABLE_WRITE": "true",
    "EXCEL_ENABLE_COM": "true",
    "EXCEL_ALLOWLIST_ROOTS": r"C:\Temp",
    "EXCEL_EXPORT_ROOTS": r"C:\Temp",
}
WORKDIR = Path("C:/Temp")
PASS_COUNT = 0
FAIL_COUNT = 0
SKIP_COUNT = 0


class McpSession:
    def __init__(self):
        self.proc = subprocess.Popen(
            SERVER_CMD,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            env=ENV,
            cwd=str(REPO),
        )
        assert self.proc.stdin is not None
        assert self.proc.stdout is not None
        self._id = 0
        init = json.dumps({
            "jsonrpc": "2.0", "id": 0, "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "uat", "version": "1"},
            },
        }) + "\n"
        self.proc.stdin.write(init.encode())
        self.proc.stdin.flush()
        self.proc.stdout.readline()
        init_note = json.dumps({
            "jsonrpc": "2.0",
            "method": "notifications/initialized",
            "params": {},
        }) + "\n"
        self.proc.stdin.write(init_note.encode())
        self.proc.stdin.flush()

    def call(self, tool, args, timeout=30):
        self._id += 1
        msg = json.dumps({
            "jsonrpc": "2.0", "id": self._id,
            "method": "tools/call",
            "params": {"name": tool, "arguments": args},
        }) + "\n"
        assert self.proc.stdin is not None
        assert self.proc.stdout is not None
        stdout = self.proc.stdout
        self.proc.stdin.write(msg.encode())
        self.proc.stdin.flush()
        result = [b""]
        def _read():
            result[0] = stdout.readline()
        t = threading.Thread(target=_read, daemon=True)
        t.start()
        t.join(timeout=timeout)
        if t.is_alive():
            # Timed out — server is stuck (likely COM hang). Return synthetic error.
            return {
                "jsonrpc": "2.0", "id": self._id,
                "result": {"content": [{"type": "text", "text": f"[UAT TIMEOUT] tool={tool} did not respond within {timeout}s"}], "isError": True},
            }
        raw = result[0]
        return json.loads(raw) if raw else {}

    def close(self):
        try:
            assert self.proc.stdin is not None
            self.proc.stdin.close()
        except Exception:
            pass
        self.proc.wait(timeout=5)


def check(tc_id, resp, desc, should_fail=False):
    global PASS_COUNT, FAIL_COUNT
    is_err = "error" in resp
    content = resp.get("result", {})
    if not is_err and isinstance(content, dict) and content.get("isError"):
        is_err = True
    if not is_err and isinstance(content, list):
        for item in content:
            if isinstance(item, dict) and item.get("isError"):
                is_err = True
                break
    ok = is_err == should_fail
    sym = "PASS" if ok else "FAIL"
    PASS_COUNT += ok
    FAIL_COUNT += (not ok)
    print(f"  {sym} {tc_id}: {desc}")
    if not ok:
        print(f"       resp: {str(resp)[:200]}")
    return content


def parse_result(resp):
    """Extract parsed JSON dict from a FastMCP tools/call response."""
    result = resp.get("result", {})
    if isinstance(result, dict) and "content" in result:
        items = result["content"]
        if items and items[0].get("type") == "text":
            try:
                return json.loads(items[0]["text"])
            except (json.JSONDecodeError, TypeError):
                return {"_raw": items[0]["text"]}
    return result


def _mc(tc_id, cond, desc):
    """Manual assertion: PASS if cond is truthy, FAIL otherwise."""
    global PASS_COUNT, FAIL_COUNT
    ok = bool(cond)
    sym = "PASS" if ok else "FAIL"
    PASS_COUNT += ok
    FAIL_COUNT += (not ok)
    print(f"  {sym} {tc_id}: {desc}")


# ---------------------------------------------------------------------------
# COM availability guard (added after harness per spec)
# ---------------------------------------------------------------------------

def _has_live_excel() -> bool:
    """Return True only when Excel COM automation can be started."""
    if sys.platform != "win32":
        return False
    try:
        import pythoncom
        import win32com.client as win32c
    except ImportError:
        return False

    initialized = False
    excel = None
    try:
        pythoncom.CoInitialize()
        initialized = True
        excel = win32c.Dispatch("Excel.Application")
        excel.Quit()
        return True
    except Exception:
        return False
    finally:
        if excel is not None:
            del excel
        if initialized:
            pythoncom.CoUninitialize()


COM_AVAILABLE = (
    os.environ.get("EXCELMCP_UAT_ENABLE_COM", "").strip().lower() == "true"
    and _has_live_excel()
)


def _skip(tc_id, desc):
    global SKIP_COUNT
    SKIP_COUNT += 1
    print(f"  SKIP {tc_id}: {desc}")


# ---------------------------------------------------------------------------
# Setup helpers
# ---------------------------------------------------------------------------

def _create(s, path, sheets=None):
    args = {"path": path, "confirm": True, "overwrite": True}
    if sheets:
        args["sheets"] = sheets
    s.call("create_workbook", args)


def _write_cell(s, path, sheet, address, value):
    s.call("cell", {
        "operation": "write",
        "path": path, "sheet": sheet,
        "address": address, "value": value, "confirm": True,
    })


# ===========================================================================
# US-BCH-01  bulk_copy_sheet
# ===========================================================================

def test_bulk_copy_sheet(s):
    print("\n[US-BCH-01] bulk_copy_sheet")
    src = r"C:\Temp\uat_bch_copy_src.xlsx"
    _create(s, src, sheets=["Sheet1", "Sheet2"])

    # 01a: copy Sheet1 -> one new sheet name -> success
    resp = s.call("bulk_copy_sheet", {
        "path": src, "source_sheet": "Sheet1",
        "new_names": ["Copy1"], "confirm": True,
    })
    check("US-BCH-01a", resp, "copy Sheet1 -> [Copy1] (single name) -> success")

    # 01b: copy Sheet1 -> two new sheet names simultaneously -> success
    resp = s.call("bulk_copy_sheet", {
        "path": src, "source_sheet": "Sheet1",
        "new_names": ["Copy2", "Copy3"], "confirm": True,
    })
    check("US-BCH-01b", resp, "copy Sheet1 -> [Copy2, Copy3] (two names) -> success")

    # 01b2: verify both new sheets appear in workbook
    resp2 = s.call("sheet", {"operation": "list", "path": src})
    d2 = parse_result(resp2)
    sheets = d2.get("sheets", []) if isinstance(d2, dict) else []
    _mc("US-BCH-01b2", "Copy2" in sheets and "Copy3" in sheets,
        f"Copy2 and Copy3 present after bulk copy (got {sheets})")

    # 01c: insert_after specified -> sheets inserted at correct position
    resp = s.call("bulk_copy_sheet", {
        "path": src, "source_sheet": "Sheet1",
        "new_names": ["AfterCopy"], "insert_after": "Sheet1", "confirm": True,
    })
    check("US-BCH-01c", resp, "bulk_copy_sheet with insert_after='Sheet1' -> success")

    # 01d: non-existent source_sheet -> error (expected)
    resp = s.call("bulk_copy_sheet", {
        "path": src, "source_sheet": "NoSuchSheet",
        "new_names": ["WillFail"], "confirm": True,
    })
    check("US-BCH-01d", resp, "non-existent source_sheet -> error (expected)", should_fail=True)

    # 01e: confirm=False -> write gate error (expected)
    resp = s.call("bulk_copy_sheet", {
        "path": src, "source_sheet": "Sheet1",
        "new_names": ["GateCopy"], "confirm": False,
    })
    check("US-BCH-01e", resp, "confirm=False -> write gate error (expected)", should_fail=True)

    # 01f: duplicate names in new_names list -> error (expected)
    resp = s.call("bulk_copy_sheet", {
        "path": src, "source_sheet": "Sheet1",
        "new_names": ["DupA", "DupA"], "confirm": True,
    })
    check("US-BCH-01f", resp, "duplicate new_names=['DupA','DupA'] -> error (expected)", should_fail=True)


# ===========================================================================
# US-BCH-02  set_row_heights
# ===========================================================================

def test_set_row_heights(s):
    print("\n[US-BCH-02] set_row_heights")
    p = r"C:\Temp\uat_bch_rowh.xlsx"
    _create(s, p, sheets=["Sheet1"])

    # 02a: set single row height via rows=[1], height=25
    resp = s.call("set_row_heights", {
        "path": p, "sheet": "Sheet1",
        "rows": [1], "height": 25.0, "confirm": True,
    })
    check("US-BCH-02a", resp, "set rows=[1] height=25 -> success")

    # 02b: set multiple rows uniformly (rows=[2,3,4], height=20)
    resp = s.call("set_row_heights", {
        "path": p, "sheet": "Sheet1",
        "rows": [2, 3, 4], "height": 20.0, "confirm": True,
    })
    check("US-BCH-02b", resp, "set rows=[2,3,4] height=20 uniformly -> success")

    # 02c: row_range '1:5', height=30
    resp = s.call("set_row_heights", {
        "path": p, "sheet": "Sheet1",
        "row_range": "1:5", "height": 30.0, "confirm": True,
    })
    check("US-BCH-02c", resp, "row_range='1:5' height=30 -> success")

    # 02d: row_specs per-row mode (3 rows with different heights)
    resp = s.call("set_row_heights", {
        "path": p, "sheet": "Sheet1",
        "row_specs": [
            {"row": 1, "height": 18.0},
            {"row": 2, "height": 24.0},
            {"row": 3, "height": 36.0},
        ],
        "confirm": True,
    })
    check("US-BCH-02d", resp, "row_specs per-row mode 3 rows different heights -> success")

    # 02e: invalid height > 409 -> error (expected)
    resp = s.call("set_row_heights", {
        "path": p, "sheet": "Sheet1",
        "rows": [1], "height": 500.0, "confirm": True,
    })
    check("US-BCH-02e", resp, "height=500 (> 409 max) -> error (expected)", should_fail=True)

    # 02f: confirm=False -> write gate error (expected)
    resp = s.call("set_row_heights", {
        "path": p, "sheet": "Sheet1",
        "rows": [1], "height": 20.0, "confirm": False,
    })
    check("US-BCH-02f", resp, "confirm=False -> write gate error (expected)", should_fail=True)


# ===========================================================================
# US-BCH-03  set_column_widths
# ===========================================================================

def test_set_column_widths(s):
    print("\n[US-BCH-03] set_column_widths")
    p = r"C:\Temp\uat_bch_colw.xlsx"
    _create(s, p, sheets=["Sheet1"])
    # Write content so autofit has something to measure
    for addr, val in [("A1", "Hello World Long Text"), ("B1", "Short"), ("C1", "Medium text here")]:
        _write_cell(s, p, "Sheet1", addr, val)

    # 03a: set single column width (columns=['A'], width=15)
    resp = s.call("set_column_widths", {
        "path": p, "sheet": "Sheet1",
        "columns": ["A"], "width": 15.0, "confirm": True,
    })
    check("US-BCH-03a", resp, "set col A width=15 -> success")

    # 03b: set multiple columns to same width (columns=['A','B','C'])
    resp = s.call("set_column_widths", {
        "path": p, "sheet": "Sheet1",
        "columns": ["A", "B", "C"], "width": 12.0, "confirm": True,
    })
    check("US-BCH-03b", resp, "set cols A,B,C width=12 -> success")

    # 03c: col_range 'A:D', uniform width
    resp = s.call("set_column_widths", {
        "path": p, "sheet": "Sheet1",
        "col_range": "A:D", "width": 10.0, "confirm": True,
    })
    check("US-BCH-03c", resp, "col_range='A:D' width=10 -> success")

    # 03d: autofit=True — width inferred from cell content
    resp = s.call("set_column_widths", {
        "path": p, "sheet": "Sheet1",
        "columns": ["A", "B", "C"], "autofit": True, "confirm": True,
    })
    check("US-BCH-03d", resp, "autofit=True for cols A,B,C -> success")

    # 03e: providing both columns AND col_range -> error (ambiguous)
    resp = s.call("set_column_widths", {
        "path": p, "sheet": "Sheet1",
        "columns": ["A"], "col_range": "A:B", "width": 10.0, "confirm": True,
    })
    check("US-BCH-03e", resp, "columns + col_range both provided -> error (expected)", should_fail=True)

    # 03f: confirm=False -> write gate error (expected)
    resp = s.call("set_column_widths", {
        "path": p, "sheet": "Sheet1",
        "columns": ["A"], "width": 10.0, "confirm": False,
    })
    check("US-BCH-03f", resp, "confirm=False -> write gate error (expected)", should_fail=True)


# ===========================================================================
# US-BCH-04  bulk_add_comments
# ===========================================================================

def test_bulk_add_comments(s):
    print("\n[US-BCH-04] bulk_add_comments")
    p = r"C:\Temp\uat_bch_comments.xlsx"
    _create(s, p, sheets=["Sheet1"])

    # 04a: add single comment -> comments_added=1
    resp = s.call("bulk_add_comments", {
        "path": p, "sheet": "Sheet1",
        "comments": [{"address": "A1", "text": "First comment"}],
        "confirm": True,
    })
    check("US-BCH-04a", resp, "add single comment to A1 -> success")
    d = parse_result(resp)
    _mc("US-BCH-04a2", isinstance(d, dict) and d.get("comments_added", 0) >= 1,
        f"comments_added >= 1 (got {d.get('comments_added') if isinstance(d, dict) else d!r})")

    # 04b: add 3 comments in one call -> comments_added=3
    resp = s.call("bulk_add_comments", {
        "path": p, "sheet": "Sheet1",
        "comments": [
            {"address": "B1", "text": "Comment B"},
            {"address": "C1", "text": "Comment C", "author": "TestUser"},
            {"address": "D1", "text": "Comment D"},
        ],
        "confirm": True,
    })
    check("US-BCH-04b", resp, "add 3 comments in one call -> success")
    d = parse_result(resp)
    _mc("US-BCH-04b2", isinstance(d, dict) and d.get("comments_added", 0) >= 3,
        f"comments_added >= 3 (got {d.get('comments_added') if isinstance(d, dict) else d!r})")

    # 04c: overwrite=True on existing comment -> success (no error)
    resp = s.call("bulk_add_comments", {
        "path": p, "sheet": "Sheet1",
        "comments": [{"address": "A1", "text": "Overwritten comment"}],
        "overwrite": True, "confirm": True,
    })
    check("US-BCH-04c", resp, "overwrite=True on existing A1 comment -> success")

    # 04d: wrong/non-existent sheet -> error (expected)
    resp = s.call("bulk_add_comments", {
        "path": p, "sheet": "NoSuchSheet",
        "comments": [{"address": "A1", "text": "fail"}],
        "confirm": True,
    })
    check("US-BCH-04d", resp, "wrong sheet -> error (expected)", should_fail=True)

    # 04e: confirm=False -> write gate error (expected)
    resp = s.call("bulk_add_comments", {
        "path": p, "sheet": "Sheet1",
        "comments": [{"address": "E1", "text": "Gate test"}],
        "confirm": False,
    })
    check("US-BCH-04e", resp, "confirm=False -> write gate error (expected)", should_fail=True)

    # 04f: empty comments list -> error (server requires non-empty list)
    resp = s.call("bulk_add_comments", {
        "path": p, "sheet": "Sheet1",
        "comments": [], "confirm": True,
    })
    check("US-BCH-04f", resp, "empty comments list -> error (server requires non-empty)", should_fail=True)


# ===========================================================================
# US-BCH-05  fill_formula_range
# ===========================================================================

def test_fill_formula_range(s):
    print("\n[US-BCH-05] fill_formula_range")
    p = r"C:\Temp\uat_bch_formula.xlsx"
    _create(s, p, sheets=["Sheet1"])
    # Seed numeric values in columns A and B
    for row in range(1, 6):
        _write_cell(s, p, "Sheet1", f"A{row}", row * 10)
        _write_cell(s, p, "Sheet1", f"B{row}", row * 5)

    # 05a: fill =A1+B1 from C1 down to C5 -> filled_cells=5
    resp = s.call("fill_formula_range", {
        "path": p, "sheet": "Sheet1",
        "anchor_cell": "C1", "end_cell": "C5",
        "formula": "=A1+B1", "confirm": True,
    })
    check("US-BCH-05a", resp, "fill =A1+B1 down C1:C5 -> success")
    d = parse_result(resp)
    _mc("US-BCH-05a2", isinstance(d, dict) and d.get("filled_cells", 0) == 5,
        f"filled_cells == 5 (got {d.get('filled_cells') if isinstance(d, dict) else d!r})")

    # 05b: fill formula across a row (D1:H1) -> filled_cells=5
    resp = s.call("fill_formula_range", {
        "path": p, "sheet": "Sheet1",
        "anchor_cell": "D1", "end_cell": "H1",
        "formula": "=A1*2", "confirm": True,
    })
    check("US-BCH-05b", resp, "fill =A1*2 across D1:H1 -> success")
    d = parse_result(resp)
    _mc("US-BCH-05b2", isinstance(d, dict) and d.get("filled_cells", 0) == 5,
        f"filled_cells == 5 across row (got {d.get('filled_cells') if isinstance(d, dict) else d!r})")

    # 05c: single-cell range (anchor == end) -> filled_cells=1
    resp = s.call("fill_formula_range", {
        "path": p, "sheet": "Sheet1",
        "anchor_cell": "E2", "end_cell": "E2",
        "formula": "=SUM(A1:A5)", "confirm": True,
    })
    check("US-BCH-05c", resp, "fill single cell E2:E2 -> filled_cells=1")
    d = parse_result(resp)
    _mc("US-BCH-05c2", isinstance(d, dict) and d.get("filled_cells", 0) >= 1,
        f"filled_cells >= 1 for single cell (got {d.get('filled_cells') if isinstance(d, dict) else d!r})")

    # 05d: formula missing leading '=' -> error (expected)
    resp = s.call("fill_formula_range", {
        "path": p, "sheet": "Sheet1",
        "anchor_cell": "F1", "end_cell": "F5",
        "formula": "SUM(A1:A5)", "confirm": True,
    })
    check("US-BCH-05d", resp, "formula without '=' -> error (expected)", should_fail=True)

    # 05e: non-existent sheet -> error (expected)
    resp = s.call("fill_formula_range", {
        "path": p, "sheet": "NoSuchSheet",
        "anchor_cell": "A1", "end_cell": "A5",
        "formula": "=A1+1", "confirm": True,
    })
    check("US-BCH-05e", resp, "non-existent sheet -> error (expected)", should_fail=True)

    # 05f: confirm=False -> write gate error (expected)
    resp = s.call("fill_formula_range", {
        "path": p, "sheet": "Sheet1",
        "anchor_cell": "G1", "end_cell": "G5",
        "formula": "=A1+1", "confirm": False,
    })
    check("US-BCH-05f", resp, "confirm=False -> write gate error (expected)", should_fail=True)

    # 05g: AVERAGE formula, 3 rows down -> filled_cells=3
    resp = s.call("fill_formula_range", {
        "path": p, "sheet": "Sheet1",
        "anchor_cell": "H1", "end_cell": "H3",
        "formula": "=AVERAGE(A1:B1)", "confirm": True,
    })
    check("US-BCH-05g", resp, "fill =AVERAGE(A1:B1) down H1:H3 -> success")
    d = parse_result(resp)
    _mc("US-BCH-05g2", isinstance(d, dict) and d.get("filled_cells", 0) == 3,
        f"filled_cells == 3 (got {d.get('filled_cells') if isinstance(d, dict) else d!r})")


# ===========================================================================
# US-CHT-01  chart  (openpyxl-backed; runs even when COM is unavailable)
# ===========================================================================

def test_chart(s):
    print("\n[US-CHT-01] chart")
    p = r"C:\Temp\uat_com_chart.xlsx"
    _create(s, p, sheets=["Data"])
    s.call("range_io", {
        "operation": "write",
        "path": p, "sheet": "Data", "address": "A1",
        "values": [
            ["Month", "Revenue", "", "Category", "Units"],
            ["Jan", 100, "", "North", 10],
            ["Feb", 150, "", "South", 15],
            ["Mar", 120, "", "East", 12],
            ["Apr", 200, "", "West", 20],
        ],
        "confirm": True,
    })

    # 01a: time-series data -> line chart using preferred payload names
    resp = s.call("chart", {
        "operation": "create",
        "path": p, "sheet": "Data",
        "chart_type": "line", "data_address": "A1:B5",
        "title": "Revenue Trend",
        "confirm": True,
    })
    check("US-CHT-01a", resp, "create line chart for seeded time-series data -> success")
    d = parse_result(resp)
    _mc("US-CHT-01a2", isinstance(d, dict) and d.get("chart_type") == "line",
        f"time-series chart_type == 'line' (got {d.get('chart_type') if isinstance(d, dict) else d!r})")
    line_chart_index = d.get("chart_index", 0) if isinstance(d, dict) else 0

    # 01b: category comparison -> bar chart using preferred target_address name
    resp = s.call("chart", {
        "operation": "create",
        "path": p, "sheet": "Data",
        "chart_type": "bar", "data_address": "D1:E5",
        "title": "Units by Category",
        "target_address": "G1",
        "confirm": True,
    })
    check("US-CHT-01b", resp, "create bar chart for category comparison (target_address=G1) -> success")
    d = parse_result(resp)
    _mc("US-CHT-01b2", isinstance(d, dict) and d.get("target_cell") == "G1",
        f"target_address honored as returned target_cell='G1' (got {d.get('target_cell') if isinstance(d, dict) else d!r})")

    # 01c: update first chart using preferred data_address payload name
    resp = s.call("chart", {
        "operation": "update",
        "path": p, "sheet": "Data",
        "chart_index": line_chart_index,
        "data_address": "A1:B3",
        "confirm": True,
    })
    check("US-CHT-01c", resp, "update chart via preferred data_address payload -> success")
    d = parse_result(resp)
    _mc("US-CHT-01c2", isinstance(d, dict) and d.get("chart_index") == line_chart_index,
        f"updated chart_index preserved as {line_chart_index} (got {d.get('chart_index') if isinstance(d, dict) else d!r})")
    _mc("US-CHT-01c3", isinstance(d, dict) and d.get("data_range") == "A1:B3",
        f"preferred data_address returned data_range='A1:B3' (got {d.get('data_range') if isinstance(d, dict) else d!r})")

    # 01d: list charts -> count >= 2
    resp = s.call("chart", {
        "operation": "list",
        "path": p, "sheet": "Data",
    })
    check("US-CHT-01d", resp, "list charts -> success")
    d = parse_result(resp)
    count = d.get("count", 0) if isinstance(d, dict) else 0
    _mc("US-CHT-01d2", count >= 2, f"chart count >= 2 after creating 2 charts (got {count})")

    # 01e: create chart on non-existent sheet -> error (expected)
    resp = s.call("chart", {
        "operation": "create",
        "path": p, "sheet": "NoSuchSheet",
        "chart_type": "bar", "data_address": "A1:B5",
        "confirm": True,
    })
    check("US-CHT-01e", resp, "chart on non-existent sheet -> error (expected)", should_fail=True)

    # 01f: create without chart_type -> error (expected)
    resp = s.call("chart", {
        "operation": "create",
        "path": p, "sheet": "Data",
        "data_address": "A1:B5",
        "confirm": True,
    })
    check("US-CHT-01f", resp, "create chart missing chart_type -> error (expected)", should_fail=True)


# ===========================================================================
# US-COM-02  recalculate_workbook
# ===========================================================================

def test_recalculate_workbook(s):
    if not COM_AVAILABLE:
        print("\n[US-COM-02] recalculate_workbook — SKIP (COM unavailable or not enabled)")
        _skip("US-COM-02a", "recalculate_workbook (COM unavailable or not enabled)")
        _skip("US-COM-02b", "recalculate_workbook full_calc=True (COM unavailable or not enabled)")
        _skip("US-COM-02c", "recalculate_workbook non-existent file (COM unavailable or not enabled)")
        _skip("US-COM-02d", "recalculate_workbook confirm=False (COM unavailable or not enabled)")
        return
    print("\n[US-COM-02] recalculate_workbook")
    p = r"C:\Temp\uat_com_recalc.xlsx"
    _create(s, p, sheets=["Sheet1"])
    _write_cell(s, p, "Sheet1", "A1", 10)
    _write_cell(s, p, "Sheet1", "A2", 20)
    s.call("cell", {
        "operation": "write",
        "path": p, "sheet": "Sheet1",
        "address": "A3", "value": "=A1+A2", "confirm": True,
    })

    # 02a: recalculate workbook -> ok=True + calc_mode key
    resp = s.call("recalculate_workbook", {"path": p, "confirm": True})
    check("US-COM-02a", resp, "recalculate_workbook -> ok=True")
    d = parse_result(resp)
    _mc("US-COM-02a2", isinstance(d, dict) and d.get("ok") is True,
        f"ok=True (got {d.get('ok') if isinstance(d, dict) else d!r})")
    _mc("US-COM-02a3", isinstance(d, dict) and "calc_mode" in d,
        f"result has 'calc_mode' key (got keys: {list(d.keys()) if isinstance(d, dict) else d!r})")

    # 02b: full_calc=True -> calc_mode='full_recalc'
    resp = s.call("recalculate_workbook", {"path": p, "full_calc": True, "confirm": True})
    check("US-COM-02b", resp, "recalculate_workbook full_calc=True -> success")
    d = parse_result(resp)
    _mc("US-COM-02b2", isinstance(d, dict) and d.get("calc_mode") == "full_recalc",
        f"calc_mode='full_recalc' (got {d.get('calc_mode') if isinstance(d, dict) else d!r})")

    # 02c: non-existent file -> error (expected)
    resp = s.call("recalculate_workbook", {
        "path": r"C:\Temp\does_not_exist_recalc.xlsx", "confirm": True,
    })
    check("US-COM-02c", resp, "non-existent file -> error (expected)", should_fail=True)

    # 02d: confirm=False -> write gate error (expected)
    resp = s.call("recalculate_workbook", {"path": p, "confirm": False})
    check("US-COM-02d", resp, "confirm=False -> write gate error (expected)", should_fail=True)


# ===========================================================================
# US-COM-03  export_as_pdf
# ===========================================================================

def test_export_as_pdf(s):
    if not COM_AVAILABLE:
        print("\n[US-COM-03] export_as_pdf — SKIP (COM unavailable or not enabled)")
        _skip("US-COM-03a", "export scope=sheet (COM unavailable or not enabled)")
        _skip("US-COM-03b", "export scope=workbook (COM unavailable or not enabled)")
        _skip("US-COM-03c", "export non-existent file (COM unavailable or not enabled)")
        _skip("US-COM-03d", "export confirm=False (COM unavailable or not enabled)")
        _skip("US-COM-03e", "export scope=sheet missing sheet param (COM unavailable or not enabled)")
        _skip("US-COM-03f", "invalid sheet preserves pre-existing target (COM unavailable or not enabled)")
        _skip("US-COM-03g", "scope=workbook rejects stray sheet and preserves target (COM unavailable or not enabled)")
        return
    print("\n[US-COM-03] export_as_pdf")
    p = r"C:\Temp\uat_com_export.xlsx"
    _create(s, p, sheets=["Report"])
    _write_cell(s, p, "Report", "A1", "Export Test Data")
    _write_cell(s, p, "Report", "B1", 42)
    Path(r"C:\Temp\uat_com_sheet_export.pdf").write_bytes(b"%PDF-1.7\nstale-sheet\n")
    Path(r"C:\Temp\uat_com_workbook_export.pdf").write_bytes(b"%PDF-1.7\nstale-workbook\n")

    # 03a: export single sheet -> ok=True OR graceful COM error if Excel unavailable
    resp = s.call("export_as_pdf", {
        "scope": "sheet",
        "path": p, "sheet": "Report",
        "output_path": r"C:\Temp\uat_com_sheet_export.pdf",
        "confirm": True,
    })
    check("US-COM-03a", resp, "export scope=sheet -> success (or graceful COM gate error)")
    d = parse_result(resp)
    _mc("US-COM-03a2", isinstance(d, dict) and d.get("size_bytes", 0) > 0,
        f"sheet export returns non-zero size_bytes (got {d.get('size_bytes') if isinstance(d, dict) else d!r})")
    _mc("US-COM-03a3", isinstance(d, dict) and isinstance(d.get("sha256"), str) and len(d.get("sha256", "")) == 64,
        f"sheet export returns sha256 digest (got {d.get('sha256') if isinstance(d, dict) else d!r})")
    sheet_pdf = Path(r"C:\Temp\uat_com_sheet_export.pdf")
    _mc("US-COM-03a4", sheet_pdf.exists() and sheet_pdf.read_bytes() != b"%PDF-1.7\nstale-sheet\n",
        "sheet export replaced any pre-existing stale target with fresh bytes")

    # 03b: export entire workbook -> success or graceful COM error
    resp = s.call("export_as_pdf", {
        "scope": "workbook",
        "path": p,
        "output_path": r"C:\Temp\uat_com_workbook_export.pdf",
        "confirm": True,
    })
    check("US-COM-03b", resp, "export scope=workbook -> success (or graceful COM gate error)")
    d = parse_result(resp)
    _mc("US-COM-03b2", isinstance(d, dict) and d.get("size_bytes", 0) > 0,
        f"workbook export returns non-zero size_bytes (got {d.get('size_bytes') if isinstance(d, dict) else d!r})")
    _mc("US-COM-03b3", isinstance(d, dict) and isinstance(d.get("sha256"), str) and len(d.get("sha256", "")) == 64,
        f"workbook export returns sha256 digest (got {d.get('sha256') if isinstance(d, dict) else d!r})")
    _mc("US-COM-03b4", isinstance(d, dict) and d.get("source") == str(Path(p).resolve()),
        f"workbook export returns source workbook path (got {d.get('source') if isinstance(d, dict) else d!r})")
    _mc("US-COM-03b5", isinstance(d, dict) and isinstance(d.get("exported_at"), str) and d.get("exported_at", "").endswith("Z"),
        f"workbook export returns exported_at UTC timestamp (got {d.get('exported_at') if isinstance(d, dict) else d!r})")
    _mc("US-COM-03b6", isinstance(d, dict) and isinstance(d.get("elapsed_ms"), int) and d.get("elapsed_ms", -1) >= 0,
        f"workbook export returns non-negative elapsed_ms (got {d.get('elapsed_ms') if isinstance(d, dict) else d!r})")
    workbook_pdf = Path(r"C:\Temp\uat_com_workbook_export.pdf")
    _mc("US-COM-03b7", workbook_pdf.exists() and workbook_pdf.read_bytes() != b"%PDF-1.7\nstale-workbook\n",
        "workbook export replaced any pre-existing stale target with fresh bytes")

    # 03c: non-existent source file -> error (expected)
    resp = s.call("export_as_pdf", {
        "scope": "sheet",
        "path": r"C:\Temp\no_such_file_for_pdf.xlsx",
        "sheet": "Sheet1",
        "output_path": r"C:\Temp\uat_com_fail_pdf.pdf",
        "confirm": True,
    })
    check("US-COM-03c", resp, "non-existent source file -> error (expected)", should_fail=True)

    # 03d: confirm=False -> error (expected)
    resp = s.call("export_as_pdf", {
        "scope": "sheet",
        "path": p, "sheet": "Report",
        "output_path": r"C:\Temp\uat_com_gate_pdf.pdf",
        "confirm": False,
    })
    check("US-COM-03d", resp, "confirm=False -> error (expected)", should_fail=True)

    # 03e: scope='sheet' without sheet param -> validation error (expected)
    resp = s.call("export_as_pdf", {
        "scope": "sheet",
        "path": p,
        "output_path": r"C:\Temp\uat_com_nosheet_pdf.pdf",
        "confirm": True,
    })
    check("US-COM-03e", resp, "scope='sheet' missing sheet param -> validation error (expected)", should_fail=True)

    preserved_pdf = Path(r"C:\Temp\uat_com_preserve_pdf.pdf")
    stale_bytes = b"%PDF-1.7\npreserve-me\n"
    preserved_pdf.write_bytes(stale_bytes)
    resp = s.call("export_as_pdf", {
        "scope": "sheet",
        "path": p, "sheet": "MissingSheet",
        "output_path": str(preserved_pdf),
        "confirm": True,
    })
    check("US-COM-03f", resp, "invalid sheet preserves pre-existing target (expected error)", should_fail=True)
    _mc("US-COM-03f2", preserved_pdf.read_bytes() == stale_bytes,
        "invalid export request preserved the pre-existing target bytes")

    workbook_preserved_pdf = Path(r"C:\Temp\uat_com_workbook_preserve_pdf.pdf")
    workbook_stale_bytes = b"%PDF-1.7\npreserve-workbook-invalid\n"
    workbook_preserved_pdf.write_bytes(workbook_stale_bytes)
    resp = s.call("export_as_pdf", {
        "scope": "workbook",
        "path": p,
        "sheet": "Report",
        "output_path": str(workbook_preserved_pdf),
        "confirm": True,
    })
    check("US-COM-03g", resp, "scope='workbook' with stray sheet -> validation error (expected)", should_fail=True)
    _mc("US-COM-03g2", workbook_preserved_pdf.read_bytes() == workbook_stale_bytes,
        "workbook export validation failure preserved the pre-existing target bytes")


# ===========================================================================
# US-COM-04  hydrate_template
# ===========================================================================

def test_hydrate_template(s):
    if not COM_AVAILABLE:
        print("\n[US-COM-04] hydrate_template — SKIP (COM unavailable or not enabled)")
        _skip("US-COM-04a", "hydrate_template with cell_writes (COM unavailable or not enabled)")
        _skip("US-COM-04b", "hydrate_template empty cell_writes (COM unavailable or not enabled)")
        _skip("US-COM-04c", "hydrate_template non-existent template (COM unavailable or not enabled)")
        _skip("US-COM-04d", "hydrate_template confirm=False (COM unavailable or not enabled)")
        _skip("US-COM-04e", "hydrate_template invalid ref (COM unavailable or not enabled)")
        return
    print("\n[US-COM-04] hydrate_template")
    tmpl = r"C:\Temp\uat_com_template_src.xlsx"
    out1 = r"C:\Temp\uat_com_hydrated_out1.xlsx"
    out2 = r"C:\Temp\uat_com_hydrated_out2.xlsx"
    _create(s, tmpl, sheets=["Sheet1"])
    _write_cell(s, tmpl, "Sheet1", "A1", "Project Name:")
    _write_cell(s, tmpl, "Sheet1", "B1", "<<PLACEHOLDER>>")
    _write_cell(s, tmpl, "Sheet1", "A2", "Date:")

    # 04a: hydrate with 2 cell_writes -> ok=True, writes_applied=2
    resp = s.call("hydrate_template", {
        "template_path": tmpl,
        "cell_writes": [
            {"ref": "B1", "value": "Alpha Project"},
            {"ref": "B2", "value": "2026-03-04"},
        ],
        "output_path": out1,
        "confirm": True,
    })
    check("US-COM-04a", resp, "hydrate_template 2 cell_writes -> success")
    d = parse_result(resp)
    _mc("US-COM-04a2", isinstance(d, dict) and d.get("writes_applied", 0) == 2,
        f"writes_applied == 2 (got {d.get('writes_applied') if isinstance(d, dict) else d!r})")

    # 04b: hydrate with empty cell_writes -> ok=True, writes_applied=0
    resp = s.call("hydrate_template", {
        "template_path": tmpl,
        "cell_writes": [],
        "output_path": out2,
        "confirm": True,
    })
    check("US-COM-04b", resp, "hydrate_template empty cell_writes -> success (writes_applied=0)")
    d = parse_result(resp)
    _mc("US-COM-04b2", isinstance(d, dict) and d.get("writes_applied", -1) == 0,
        f"writes_applied == 0 (got {d.get('writes_applied') if isinstance(d, dict) else d!r})")

    # 04c: non-existent template -> error (expected)
    resp = s.call("hydrate_template", {
        "template_path": r"C:\Temp\no_such_template.xlsx",
        "cell_writes": [],
        "output_path": r"C:\Temp\uat_com_ht_fail.xlsx",
        "confirm": True,
    })
    check("US-COM-04c", resp, "non-existent template -> error (expected)", should_fail=True)

    # 04d: confirm=False -> write gate error (expected)
    resp = s.call("hydrate_template", {
        "template_path": tmpl,
        "cell_writes": [],
        "output_path": r"C:\Temp\uat_com_ht_gate.xlsx",
        "confirm": False,
    })
    check("US-COM-04d", resp, "confirm=False -> write gate error (expected)", should_fail=True)

    # 04e: invalid ref in cell_writes -> validation error (expected)
    resp = s.call("hydrate_template", {
        "template_path": tmpl,
        "cell_writes": [{"ref": "!invalid!ref", "value": "bad"}],
        "output_path": r"C:\Temp\uat_com_ht_invalid.xlsx",
        "confirm": True,
    })
    check("US-COM-04e", resp, "invalid ref in cell_writes -> validation error (expected)", should_fail=True)


# ===========================================================================
# US-COM-05  pivot_table
# ===========================================================================

def test_pivot_table(s):
    if not COM_AVAILABLE:
        print("\n[US-COM-05] pivot_table — SKIP (COM unavailable or not enabled)")
        _skip("US-COM-05a", "pivot_table create (COM unavailable or not enabled)")
        _skip("US-COM-05b", "pivot_table read (COM unavailable or not enabled)")
        _skip("US-COM-05c", "pivot_table refresh (COM unavailable or not enabled)")
        _skip("US-COM-05d", "pivot_table missing required params (COM unavailable or not enabled)")
        _skip("US-COM-05e", "pivot_table invalid operation (COM unavailable or not enabled)")
        return
    print("\n[US-COM-05] pivot_table")
    p = r"C:\Temp\uat_com_pivot.xlsx"
    _create(s, p, sheets=["Data", "PivotOut"])
    s.call("range_io", {
        "operation": "write",
        "path": p, "sheet": "Data", "address": "A1",
        "values": [
            ["Region", "Product", "Sales"],
            ["North", "Widget", 100],
            ["South", "Widget", 150],
            ["North", "Gadget", 200],
            ["South", "Gadget", 120],
            ["North", "Widget", 80],
        ],
        "confirm": True,
    })

    # 05a: create pivot table from Data!A1:C6 -> success or graceful COM error
    resp = s.call("pivot_table", {
        "operation": "create",
        "path": p,
        "source_range": "Data!A1:C6",
        "dest_sheet": "PivotOut",
        "dest_cell": "A1",
        "row_fields": ["Region"],
        "value_fields": ["Sales"],
        "table_name": "RegionSales",
        "confirm": True,
    })
    check("US-COM-05a", resp, "pivot_table create -> success (or graceful COM error)")

    # 05b: read pivot table -> success (or graceful error if create failed)
    resp = s.call("pivot_table", {
        "operation": "read",
        "path": p,
        "sheet": "PivotOut",
        "pivot_name": "RegionSales",
    })
    check("US-COM-05b", resp, "pivot_table read -> success (or graceful error if not created)")

    # 05c: refresh all pivot tables -> success or graceful COM error
    resp = s.call("pivot_table", {
        "operation": "refresh",
        "path": p,
        "confirm": True,
    })
    check("US-COM-05c", resp, "pivot_table refresh -> success (or graceful COM error)")

    # 05d: create with missing required params -> error (expected)
    resp = s.call("pivot_table", {
        "operation": "create",
        "path": p,
        # Omit source_range, dest_sheet, dest_cell, row_fields, value_fields
        "confirm": True,
    })
    check("US-COM-05d", resp, "pivot_table create missing required params -> error (expected)", should_fail=True)

    # 05e: unsupported operation string -> error (expected)
    resp = s.call("pivot_table", {
        "operation": "invalidop",
        "path": p,
    })
    check("US-COM-05e", resp, "pivot_table invalid operation -> error (expected)", should_fail=True)


# ===========================================================================
# Main
# ===========================================================================

def main():
    print("=" * 65)
    print("excelmcp UAT: server_batch (5 tools) + chart + COM-backed (4 tools)")
    print(f"Platform: {sys.platform}  COM_AVAILABLE: {COM_AVAILABLE}")
    print("=" * 65)

    s = McpSession()
    try:
        # --- Batch tools (always run, no COM required) ---
        test_bulk_copy_sheet(s)
        test_set_row_heights(s)
        test_set_column_widths(s)
        test_bulk_add_comments(s)
        test_fill_formula_range(s)
        test_chart(s)

        # --- COM tools (skip on non-Windows) ---
        test_recalculate_workbook(s)
        test_export_as_pdf(s)
        test_hydrate_template(s)
        test_pivot_table(s)
    finally:
        s.close()

    total = PASS_COUNT + FAIL_COUNT + SKIP_COUNT
    print()
    print("=" * 65)
    print(
        f"RESULTS: {PASS_COUNT} PASS / {FAIL_COUNT} FAIL"
        f" / {SKIP_COUNT} SKIP / {total} total"
    )
    print("=" * 65)
    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
