"""UAT User Stories: excelmcp server_io tools (part 2).
Run: python excel/tests/uat_excelmcp_io_part2.py

Covers: named_range, get_cell_metadata, evaluate_formula, list_tables,
        read_table, append_rows, find_replace, import_csv_to_sheet
"""
import csv
import json
import subprocess
import sys
import os
import traceback
from pathlib import Path

PYTHON = sys.executable
REPO = Path(__file__).parent.parent.parent
SERVER_CMD = [PYTHON, "-m", "excelmcp.server"]
ENV = {
    **os.environ,
    "PYTHONPATH": str(REPO / "excel" / "src"),
    "EXCEL_ENABLE_WRITE": "true",
    "EXCEL_ALLOWLIST_ROOTS": r"C:\Temp",
}
WORKDIR = Path("C:/Temp")
PASS_COUNT = 0
FAIL_COUNT = 0


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

    def call(self, tool, args):
        self._id += 1
        msg = json.dumps({
            "jsonrpc": "2.0", "id": self._id,
            "method": "tools/call",
            "params": {"name": tool, "arguments": args},
        }) + "\n"
        self.proc.stdin.write(msg.encode())
        self.proc.stdin.flush()
        raw = self.proc.stdout.readline()
        return json.loads(raw) if raw else {}

    def close(self):
        try:
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


def result_text(content):
    if isinstance(content, list):
        for item in content:
            if isinstance(item, dict) and item.get("type") == "text":
                return item["text"]
    return str(content)


def parse_result(resp):
    """Extract parsed JSON dict/list from a FastMCP tools/call response."""
    result = resp.get("result", {})
    if isinstance(result, dict) and "content" in result:
        items = result["content"]
        if items and items[0].get("type") == "text":
            try:
                return json.loads(items[0]["text"])
            except (json.JSONDecodeError, TypeError):
                return {"_raw": items[0]["text"]}
        # list-typed tools return via structuredContent.result when content=[]
        sc = result.get("structuredContent", {})
        if "result" in sc:
            return sc["result"]
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
# US-IO2-01  named_range
# ---------------------------------------------------------------------------

def test_named_range(s):
    print("\n[US-IO2-01] named_range")
    p = r"C:\Temp\uat_io2_named_range.xlsx"
    s.call("create_workbook", {"path": p, "confirm": True})
    # Seed a value so the named range points to real data
    s.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                    "address": "B2", "value": 42, "confirm": True})

    # 01a: list names on fresh workbook → empty / success
    resp = s.call("named_range", {"operation": "list", "path": p})
    check("US-IO2-01a", resp, "list named ranges on fresh workbook → success")
    d = parse_result(resp)
    _mc("US-IO2-01a2", isinstance(d, dict) and "named_ranges" in d,
        f"result has 'named_ranges' key (got {list(d.keys()) if isinstance(d, dict) else d!r})")

    # 01b: define a named range 'MyRange' → success
    resp = s.call("named_range", {
        "operation": "write", "path": p, "name": "MyRange",
        "sheet": "Sheet1", "range_address": "B2", "confirm": True,
    })
    check("US-IO2-01b", resp, "write named range 'MyRange' → success")

    # 01c: list names after define → MyRange present
    resp = s.call("named_range", {"operation": "list", "path": p})
    check("US-IO2-01c", resp, "list after define → success")
    d = parse_result(resp)
    names = [nr.get("name", "") for nr in d.get("named_ranges", [])] if isinstance(d, dict) else []
    _mc("US-IO2-01c2", "MyRange" in names, f"'MyRange' in names list (got {names})")

    # 01d: read named range 'MyRange' → returns data
    resp = s.call("named_range", {"operation": "read", "path": p, "name": "MyRange"})
    check("US-IO2-01d", resp, "read named range 'MyRange' → success")

    # 01e: read non-existent named range → error
    resp = s.call("named_range", {"operation": "read", "path": p, "name": "DoesNotExist"})
    check("US-IO2-01e", resp, "read non-existent named range → error (expected)", should_fail=True)

    # 01f: define multi-cell named range 'BigRange'
    s.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                        "address": "A1", "values": [[1, 2, 3], [4, 5, 6]], "confirm": True})
    resp = s.call("named_range", {
        "operation": "write", "path": p, "name": "BigRange",
        "sheet": "Sheet1", "range_address": "A1:C2", "confirm": True,
    })
    check("US-IO2-01f", resp, "write multi-cell named range 'BigRange' → success")

    # 01g: redefine 'MyRange' with different address → success (overwrite)
    resp = s.call("named_range", {
        "operation": "write", "path": p, "name": "MyRange",
        "sheet": "Sheet1", "range_address": "A1", "confirm": True,
    })
    check("US-IO2-01g", resp, "redefine existing named range 'MyRange' → success")

    # 01h: unknown operation → error
    resp = s.call("named_range", {"operation": "delete", "path": p, "name": "MyRange"})
    check("US-IO2-01h", resp, "unknown operation='delete' → error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-IO2-02  get_cell_metadata
# ---------------------------------------------------------------------------

def test_get_cell_metadata(s):
    print("\n[US-IO2-02] get_cell_metadata")
    p = r"C:\Temp\uat_io2_metadata.xlsx"
    s.call("create_workbook", {"path": p, "confirm": True})

    # Seed cells
    s.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                    "address": "A1", "value": "Hello", "confirm": True})
    s.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                    "address": "B1", "value": 3.14, "confirm": True})
    s.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                    "address": "C1", "value": "=1+1", "confirm": True})

    # 02a: metadata of string cell A1 → success
    resp = s.call("get_cell_metadata", {"path": p, "sheet": "Sheet1", "address": "A1"})
    check("US-IO2-02a", resp, "get_cell_metadata string cell A1 → success")
    d = parse_result(resp)
    _mc("US-IO2-02a2", isinstance(d, dict), f"result is dict (got {type(d).__name__})")

    # 02b: metadata of numeric cell B1 → success
    resp = s.call("get_cell_metadata", {"path": p, "sheet": "Sheet1", "address": "B1"})
    check("US-IO2-02b", resp, "get_cell_metadata numeric cell B1 → success")

    # 02c: metadata of formula cell C1 → success
    resp = s.call("get_cell_metadata", {"path": p, "sheet": "Sheet1", "address": "C1"})
    check("US-IO2-02c", resp, "get_cell_metadata formula cell C1 → success")

    # 02d: metadata of empty cell D1 → success (should not error on an empty cell)
    resp = s.call("get_cell_metadata", {"path": p, "sheet": "Sheet1", "address": "D1"})
    check("US-IO2-02d", resp, "get_cell_metadata empty cell D1 → success (no error)")

    # 02e: result contains 'address' field
    d = parse_result(resp)
    _mc("US-IO2-02e", isinstance(d, dict) and "address" in d,
        f"result has 'address' field (keys: {list(d.keys()) if isinstance(d, dict) else d!r})")

    # 02f: wrong sheet name → error
    resp = s.call("get_cell_metadata", {"path": p, "sheet": "NoSheet", "address": "A1"})
    check("US-IO2-02f", resp, "get_cell_metadata wrong sheet → error (expected)", should_fail=True)

    # 02g: result for A1 contains 'data_type' or type-like field
    resp = s.call("get_cell_metadata", {"path": p, "sheet": "Sheet1", "address": "A1"})
    d2 = parse_result(resp)
    has_type_field = isinstance(d2, dict) and any(
        k in d2 for k in ("data_type", "type", "value_type", "cell_type")
    )
    _mc("US-IO2-02g", has_type_field,
        f"result has a type-like field (keys: {list(d2.keys()) if isinstance(d2, dict) else d2!r})")


# ---------------------------------------------------------------------------
# US-IO2-03  evaluate_formula
# ---------------------------------------------------------------------------

def test_evaluate_formula(s):
    print("\n[US-IO2-03] evaluate_formula")
    p = r"C:\Temp\uat_io2_formula.xlsx"
    s.call("create_workbook", {"path": p, "confirm": True})

    # Seed numeric data A1:A3 = 10, 20, 30 for SUM
    s.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                        "address": "A1", "values": [[10], [20], [30]], "confirm": True})
    # Write formulas
    s.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                    "address": "B1", "value": "=SUM(A1:A3)", "confirm": True})
    s.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                    "address": "B2", "value": '=IF(A1>5,"big","small")', "confirm": True})
    s.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                    "address": "B3", "value": '=CONCATENATE("hello"," ","world")', "confirm": True})
    s.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                    "address": "B4", "value": "=A1+A2", "confirm": True})
    s.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                    "address": "B5", "value": "=IF(SUM(A1:A3)>50,\"yes\",\"no\")", "confirm": True})

    # 03a: SUM formula in B1 → formula field present
    resp = s.call("evaluate_formula", {"path": p, "sheet": "Sheet1", "address": "B1"})
    check("US-IO2-03a", resp, "evaluate_formula SUM → success")
    d = parse_result(resp)
    _mc("US-IO2-03a2", isinstance(d, dict) and "formula" in d,
        f"result has 'formula' key (keys: {list(d.keys()) if isinstance(d, dict) else d!r})")

    # 03b: IF formula in B2 → success
    resp = s.call("evaluate_formula", {"path": p, "sheet": "Sheet1", "address": "B2"})
    check("US-IO2-03b", resp, "evaluate_formula IF → success")

    # 03c: CONCATENATE formula in B3 → success
    resp = s.call("evaluate_formula", {"path": p, "sheet": "Sheet1", "address": "B3"})
    check("US-IO2-03c", resp, "evaluate_formula CONCATENATE literal → success")

    # 03d: cell reference formula in B4 → success
    resp = s.call("evaluate_formula", {"path": p, "sheet": "Sheet1", "address": "B4"})
    check("US-IO2-03d", resp, "evaluate_formula cell-reference formula B4 → success")

    # 03e: nested IF(SUM) formula in B5 → success
    resp = s.call("evaluate_formula", {"path": p, "sheet": "Sheet1", "address": "B5"})
    check("US-IO2-03e", resp, "evaluate_formula nested IF(SUM) B5 → success")

    # 03f: empty/plain cell (A1 has a numeric value, not a formula) → success + formula is None/empty
    resp = s.call("evaluate_formula", {"path": p, "sheet": "Sheet1", "address": "A1"})
    check("US-IO2-03f", resp, "evaluate_formula plain numeric cell A1 → success (no error)")

    # 03g: wrong sheet → error
    resp = s.call("evaluate_formula", {"path": p, "sheet": "BadSheet", "address": "B1"})
    check("US-IO2-03g", resp, "evaluate_formula wrong sheet → error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-IO2-04  list_tables
# ---------------------------------------------------------------------------

def test_list_tables(s):
    print("\n[US-IO2-04] list_tables")
    p = r"C:\Temp\uat_io2_list_tables.xlsx"
    s.call("create_workbook", {"path": p, "confirm": True})

    # Seed header + data rows for table
    s.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                        "address": "A1",
                        "values": [["Name", "Score"], ["Alice", 90], ["Bob", 80]],
                        "confirm": True})

    # 04a: list tables on empty workbook (no tables yet) → success, empty list
    resp = s.call("list_tables", {"path": p})
    check("US-IO2-04a", resp, "list_tables on workbook with no tables → success")
    d = parse_result(resp)
    table_count_before = len(d) if isinstance(d, list) else 0
    _mc("US-IO2-04a2", isinstance(d, list) and table_count_before == 0,
        f"no tables before create (got {d!r})")

    # 04b: create a table then list → 1 table
    s.call("create_table", {"path": p, "sheet": "Sheet1",
                            "table_range": "A1:B3", "table_name": "tbl_Scores", "confirm": True})
    resp = s.call("list_tables", {"path": p})
    check("US-IO2-04b", resp, "list_tables after create → success")
    d = parse_result(resp)
    _mc("US-IO2-04b2", isinstance(d, list) and len(d) == 1,
        f"1 table after create (got count={len(d) if isinstance(d, list) else d!r})")

    # 04c: table name matches what we created
    tbl_names = [t.get("name", t.get("display_name", "")) for t in d] if isinstance(d, list) else []
    _mc("US-IO2-04c", any("tbl_Scores" in n for n in tbl_names),
        f"'tbl_Scores' in table list (got {tbl_names})")

    # 04d: list tables with sheet filter → still success
    resp = s.call("list_tables", {"path": p, "sheet": "Sheet1"})
    check("US-IO2-04d", resp, "list_tables with sheet='Sheet1' filter → success")

    # 04e: add second table on a new sheet, then list all → 2 tables
    s.call("sheet", {"operation": "add", "path": p, "sheet_name": "Sheet2", "confirm": True})
    s.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet2",
                        "address": "A1",
                        "values": [["Item", "Qty"], ["Pen", 5], ["Pad", 10]],
                        "confirm": True})
    s.call("create_table", {"path": p, "sheet": "Sheet2",
                            "table_range": "A1:B3", "table_name": "tbl_Items", "confirm": True})
    resp = s.call("list_tables", {"path": p})
    check("US-IO2-04e", resp, "list_tables after 2nd create → success")
    d2 = parse_result(resp)
    _mc("US-IO2-04e2", isinstance(d2, list) and len(d2) == 2,
        f"2 tables total (got count={len(d2) if isinstance(d2, list) else d2!r})")

    # 04f: list tables on non-existent file → error
    resp = s.call("list_tables", {"path": r"C:\Temp\no_such_file_io2.xlsx"})
    check("US-IO2-04f", resp, "list_tables non-existent file → error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-IO2-05  read_table
# ---------------------------------------------------------------------------

def test_read_table(s):
    print("\n[US-IO2-05] read_table")
    p = r"C:\Temp\uat_io2_read_table.xlsx"
    s.call("create_workbook", {"path": p, "confirm": True})

    # Seed header + data rows
    s.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                        "address": "A1",
                        "values": [["Product", "Price", "Qty"],
                                   ["Apple", 1.5, 100],
                                   ["Banana", 0.75, 200],
                                   ["Cherry", 3.0, 50]],
                        "confirm": True})
    s.call("create_table", {"path": p, "sheet": "Sheet1",
                            "table_range": "A1:C4", "table_name": "tbl_Products", "confirm": True})

    # 05a: read_table returns success
    resp = s.call("read_table", {"path": p, "sheet": "Sheet1", "table_name": "tbl_Products"})
    check("US-IO2-05a", resp, "read_table tbl_Products → success")
    d = parse_result(resp)
    _mc("US-IO2-05a2", isinstance(d, dict), f"result is dict (got {type(d).__name__})")

    # 05b: result contains 'headers' key with correct columns
    headers = d.get("headers", []) if isinstance(d, dict) else []
    _mc("US-IO2-05b", "Product" in headers and "Price" in headers and "Qty" in headers,
        f"headers contain Product/Price/Qty (got {headers})")

    # 05c: result contains data rows
    rows = d.get("rows", d.get("data", [])) if isinstance(d, dict) else []
    _mc("US-IO2-05c", isinstance(rows, list) and len(rows) == 3,
        f"3 data rows (got {len(rows) if isinstance(rows, list) else rows!r})")

    # 05d: wrong table name → error
    resp = s.call("read_table", {"path": p, "sheet": "Sheet1", "table_name": "tbl_NoSuch"})
    check("US-IO2-05d", resp, "read_table wrong table name → error (expected)", should_fail=True)

    # 05e: wrong sheet → error
    resp = s.call("read_table", {"path": p, "sheet": "BadSheet", "table_name": "tbl_Products"})
    check("US-IO2-05e", resp, "read_table wrong sheet → error (expected)", should_fail=True)

    # 05f: numeric values in table accessible (Price column has floats)
    first_row = rows[0] if rows else {}
    _mc("US-IO2-05f", isinstance(first_row, (dict, list)),
        f"first row is dict or list (got {type(first_row).__name__})")

    # 05g: create a minimal table (header + 1 blank row) and read it → success
    s.call("sheet", {"operation": "add", "path": p, "sheet_name": "EmptyData", "confirm": True})
    s.call("range_io", {"operation": "write", "path": p, "sheet": "EmptyData",
                        "address": "A1", "values": [["Col1", "Col2"], ["", ""]], "confirm": True})
    s.call("create_table", {"path": p, "sheet": "EmptyData",
                            "table_range": "A1:B2", "table_name": "tbl_Empty", "confirm": True})
    resp = s.call("read_table", {"path": p, "sheet": "EmptyData", "table_name": "tbl_Empty"})
    check("US-IO2-05g", resp, "read_table minimal table (header + 1 blank data row) → success")


# ---------------------------------------------------------------------------
# US-IO2-06  append_rows
# ---------------------------------------------------------------------------

def test_append_rows(s):
    print("\n[US-IO2-06] append_rows")
    p = r"C:\Temp\uat_io2_append.xlsx"
    s.call("create_workbook", {"path": p, "confirm": True})
    # Seed 2 initial rows
    s.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                        "address": "A1",
                        "values": [["Name", "Value"], ["Row1", 10]],
                        "confirm": True})

    # 06a: append single row → success
    resp = s.call("append_rows", {"path": p, "sheet": "Sheet1",
                                  "rows": [["Row2", 20]], "confirm": True})
    check("US-IO2-06a", resp, "append_rows single row → success")

    # 06b: verify row count increased to 3 (2 seed + 1 appended)
    rsa = s.call("read_sheet_all", {"path": p, "sheet": "Sheet1"})
    da = parse_result(rsa)
    _mc("US-IO2-06b", isinstance(da, dict) and da.get("rows") == 3,
        f"row count = 3 after append (got {da.get('rows') if isinstance(da, dict) else da!r})")

    # 06c: append multiple rows at once → success
    resp = s.call("append_rows", {"path": p, "sheet": "Sheet1",
                                  "rows": [["Row3", 30], ["Row4", 40], ["Row5", 50]],
                                  "confirm": True})
    check("US-IO2-06c", resp, "append_rows 3 rows at once → success")

    # 06d: row count after multi-append = 6 (3 prior + 3 new)
    rsa2 = s.call("read_sheet_all", {"path": p, "sheet": "Sheet1"})
    da2 = parse_result(rsa2)
    _mc("US-IO2-06d", isinstance(da2, dict) and da2.get("rows") == 6,
        f"row count = 6 after multi-append (got {da2.get('rows') if isinstance(da2, dict) else da2!r})")

    # 06e: confirm=False → write gate error
    resp = s.call("append_rows", {"path": p, "sheet": "Sheet1",
                                  "rows": [["Blocked", 0]], "confirm": False})
    check("US-IO2-06e", resp, "append_rows confirm=False → write gate error (expected)", should_fail=True)

    # 06f: mixed type row (string + int + float) → success
    resp = s.call("append_rows", {"path": p, "sheet": "Sheet1",
                                  "rows": [["MixedRow", 99, 3.14]], "confirm": True})
    check("US-IO2-06f", resp, "append_rows mixed-type row → success")

    # 06g: append to fresh sheet (no existing data) → success
    s.call("sheet", {"operation": "add", "path": p, "sheet_name": "FreshSheet", "confirm": True})
    resp = s.call("append_rows", {"path": p, "sheet": "FreshSheet",
                                  "rows": [["A", "B"], [1, 2]], "confirm": True})
    check("US-IO2-06g", resp, "append_rows to brand-new empty sheet → success")


# ---------------------------------------------------------------------------
# US-IO2-07  find_replace
# ---------------------------------------------------------------------------

def test_find_replace(s):
    print("\n[US-IO2-07] find_replace")
    p = r"C:\Temp\uat_io2_find_replace.xlsx"
    s.call("create_workbook", {"path": p, "confirm": True})
    # Seed data
    s.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                        "address": "A1",
                        "values": [["apple", "banana", "apple"],
                                   ["APPLE", "cherry", "banana"],
                                   ["hello world", "number", 42]],
                        "confirm": True})

    # 07a: replace existing exact string → success, replacements > 0
    resp = s.call("find_replace", {"path": p, "sheet": "Sheet1",
                                   "find_value": "banana", "replace_value": "MANGO",
                                   "confirm": True})
    check("US-IO2-07a", resp, "find_replace existing string → success")
    d = parse_result(resp)
    replaced = d.get("replacements_made", d.get("replacements", d.get("count", d.get("replaced", -1)))) if isinstance(d, dict) else -1
    _mc("US-IO2-07a2", isinstance(d, dict) and replaced not in (-1, 0, None),
        f"replacements > 0 (got {replaced!r}, keys={list(d.keys()) if isinstance(d, dict) else d!r})")

    # 07b: replace with no match → success, count = 0 or graceful empty
    resp = s.call("find_replace", {"path": p, "sheet": "Sheet1",
                                   "find_value": "NOTEXIST", "replace_value": "X",
                                   "confirm": True})
    check("US-IO2-07b", resp, "find_replace no match → success (graceful, not error)")

    # 07c: replace with empty string (deletion) → success
    resp = s.call("find_replace", {"path": p, "sheet": "Sheet1",
                                   "find_value": "cherry", "replace_value": "",
                                   "confirm": True})
    check("US-IO2-07c", resp, "find_replace replace with empty string → success")

    # 07d: case-insensitive substring replace (whole_cell=False) → success
    resp = s.call("find_replace", {"path": p, "sheet": "Sheet1",
                                   "find_value": "hello", "replace_value": "HI",
                                   "whole_cell": False, "match_case": False,
                                   "confirm": True})
    check("US-IO2-07d", resp, "find_replace case-insensitive substring replace → success")

    # 07e: case-sensitive exact match (match_case=True) → success
    resp = s.call("find_replace", {"path": p, "sheet": "Sheet1",
                                   "find_value": "apple", "replace_value": "pear",
                                   "match_case": True, "whole_cell": True,
                                   "confirm": True})
    check("US-IO2-07e", resp, "find_replace case-sensitive exact match → success")

    # 07f: confirm=False → write gate error
    resp = s.call("find_replace", {"path": p, "sheet": "Sheet1",
                                   "find_value": "pear", "replace_value": "plum",
                                   "confirm": False})
    check("US-IO2-07f", resp, "find_replace confirm=False → write gate error (expected)", should_fail=True)

    # 07g: wrong sheet → error
    resp = s.call("find_replace", {"path": p, "sheet": "NoSheet",
                                   "find_value": "apple", "replace_value": "fig",
                                   "confirm": True})
    check("US-IO2-07g", resp, "find_replace wrong sheet → error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-IO2-08  import_csv_to_sheet
# ---------------------------------------------------------------------------

def test_import_csv(s):
    print("\n[US-IO2-08] import_csv_to_sheet")
    p = r"C:\Temp\uat_io2_csv_import.xlsx"
    csv_path = r"C:\Temp\uat_io2_test.csv"
    csv_noheader = r"C:\Temp\uat_io2_noheader.csv"
    csv_numeric = r"C:\Temp\uat_io2_numeric.csv"
    csv_nonexistent = r"C:\Temp\uat_io2_does_not_exist.csv"

    # Create test CSV files
    with open(csv_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Name", "Age", "City"])
        w.writerow(["Alice", "30", "London"])
        w.writerow(["Bob", "25", "Paris"])
        w.writerow(["Carol", "35", "Berlin"])

    with open(csv_noheader, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["HeaderA", "HeaderB"])
        w.writerow(["Val1", "Val2"])

    with open(csv_numeric, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Product", "Price", "Qty"])
        w.writerow(["Widget", "9.99", "100"])
        w.writerow(["Gadget", "24.50", "50"])

    s.call("create_workbook", {"path": p, "confirm": True})

    # 08a: import CSV to new sheet → success
    resp = s.call("import_csv_to_sheet", {
        "path": p, "csv_path": csv_path, "sheet_name": "ImportedData", "confirm": True,
    })
    check("US-IO2-08a", resp, "import_csv_to_sheet to new sheet → success")

    # 08b: verify row count (3 data + 1 header = 4 rows)
    rsa = s.call("read_sheet_all", {"path": p, "sheet": "ImportedData"})
    d = parse_result(rsa)
    _mc("US-IO2-08b", isinstance(d, dict) and d.get("rows") == 4,
        f"imported 4 rows including header (got {d.get('rows') if isinstance(d, dict) else d!r})")

    # 08c: skip_header=True → only data rows (3 rows)
    resp = s.call("import_csv_to_sheet", {
        "path": p, "csv_path": csv_noheader, "sheet_name": "NoHeaderImport",
        "skip_header": True, "confirm": True,
    })
    check("US-IO2-08c", resp, "import_csv_to_sheet skip_header=True → success")
    rsa2 = s.call("read_sheet_all", {"path": p, "sheet": "NoHeaderImport"})
    d2 = parse_result(rsa2)
    _mc("US-IO2-08c2", isinstance(d2, dict) and d2.get("rows") == 1,
        f"skip_header=True → 1 data row (got {d2.get('rows') if isinstance(d2, dict) else d2!r})")

    # 08d: non-existent CSV → error
    resp = s.call("import_csv_to_sheet", {
        "path": p, "csv_path": csv_nonexistent, "sheet_name": "NoCSV", "confirm": True,
    })
    check("US-IO2-08d", resp, "import_csv_to_sheet non-existent CSV → error (expected)", should_fail=True)

    # 08e: numeric values CSV → success (values importable)
    resp = s.call("import_csv_to_sheet", {
        "path": p, "csv_path": csv_numeric, "sheet_name": "NumericData", "confirm": True,
    })
    check("US-IO2-08e", resp, "import_csv_to_sheet numeric CSV → success")
    rsa3 = s.call("read_sheet_all", {"path": p, "sheet": "NumericData"})
    d3 = parse_result(rsa3)
    _mc("US-IO2-08e2", isinstance(d3, dict) and d3.get("rows") == 3,
        f"numeric CSV: 3 rows (header + 2 data) (got {d3.get('rows') if isinstance(d3, dict) else d3!r})")

    # 08f: overwrite=True on existing sheet → success
    resp = s.call("import_csv_to_sheet", {
        "path": p, "csv_path": csv_path, "sheet_name": "ImportedData",
        "overwrite": True, "confirm": True,
    })
    check("US-IO2-08f", resp, "import_csv_to_sheet overwrite=True existing sheet → success")

    # 08g: confirm=False → write gate error
    resp = s.call("import_csv_to_sheet", {
        "path": p, "csv_path": csv_path, "sheet_name": "WontCreate", "confirm": False,
    })
    check("US-IO2-08g", resp, "import_csv_to_sheet confirm=False → write gate error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    WORKDIR.mkdir(exist_ok=True)
    print("=" * 60)
    print("excelmcp UAT: server_io (part 2)")
    print("=" * 60)

    # Clean up test artefacts from any previous run
    _test_files = [
        "uat_io2_named_range.xlsx",
        "uat_io2_metadata.xlsx",
        "uat_io2_formula.xlsx",
        "uat_io2_list_tables.xlsx",
        "uat_io2_read_table.xlsx",
        "uat_io2_append.xlsx",
        "uat_io2_find_replace.xlsx",
        "uat_io2_csv_import.xlsx",
        "uat_io2_test.csv",
        "uat_io2_noheader.csv",
        "uat_io2_numeric.csv",
    ]
    for _f in _test_files:
        try:
            (WORKDIR / _f).unlink(missing_ok=True)
        except Exception:
            pass

    s = McpSession()
    try:
        test_named_range(s)
        test_get_cell_metadata(s)
        test_evaluate_formula(s)
        test_list_tables(s)
        test_read_table(s)
        test_append_rows(s)
        test_find_replace(s)
        test_import_csv(s)
    except Exception:
        traceback.print_exc()
    finally:
        s.close()

    print(f"\nRESULTS: {PASS_COUNT} PASS | {FAIL_COUNT} FAIL")
    sys.exit(0 if FAIL_COUNT == 0 else 1)
