"""UAT User Stories: excelmcp server_io tools (part 1). Run: python excel/tests/uat_excelmcp_io_part1.py"""
import json
import subprocess
import sys
import os
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
    # FastMCP tool errors: result is a dict with isError key
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


# ---------------------------------------------------------------------------
# Extra helpers â€” parse FastMCP dict-format responses and manual assertions
# ---------------------------------------------------------------------------

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
# US-IO-01  capabilities
# ---------------------------------------------------------------------------

def test_capabilities(mcp):
    print("\n[US-IO-01] capabilities")
    # 01a: basic call succeeds
    resp = mcp.call("capabilities", {})
    check("US-IO-01a", resp, "capabilities() call â†’ no error")
    d = parse_result(resp)
    # 01b: result has 'tools' key
    _mc("US-IO-01b", isinstance(d, dict) and "tools" in d, "result has 'tools' key")
    count = len(d.get("tools", [])) if isinstance(d, dict) else 0
    # 01c: tool count > 0
    _mc("US-IO-01c", count > 0, f"tools count > 0 (got {count})")
    # 01d: result has 'version' key
    _mc("US-IO-01d", isinstance(d, dict) and "version" in d, "result has 'version' key")
    # 01e: result has 'backend' key
    _mc("US-IO-01e", isinstance(d, dict) and "backend" in d, "result has 'backend' field")
    # 01f: repeated call is stable (same tool count)
    resp2 = mcp.call("capabilities", {})
    check("US-IO-01f", resp2, "repeated capabilities() call â†’ no error")
    d2 = parse_result(resp2)
    count2 = len(d2.get("tools", [])) if isinstance(d2, dict) else -1
    _mc("US-IO-01g", count == count2, f"repeated call stable (count={count})")


# ---------------------------------------------------------------------------
# US-IO-02  create_workbook
# ---------------------------------------------------------------------------

def test_create_workbook(mcp):
    print("\n[US-IO-02] create_workbook")
    p1 = r"C:\Temp\uat_io_cw1.xlsx"
    p2 = r"C:\Temp\uat_io_cw2.xlsx"
    p3 = r"C:\Temp\uat_io_cw3.xlsx"

    # 02a: create single-sheet workbook with confirm=True
    resp = mcp.call("create_workbook", {"path": p1, "confirm": True})
    check("US-IO-02a", resp, "create workbook uat_io_cw1.xlsx â†’ success")
    # 02b: file exists on disk immediately after
    _mc("US-IO-02b", Path(p1).exists(), "file exists on disk after create")

    # 02c: create with multiple named sheets
    resp = mcp.call("create_workbook", {"path": p2, "sheets": ["Alpha", "Beta", "Gamma"], "confirm": True})
    check("US-IO-02c", resp, "create workbook with 3 sheets ['Alpha','Beta','Gamma'] â†’ success")
    d = parse_result(resp)
    _mc("US-IO-02c2", isinstance(d, dict) and len(d.get("sheets", [])) == 3,
        f"returned 'sheets' list has 3 entries (got {d.get('sheets') if isinstance(d,dict) else '?'})")

    # 02d: overwrite=True on existing file â†’ success
    resp = mcp.call("create_workbook", {"path": p1, "overwrite": True, "confirm": True})
    check("US-IO-02d", resp, "create with overwrite=True on existing file â†’ success")

    # 02e: overwrite=False (default) on existing file â†’ error
    resp = mcp.call("create_workbook", {"path": p1, "overwrite": False, "confirm": True})
    check("US-IO-02e", resp, "create with overwrite=False on existing file â†’ error (expected)", should_fail=True)

    # 02f: path outside allowlist â†’ error
    resp = mcp.call("create_workbook", {"path": r"D:\outside\test.xlsx", "confirm": True})
    check("US-IO-02f", resp, "path outside allowlist â†’ error (expected)", should_fail=True)

    # 02g: confirm=False â†’ write gate error
    resp = mcp.call("create_workbook", {"path": p3, "confirm": False})
    check("US-IO-02g", resp, "confirm=False â†’ write gate error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-IO-03  save
# ---------------------------------------------------------------------------

def test_save(mcp):
    print("\n[US-IO-03] save")
    p = r"C:\Temp\uat_io_save1.xlsx"

    # Prepare: create the workbook
    mcp.call("create_workbook", {"path": p, "confirm": True})

    # 03a: save with confirm=False â†’ write gate error
    resp = mcp.call("save", {"path": p, "confirm": False})
    check("US-IO-03a", resp, "save confirm=False â†’ write gate error (expected)", should_fail=True)

    # 03b: save with path outside allowlist â†’ error
    resp = mcp.call("save", {"path": r"D:\outside\test.xlsx", "confirm": True})
    check("US-IO-03b", resp, "save path outside allowlist â†’ error (expected)", should_fail=True)

    # 03c: save valid path confirm=True â†’ ok (no_pending_writes graceful no-op)
    resp = mcp.call("save", {"path": p, "confirm": True})
    check("US-IO-03c", resp, "save valid path confirm=True â†’ ok (no pending writes)")
    d = parse_result(resp)
    _mc("US-IO-03c2", isinstance(d, dict) and d.get("ok") is True,
        f"save result ok=True (got {d.get('ok') if isinstance(d,dict) else d!r})")

    # 03d: write cell then save â€” write_cell auto-saves, so save is still a no-op
    mcp.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                      "address": "A1", "value": "SaveTest", "confirm": True})
    resp = mcp.call("save", {"path": p, "confirm": True})
    check("US-IO-03d", resp, "save after cell write â†’ ok")

    # 03e: file exists on disk after all operations
    _mc("US-IO-03e", Path(p).exists(), "file exists on disk after save workflow")

    # 03f: re-read A1 confirms value persisted (write_cell auto-saved it)
    resp = mcp.call("cell", {"operation": "read", "path": p, "sheet": "Sheet1", "address": "A1"})
    check("US-IO-03f", resp, "read A1 after save workflow â†’ success")
    d = parse_result(resp)
    _mc("US-IO-03g", isinstance(d, dict) and d.get("value") == "SaveTest",
        f"A1 value persisted as 'SaveTest' (got {d.get('value') if isinstance(d,dict) else d!r})")


# ---------------------------------------------------------------------------
# US-IO-04  sheet (list / add / rename / delete)
# ---------------------------------------------------------------------------

def test_sheet(mcp):
    print("\n[US-IO-04] sheet")
    p = r"C:\Temp\uat_io_sheet1.xlsx"
    mcp.call("create_workbook", {"path": p, "sheets": ["Sheet1"], "confirm": True})

    # 04a: list sheets â†’ success + has 'sheets' key
    resp = mcp.call("sheet", {"operation": "list", "path": p})
    check("US-IO-04a", resp, "sheet list â†’ success")
    d = parse_result(resp)
    _mc("US-IO-04a2", isinstance(d, dict) and "sheets" in d,
        f"result has 'sheets' key (keys: {list(d.keys()) if isinstance(d,dict) else '?'})")

    # 04b: add new sheet 'NewSheet'
    resp = mcp.call("sheet", {"operation": "add", "path": p, "sheet_name": "NewSheet", "confirm": True})
    check("US-IO-04b", resp, "sheet add 'NewSheet' â†’ success")

    # 04c: 'NewSheet' appears in list after add
    resp2 = mcp.call("sheet", {"operation": "list", "path": p})
    d2 = parse_result(resp2)
    sheets_after = d2.get("sheets", []) if isinstance(d2, dict) else []
    _mc("US-IO-04c", "NewSheet" in sheets_after, f"'NewSheet' in list after add (got {sheets_after})")

    # 04d: rename 'NewSheet' â†’ 'RenamedSheet'
    resp = mcp.call("sheet", {"operation": "rename", "path": p,
                               "sheet_name": "NewSheet", "new_name": "RenamedSheet", "confirm": True})
    check("US-IO-04d", resp, "sheet rename 'NewSheet' â†’ 'RenamedSheet' â†’ success")

    # 04e: delete 'RenamedSheet'
    resp = mcp.call("sheet", {"operation": "delete", "path": p,
                               "sheet_name": "RenamedSheet", "confirm": True})
    check("US-IO-04e", resp, "sheet delete 'RenamedSheet' â†’ success")

    # 04f: rename to existing name â†’ error
    mcp.call("sheet", {"operation": "add", "path": p, "sheet_name": "DupSheet", "confirm": True})
    resp = mcp.call("sheet", {"operation": "rename", "path": p,
                               "sheet_name": "Sheet1", "new_name": "DupSheet", "confirm": True})
    check("US-IO-04f", resp, "rename to existing name 'DupSheet' â†’ error (expected)", should_fail=True)

    # 04g: delete non-existent sheet â†’ error
    resp = mcp.call("sheet", {"operation": "delete", "path": p,
                               "sheet_name": "NonExistentSheet", "confirm": True})
    check("US-IO-04g", resp, "delete non-existent sheet â†’ error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-IO-05  get_used_range
# ---------------------------------------------------------------------------

def test_get_used_range(mcp):
    print("\n[US-IO-05] get_used_range")
    p = r"C:\Temp\uat_io_gur1.xlsx"
    mcp.call("create_workbook", {"path": p, "confirm": True})

    # 05a: freshly created empty sheet â†’ success (not an error)
    resp = mcp.call("get_used_range", {"path": p, "sheet": "Sheet1"})
    check("US-IO-05a", resp, "get_used_range on empty Sheet1 â†’ success (no error)")

    # 05b: wrong sheet name â†’ error
    resp = mcp.call("get_used_range", {"path": p, "sheet": "NoSuchSheet"})
    check("US-IO-05b", resp, "wrong sheet name â†’ error (expected)", should_fail=True)

    # 05c: non-existent file â†’ error
    resp = mcp.call("get_used_range", {"path": r"C:\Temp\no_such_file_gur.xlsx", "sheet": "Sheet1"})
    check("US-IO-05c", resp, "non-existent file â†’ error (expected)", should_fail=True)

    # 05d: write 4Ã—3 data range, then get_used_range
    rows_data = [["H1", "H2", "H3"], ["R1", 1, 2], ["R2", 3, 4], ["R3", 5, 6]]
    mcp.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                          "address": "A1", "values": rows_data, "confirm": True})
    resp = mcp.call("get_used_range", {"path": p, "sheet": "Sheet1"})
    check("US-IO-05d", resp, "get_used_range after 4-row 3-col write â†’ success")
    d = parse_result(resp)

    # 05e: max_row reflects the 4 rows written
    _mc("US-IO-05e", isinstance(d, dict) and d.get("max_row") == 4,
        f"max_row=4 after writing 4 rows (got {d.get('max_row') if isinstance(d,dict) else d!r})")

    # 05f: max_col reflects the 3 columns written
    _mc("US-IO-05f", isinstance(d, dict) and d.get("max_col") == 3,
        f"max_col=3 after writing 3 cols (got {d.get('max_col') if isinstance(d,dict) else d!r})")

    # 05g: min_row = 1 (data starts at row 1)
    _mc("US-IO-05g", isinstance(d, dict) and d.get("min_row") == 1,
        f"min_row=1 (data starts at A1) (got {d.get('min_row') if isinstance(d,dict) else d!r})")


# ---------------------------------------------------------------------------
# US-IO-06  cell  (single-cell read / write)
# ---------------------------------------------------------------------------

def test_cell(mcp):
    print("\n[US-IO-06] cell")
    p = r"C:\Temp\uat_io_cell1.xlsx"
    mcp.call("create_workbook", {"path": p, "confirm": True})

    # 06a: write string value to A1
    resp = mcp.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                              "address": "A1", "value": "Hello", "confirm": True})
    check("US-IO-06a", resp, "cell write string 'Hello' â†’ A1 success")

    # 06b: write numeric value to B2
    resp = mcp.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                              "address": "B2", "value": 42.5, "confirm": True})
    check("US-IO-06b", resp, "cell write number 42.5 â†’ B2 success")

    # 06c: read back A1 â†’ correct string
    resp = mcp.call("cell", {"operation": "read", "path": p, "sheet": "Sheet1", "address": "A1"})
    check("US-IO-06c", resp, "cell read A1 â†’ success")
    d = parse_result(resp)
    _mc("US-IO-06c2", isinstance(d, dict) and d.get("value") == "Hello",
        f"A1 value='Hello' (got {d.get('value') if isinstance(d, dict) else d!r})")

    # 06d: read back B2 â†’ correct number
    resp = mcp.call("cell", {"operation": "read", "path": p, "sheet": "Sheet1", "address": "B2"})
    d = parse_result(resp)
    _mc("US-IO-06d", isinstance(d, dict) and d.get("value") == 42.5,
        f"B2 value=42.5 (got {d.get('value') if isinstance(d, dict) else d!r})")

    # 06e: read empty cell Z99 â†’ success, value is None
    resp = mcp.call("cell", {"operation": "read", "path": p, "sheet": "Sheet1", "address": "Z99"})
    check("US-IO-06e", resp, "cell read empty Z99 â†’ success (no error)")
    d = parse_result(resp)
    _mc("US-IO-06e2", isinstance(d, dict) and d.get("value") is None,
        f"empty cell Z99 value is None (got {d.get('value') if isinstance(d, dict) else d!r})")

    # 06f: write with confirm=False â†’ write gate error
    resp = mcp.call("cell", {"operation": "write", "path": p, "sheet": "Sheet1",
                              "address": "C3", "value": "x", "confirm": False})
    check("US-IO-06f", resp, "cell write confirm=False â†’ write gate error (expected)", should_fail=True)

    # 06g: invalid cell address â†’ error
    resp = mcp.call("cell", {"operation": "read", "path": p, "sheet": "Sheet1",
                              "address": "notanaddress"})
    check("US-IO-06g", resp, "cell read invalid address 'notanaddress' â†’ error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-IO-07  range_io  (rectangular range read / write)
# ---------------------------------------------------------------------------

def test_range_io(mcp):
    print("\n[US-IO-07] range_io")
    p = r"C:\Temp\uat_io_range1.xlsx"
    mcp.call("create_workbook", {"path": p, "confirm": True})

    # 07a: write 3Ã—3 array with headers
    data = [["Name", "Score", "Grade"], ["Alice", 95, "A"], ["Bob", 82, "B"]]
    resp = mcp.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                                  "address": "A1", "values": data, "confirm": True})
    check("US-IO-07a", resp, "range_io write 3Ã—3 with headers â†’ success")

    # 07b: read back A1:C3 â†’ rows=3, cols=3
    resp = mcp.call("range_io", {"operation": "read", "path": p, "sheet": "Sheet1",
                                  "address": "A1:C3"})
    check("US-IO-07b", resp, "range_io read A1:C3 â†’ success")
    d = parse_result(resp)
    _mc("US-IO-07b2", isinstance(d, dict) and d.get("rows") == 3,
        f"read rows=3 (got {d.get('rows') if isinstance(d, dict) else d!r})")
    _mc("US-IO-07b3", isinstance(d, dict) and d.get("cols") == 3,
        f"read cols=3 (got {d.get('cols') if isinstance(d, dict) else d!r})")

    # 07c: write numeric 2Ã—3 block to E1
    resp = mcp.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                                  "address": "E1", "values": [[1, 2, 3], [4, 5, 6]], "confirm": True})
    check("US-IO-07c", resp, "range_io write 2Ã—3 numeric to E1 â†’ success")

    # 07d: partial range read â€” just first row E1:G1
    resp = mcp.call("range_io", {"operation": "read", "path": p, "sheet": "Sheet1",
                                  "address": "E1:G1"})
    check("US-IO-07d", resp, "range_io read partial E1:G1 â†’ success")
    d = parse_result(resp)
    _mc("US-IO-07d2", isinstance(d, dict) and d.get("rows") == 1,
        f"partial read rows=1 (got {d.get('rows') if isinstance(d, dict) else d!r})")

    # 07e: write without confirm â†’ write gate error
    resp = mcp.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                                  "address": "A5", "values": [[1]], "confirm": False})
    check("US-IO-07e", resp, "range_io write confirm=False â†’ write gate error (expected)", should_fail=True)

    # 07f: write operation without 'values' â†’ error
    resp = mcp.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                                  "address": "A5", "confirm": True})
    check("US-IO-07f", resp, "range_io write without values â†’ error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-IO-08  read_sheet_all
# ---------------------------------------------------------------------------

def test_read_sheet_all(mcp):
    print("\n[US-IO-08] read_sheet_all")
    p = r"C:\Temp\uat_io_all1.xlsx"
    mcp.call("create_workbook", {"path": p, "confirm": True})

    # 08a: freshly created sheet (empty or 1-cell placeholder) â†’ success
    resp = mcp.call("read_sheet_all", {"path": p, "sheet": "Sheet1"})
    check("US-IO-08a", resp, "read_sheet_all empty Sheet1 â†’ success (no error)")

    # 08b: wrong sheet name â†’ error
    resp = mcp.call("read_sheet_all", {"path": p, "sheet": "NoSuchSheet"})
    check("US-IO-08b", resp, "read_sheet_all wrong sheet â†’ error (expected)", should_fail=True)

    # 08c: file path outside allowlist â†’ error
    resp = mcp.call("read_sheet_all", {"path": r"D:\outside\test.xlsx", "sheet": "Sheet1"})
    check("US-IO-08c", resp, "read_sheet_all path outside allowlist â†’ error (expected)", should_fail=True)

    # 08d: write 5Ã—3 rows, then read_sheet_all â†’ success with correct counts
    rows5 = [["A", "B", "C"], [1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12]]
    mcp.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet1",
                          "address": "A1", "values": rows5, "confirm": True})
    resp = mcp.call("read_sheet_all", {"path": p, "sheet": "Sheet1"})
    check("US-IO-08d", resp, "read_sheet_all after 5-row write â†’ success")
    d = parse_result(resp)

    # 08e: row count = 5
    _mc("US-IO-08e", isinstance(d, dict) and d.get("rows") == 5,
        f"rows=5 (got {d.get('rows') if isinstance(d, dict) else d!r})")

    # 08f: data is non-empty nested list (list of lists)
    data_val = d.get("data") if isinstance(d, dict) else None
    _mc("US-IO-08f",
        isinstance(data_val, list) and len(data_val) > 0 and isinstance(data_val[0], list),
        f"data is non-empty list-of-lists (outer={type(data_val).__name__}, "
        f"inner={type(data_val[0]).__name__ if data_val else 'N/A'})")

    # 08g: add a second sheet, write different data, read_sheet_all on it
    mcp.call("sheet", {"operation": "add", "path": p, "sheet_name": "Sheet2", "confirm": True})
    mcp.call("range_io", {"operation": "write", "path": p, "sheet": "Sheet2",
                          "address": "A1", "values": [["X", "Y"], [10, 20]], "confirm": True})
    resp = mcp.call("read_sheet_all", {"path": p, "sheet": "Sheet2"})
    check("US-IO-08g", resp, "read_sheet_all on Sheet2 â†’ success")
    d2 = parse_result(resp)
    _mc("US-IO-08h", isinstance(d2, dict) and d2.get("cols") == 2,
        f"Sheet2 cols=2 (got {d2.get('cols') if isinstance(d2, dict) else d2!r})")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=" * 60)
    print("excelmcp UAT: server_io (part 1)")
    print("=" * 60)

    # Ensure C:\Temp exists
    Path(r"C:\Temp").mkdir(parents=True, exist_ok=True)

    # Clean up test artifacts from any previous run to ensure idempotency
    _test_files = [
        "uat_io_cw1.xlsx", "uat_io_cw2.xlsx", "uat_io_cw3.xlsx",
        "uat_io_save1.xlsx", "uat_io_sheet1.xlsx", "uat_io_gur1.xlsx",
        "uat_io_cell1.xlsx", "uat_io_range1.xlsx", "uat_io_all1.xlsx",
    ]
    for _f in _test_files:
        try:
            (Path(r"C:\Temp") / _f).unlink(missing_ok=True)
        except Exception:
            pass

    mcp = McpSession()
    try:
        test_capabilities(mcp)
        test_create_workbook(mcp)
        test_save(mcp)
        test_sheet(mcp)
        test_get_used_range(mcp)
        test_cell(mcp)
        test_range_io(mcp)
        test_read_sheet_all(mcp)
    finally:
        mcp.close()

    print(f"\nRESULTS: {PASS_COUNT} PASS | {FAIL_COUNT} FAIL")
    sys.exit(0 if FAIL_COUNT == 0 else 1)
