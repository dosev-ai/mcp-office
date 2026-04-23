"""UAT User Stories: excelmcp server_ops tools (part 1). Run: python excel/tests/uat_excelmcp_ops_part1.py

Covers the first 8 tools from server_ops.py in declaration order:
  US-OPS-01  merge         (merge / unmerge cells)
  US-OPS-02  freeze_panes  (freeze / unfreeze panes)
  US-OPS-03  auto_filter   (apply / clear AutoFilter)
  US-OPS-04  rows          (insert / delete rows)
  US-OPS-05  cols          (insert / delete columns)
  US-OPS-06  protect       (protect sheet / workbook)
  US-OPS-07  set_print_area
  US-OPS-08  cell_comment  (get / delete cell comments)
"""
import json
import os
import subprocess
import sys
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


# ---------------------------------------------------------------------------
# Harness
# ---------------------------------------------------------------------------

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
    FAIL_COUNT += not ok
    print(f"  {sym} {tc_id}: {desc}")
    if not ok:
        print(f"       resp: {str(resp)[:300]}")
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
    FAIL_COUNT += not ok
    print(f"  {sym} {tc_id}: {desc}")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_workbook(mcp, path, sheets=None, data=None):
    """Create a workbook, optionally writing initial data to Sheet1 A1."""
    kwargs = {"path": path, "confirm": True}
    if sheets:
        kwargs["sheets"] = sheets
    mcp.call("create_workbook", kwargs)
    if data:
        mcp.call("range_io", {
            "operation": "write",
            "path": path,
            "sheet": sheets[0] if sheets else "Sheet1",
            "address": "A1",
            "values": data,
            "confirm": True,
        })


# ---------------------------------------------------------------------------
# US-OPS-01  merge
# ---------------------------------------------------------------------------

def test_merge(mcp):
    print("\n[US-OPS-01] merge / unmerge_cells")
    p = r"C:\Temp\uat_ops1_merge.xlsx"
    _make_workbook(mcp, p)

    # 01a: list merged cells on fresh sheet -> success + empty list
    resp = mcp.call("merge", {"operation": "list", "path": p, "sheet": "Sheet1"})
    check("US-OPS-01a", resp, "list merged ranges on empty sheet -> success")
    d = parse_result(resp)
    _mc("US-OPS-01a2", isinstance(d, dict) and d.get("count") == 0,
        f"count=0 on empty sheet (got {d.get('count') if isinstance(d, dict) else d!r})")

    # 01b: merge A1:C1 with confirm=True -> success
    resp = mcp.call("merge", {"operation": "merge", "path": p, "sheet": "Sheet1",
                               "address": "A1:C1", "confirm": True})
    check("US-OPS-01b", resp, "merge A1:C1 -> success")

    # 01c: list after merge -> count=1 and A1:C1 appears
    resp = mcp.call("merge", {"operation": "list", "path": p, "sheet": "Sheet1"})
    check("US-OPS-01c", resp, "list after merge -> success")
    d = parse_result(resp)
    _mc("US-OPS-01c2", isinstance(d, dict) and d.get("count") == 1,
        f"count=1 after merge (got {d.get('count') if isinstance(d, dict) else d!r})")
    ranges = d.get("merged_ranges", []) if isinstance(d, dict) else []
    _mc("US-OPS-01c3", any("A1" in r for r in ranges),
        f"A1:C1 in merged_ranges (got {ranges})")

    # 01d: unmerge A1:C1 with confirm=True -> success
    resp = mcp.call("merge", {"operation": "unmerge", "path": p, "sheet": "Sheet1",
                               "address": "A1:C1", "confirm": True})
    check("US-OPS-01d", resp, "unmerge A1:C1 -> success")

    # 01e: list after unmerge -> count=0
    resp = mcp.call("merge", {"operation": "list", "path": p, "sheet": "Sheet1"})
    d = parse_result(resp)
    _mc("US-OPS-01e", isinstance(d, dict) and d.get("count") == 0,
        f"count=0 after unmerge (got {d.get('count') if isinstance(d, dict) else d!r})")

    # 01f: merge without confirm=True -> write gate error
    resp = mcp.call("merge", {"operation": "merge", "path": p, "sheet": "Sheet1",
                               "address": "B2:D2", "confirm": False})
    check("US-OPS-01f", resp, "merge confirm=False -> write gate error (expected)", should_fail=True)

    # 01g: unknown operation -> error
    resp = mcp.call("merge", {"operation": "explode", "path": p, "sheet": "Sheet1"})
    check("US-OPS-01g", resp, "invalid operation 'explode' -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS-02  freeze_panes
# ---------------------------------------------------------------------------

def test_freeze_panes(mcp):
    print("\n[US-OPS-02] freeze_panes")
    p = r"C:\Temp\uat_ops1_freeze.xlsx"
    _make_workbook(mcp, p)

    # 02a: freeze at B2 (first row + first col frozen)
    resp = mcp.call("freeze_panes", {"path": p, "sheet": "Sheet1",
                                      "freeze_at": "B2", "confirm": True})
    check("US-OPS-02a", resp, "freeze_at=B2 -> success")

    # 02b: freeze at A2 (first row only)
    resp = mcp.call("freeze_panes", {"path": p, "sheet": "Sheet1",
                                      "freeze_at": "A2", "confirm": True})
    check("US-OPS-02b", resp, "freeze_at=A2 (row only) -> success")

    # 02c: freeze at B1 (first column only)
    resp = mcp.call("freeze_panes", {"path": p, "sheet": "Sheet1",
                                      "freeze_at": "B1", "confirm": True})
    check("US-OPS-02c", resp, "freeze_at=B1 (column only) -> success")

    # 02d: unfreeze by passing freeze_at=None
    resp = mcp.call("freeze_panes", {"path": p, "sheet": "Sheet1",
                                      "freeze_at": None, "confirm": True})
    check("US-OPS-02d", resp, "freeze_at=None (unfreeze) -> success")

    # 02e: confirm=False -> write gate error
    resp = mcp.call("freeze_panes", {"path": p, "sheet": "Sheet1",
                                      "freeze_at": "B2", "confirm": False})
    check("US-OPS-02e", resp, "confirm=False -> write gate error (expected)", should_fail=True)

    # 02f: non-existent sheet -> error
    resp = mcp.call("freeze_panes", {"path": p, "sheet": "NoSuchSheet",
                                      "freeze_at": "B2", "confirm": True})
    check("US-OPS-02f", resp, "non-existent sheet -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS-03  auto_filter
# ---------------------------------------------------------------------------

def test_auto_filter(mcp):
    print("\n[US-OPS-03] auto_filter")
    p = r"C:\Temp\uat_ops1_autofilter.xlsx"
    _make_workbook(mcp, p, data=[
        ["Name", "Score", "Grade"],
        ["Alice", 90, "A"],
        ["Bob", 75, "B"],
        ["Carol", 82, "B+"],
    ])

    # 03a: apply AutoFilter to A1:C4 with confirm=True -> success
    resp = mcp.call("auto_filter", {"path": p, "sheet": "Sheet1",
                                     "address": "A1:C4", "confirm": True})
    check("US-OPS-03a", resp, "auto_filter A1:C4 -> success")

    # 03b: apply to header row only A1:C1 -> success
    resp = mcp.call("auto_filter", {"path": p, "sheet": "Sheet1",
                                     "address": "A1:C1", "confirm": True})
    check("US-OPS-03b", resp, "auto_filter header row A1:C1 -> success")

    # 03c: clear AutoFilter by passing address=None -> success
    resp = mcp.call("auto_filter", {"path": p, "sheet": "Sheet1",
                                     "address": None, "confirm": True})
    check("US-OPS-03c", resp, "auto_filter address=None (clear) -> success")

    # 03d: re-apply then result has expected keys
    resp = mcp.call("auto_filter", {"path": p, "sheet": "Sheet1",
                                     "address": "A1:C4", "confirm": True})
    check("US-OPS-03d", resp, "re-apply auto_filter -> success")
    d = parse_result(resp)
    _mc("US-OPS-03d2", isinstance(d, dict),
        f"result is a dict (got {type(d).__name__})")

    # 03e: confirm=False -> write gate error
    resp = mcp.call("auto_filter", {"path": p, "sheet": "Sheet1",
                                     "address": "A1:C4", "confirm": False})
    check("US-OPS-03e", resp, "confirm=False -> write gate error (expected)", should_fail=True)

    # 03f: non-existent sheet -> error
    resp = mcp.call("auto_filter", {"path": p, "sheet": "GhostSheet",
                                     "address": "A1:C4", "confirm": True})
    check("US-OPS-03f", resp, "non-existent sheet -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS-04  rows  (insert / delete)
# ---------------------------------------------------------------------------

def test_rows(mcp):
    print("\n[US-OPS-04] rows (insert / delete)")
    p = r"C:\Temp\uat_ops1_rows.xlsx"
    _make_workbook(mcp, p, data=[
        ["R1C1", "R1C2"],
        ["R2C1", "R2C2"],
        ["R3C1", "R3C2"],
    ])

    # 04a: insert 1 row before row 2 -> success
    resp = mcp.call("rows", {"operation": "insert", "path": p, "sheet": "Sheet1",
                              "start_row": 2, "confirm": True})
    check("US-OPS-04a", resp, "rows insert before row 2 -> success")

    # 04b: after insert, row 3 should contain original row 2 data (R2C1)
    resp = mcp.call("cell", {"operation": "read", "path": p, "sheet": "Sheet1", "address": "A3"})
    check("US-OPS-04b", resp, "read A3 after row insert -> success")
    d = parse_result(resp)
    _mc("US-OPS-04b2", isinstance(d, dict) and d.get("value") == "R2C1",
        f"A3 value is 'R2C1' after insert (got {d.get('value') if isinstance(d, dict) else d!r})")

    # 04c: insert multiple rows at once (count=2) -> success
    resp = mcp.call("rows", {"operation": "insert", "path": p, "sheet": "Sheet1",
                              "start_row": 1, "count": 2, "confirm": True})
    check("US-OPS-04c", resp, "rows insert 2 rows before row 1 -> success")

    # 04d: delete row 1 (now blank) -> success
    resp = mcp.call("rows", {"operation": "delete", "path": p, "sheet": "Sheet1",
                              "start_row": 1, "confirm": True})
    check("US-OPS-04d", resp, "rows delete row 1 -> success")

    # 04e: delete a range of rows (start_row=1, end_row=2) -> success
    resp = mcp.call("rows", {"operation": "delete", "path": p, "sheet": "Sheet1",
                              "start_row": 1, "end_row": 2, "confirm": True})
    check("US-OPS-04e", resp, "rows delete rows 1-2 (end_row) -> success")

    # 04f: confirm=False -> write gate error
    resp = mcp.call("rows", {"operation": "insert", "path": p, "sheet": "Sheet1",
                              "start_row": 1, "confirm": False})
    check("US-OPS-04f", resp, "rows insert confirm=False -> write gate error (expected)", should_fail=True)

    # 04g: invalid operation -> error
    resp = mcp.call("rows", {"operation": "clone", "path": p, "sheet": "Sheet1",
                              "start_row": 1, "confirm": True})
    check("US-OPS-04g", resp, "rows invalid operation 'clone' -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS-05  cols  (insert / delete)
# ---------------------------------------------------------------------------

def test_cols(mcp):
    print("\n[US-OPS-05] cols (insert / delete)")
    p = r"C:\Temp\uat_ops1_cols.xlsx"
    _make_workbook(mcp, p, data=[
        ["A1", "B1", "C1"],
        ["A2", "B2", "C2"],
        ["A3", "B3", "C3"],
    ])

    # 05a: insert 1 column before column 2 (B) -> success
    resp = mcp.call("cols", {"operation": "insert", "path": p, "sheet": "Sheet1",
                              "start_col": 2, "confirm": True})
    check("US-OPS-05a", resp, "cols insert before col 2 -> success")

    # 05b: after insert, col C should contain original col B data (B1)
    resp = mcp.call("cell", {"operation": "read", "path": p, "sheet": "Sheet1", "address": "C1"})
    check("US-OPS-05b", resp, "read C1 after col insert -> success")
    d = parse_result(resp)
    _mc("US-OPS-05b2", isinstance(d, dict) and d.get("value") == "B1",
        f"C1 value is 'B1' after insert (got {d.get('value') if isinstance(d, dict) else d!r})")

    # 05c: insert multiple columns (count=2) -> success
    resp = mcp.call("cols", {"operation": "insert", "path": p, "sheet": "Sheet1",
                              "start_col": 1, "count": 2, "confirm": True})
    check("US-OPS-05c", resp, "cols insert 2 cols before col 1 -> success")

    # 05d: delete column 1 -> success
    resp = mcp.call("cols", {"operation": "delete", "path": p, "sheet": "Sheet1",
                              "start_col": 1, "confirm": True})
    check("US-OPS-05d", resp, "cols delete col 1 -> success")

    # 05e: delete a range of columns (start_col=1 end_col=2) -> success
    resp = mcp.call("cols", {"operation": "delete", "path": p, "sheet": "Sheet1",
                              "start_col": 1, "end_col": 2, "confirm": True})
    check("US-OPS-05e", resp, "cols delete cols 1-2 (end_col) -> success")

    # 05f: confirm=False -> write gate error
    resp = mcp.call("cols", {"operation": "insert", "path": p, "sheet": "Sheet1",
                              "start_col": 1, "confirm": False})
    check("US-OPS-05f", resp, "cols insert confirm=False -> write gate error (expected)", should_fail=True)

    # 05g: invalid operation -> error
    resp = mcp.call("cols", {"operation": "duplicate", "path": p, "sheet": "Sheet1",
                              "start_col": 1, "confirm": True})
    check("US-OPS-05g", resp, "cols invalid operation 'duplicate' -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS-06  protect  (sheet / workbook)
# ---------------------------------------------------------------------------

def test_protect(mcp):
    print("\n[US-OPS-06] protect (sheet / workbook)")
    p = r"C:\Temp\uat_ops1_protect.xlsx"
    _make_workbook(mcp, p, data=[["Sensitive", "Data"]])

    # 06a: protect sheet with password -> success
    resp = mcp.call("protect", {"scope": "sheet", "path": p, "sheet": "Sheet1",
                                  "password": "s3cr3t", "confirm": True})
    check("US-OPS-06a", resp, "protect sheet with password -> success")

    # 06b: unprotect sheet (password=None clears protection) -> success
    resp = mcp.call("protect", {"scope": "sheet", "path": p, "sheet": "Sheet1",
                                  "password": None, "confirm": True})
    check("US-OPS-06b", resp, "unprotect sheet (password=None) -> success")

    # 06c: protect sheet without password -> success
    resp = mcp.call("protect", {"scope": "sheet", "path": p, "sheet": "Sheet1",
                                  "confirm": True})
    check("US-OPS-06c", resp, "protect sheet without password -> success")
    d = parse_result(resp)
    _mc("US-OPS-06c2", isinstance(d, dict),
        f"protect sheet result is dict (type={type(d).__name__})")

    # 06d: protect workbook structure -> success
    resp = mcp.call("protect", {"scope": "workbook", "path": p,
                                  "lock_structure": True, "lock_windows": False,
                                  "password": "wb_pass", "confirm": True})
    check("US-OPS-06d", resp, "protect workbook structure -> success")

    # 06e: unprotect workbook (password=None) -> success
    resp = mcp.call("protect", {"scope": "workbook", "path": p,
                                  "password": None, "confirm": True})
    check("US-OPS-06e", resp, "unprotect workbook (password=None) -> success")

    # 06f: protect scope=sheet missing sheet param -> error
    resp = mcp.call("protect", {"scope": "sheet", "path": p,
                                  "confirm": True})
    # sheet param is missing — server should reject
    # Note: FastMCP will default sheet=None; protect tool raises ToolError for scope='sheet' + sheet=None
    check("US-OPS-06f", resp, "protect scope=sheet missing sheet -> error (expected)", should_fail=True)

    # 06g: confirm=False -> write gate error
    resp = mcp.call("protect", {"scope": "sheet", "path": p, "sheet": "Sheet1",
                                  "confirm": False})
    check("US-OPS-06g", resp, "protect confirm=False -> write gate error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS-07  set_print_area
# ---------------------------------------------------------------------------

def test_set_print_area(mcp):
    print("\n[US-OPS-07] set_print_area")
    p = r"C:\Temp\uat_ops1_printarea.xlsx"
    _make_workbook(mcp, p, data=[
        ["Q1", "Q2", "Q3"],
        [100, 200, 300],
        [400, 500, 600],
    ])

    # 07a: set print area A1:C3 -> success
    resp = mcp.call("set_print_area", {"path": p, "sheet": "Sheet1",
                                        "address": "A1:C3", "confirm": True})
    check("US-OPS-07a", resp, "set_print_area A1:C3 -> success")

    # 07b: set print area to a single cell range A1:A1 -> success
    resp = mcp.call("set_print_area", {"path": p, "sheet": "Sheet1",
                                        "address": "A1:A1", "confirm": True})
    check("US-OPS-07b", resp, "set_print_area single cell A1:A1 -> success")

    # 07c: clear print area by passing address=None -> success
    resp = mcp.call("set_print_area", {"path": p, "sheet": "Sheet1",
                                        "address": None, "confirm": True})
    check("US-OPS-07c", resp, "set_print_area address=None (clear) -> success")

    # 07d: re-set a wider print area A1:F20 -> success
    resp = mcp.call("set_print_area", {"path": p, "sheet": "Sheet1",
                                        "address": "A1:F20", "confirm": True})
    check("US-OPS-07d", resp, "set_print_area A1:F20 -> success")
    d = parse_result(resp)
    _mc("US-OPS-07d2", isinstance(d, dict),
        f"result is dict (type={type(d).__name__})")

    # 07e: confirm=False -> write gate error
    resp = mcp.call("set_print_area", {"path": p, "sheet": "Sheet1",
                                        "address": "A1:C3", "confirm": False})
    check("US-OPS-07e", resp, "set_print_area confirm=False -> write gate error (expected)", should_fail=True)

    # 07f: non-existent sheet -> error
    resp = mcp.call("set_print_area", {"path": p, "sheet": "Phantom",
                                        "address": "A1:C3", "confirm": True})
    check("US-OPS-07f", resp, "set_print_area non-existent sheet -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS-08  cell_comment  (get / delete)
# ---------------------------------------------------------------------------

def test_cell_comment(mcp):
    print("\n[US-OPS-08] cell_comment (get / delete)")
    p = r"C:\Temp\uat_ops1_comment.xlsx"
    _make_workbook(mcp, p, data=[["Header", "Value"], ["Alice", 42]])

    # Setup: add comments via bulk_add_comments so we can test get/delete
    mcp.call("bulk_add_comments", {
        "path": p, "sheet": "Sheet1",
        "comments": [
            {"address": "A1", "text": "This is the header cell", "author": "UAT"},
            {"address": "B2", "text": "Answer to everything", "author": "UAT"},
        ],
        "confirm": True,
    })

    # 08a: get comment from A1 -> success + has_comment=True
    resp = mcp.call("cell_comment", {"operation": "get", "path": p, "sheet": "Sheet1",
                                      "address": "A1"})
    check("US-OPS-08a", resp, "cell_comment get A1 -> success")
    d = parse_result(resp)
    _mc("US-OPS-08a2", isinstance(d, dict) and d.get("has_comment") is True,
        f"has_comment=True for A1 (got {d.get('has_comment') if isinstance(d, dict) else d!r})")
    _mc("US-OPS-08a3", isinstance(d, dict) and "header" in (d.get("text") or "").lower(),
        f"comment text contains 'header' (got {d.get('text') if isinstance(d, dict) else d!r})")

    # 08b: get comment from B2 -> author='UAT'
    resp = mcp.call("cell_comment", {"operation": "get", "path": p, "sheet": "Sheet1",
                                      "address": "B2"})
    check("US-OPS-08b", resp, "cell_comment get B2 -> success")
    d = parse_result(resp)
    _mc("US-OPS-08b2", isinstance(d, dict) and d.get("author") == "UAT",
        f"author='UAT' for B2 (got {d.get('author') if isinstance(d, dict) else d!r})")

    # 08c: get comment from empty cell C3 -> has_comment=False
    resp = mcp.call("cell_comment", {"operation": "get", "path": p, "sheet": "Sheet1",
                                      "address": "C3"})
    check("US-OPS-08c", resp, "cell_comment get C3 (no comment) -> success")
    d = parse_result(resp)
    _mc("US-OPS-08c2", isinstance(d, dict) and d.get("has_comment") is False,
        f"has_comment=False for C3 (got {d.get('has_comment') if isinstance(d, dict) else d!r})")

    # 08d: delete comment from A1 with confirm=True -> success
    resp = mcp.call("cell_comment", {"operation": "delete", "path": p, "sheet": "Sheet1",
                                      "address": "A1", "confirm": True})
    check("US-OPS-08d", resp, "cell_comment delete A1 -> success")

    # 08e: get A1 after delete -> has_comment=False
    resp = mcp.call("cell_comment", {"operation": "get", "path": p, "sheet": "Sheet1",
                                      "address": "A1"})
    d = parse_result(resp)
    _mc("US-OPS-08e", isinstance(d, dict) and d.get("has_comment") is False,
        f"has_comment=False after delete (got {d.get('has_comment') if isinstance(d, dict) else d!r})")

    # 08f: delete confirm=False -> write gate error
    resp = mcp.call("cell_comment", {"operation": "delete", "path": p, "sheet": "Sheet1",
                                      "address": "B2", "confirm": False})
    check("US-OPS-08f", resp, "cell_comment delete confirm=False -> write gate error (expected)", should_fail=True)

    # 08g: invalid operation -> error
    resp = mcp.call("cell_comment", {"operation": "edit", "path": p, "sheet": "Sheet1",
                                      "address": "B2"})
    check("US-OPS-08g", resp, "cell_comment invalid operation 'edit' -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=" * 60)
    print("excelmcp UAT: server_ops (part 1)")
    print("=" * 60)

    Path(r"C:\Temp").mkdir(parents=True, exist_ok=True)

    _test_files = [
        "uat_ops1_merge.xlsx",
        "uat_ops1_freeze.xlsx",
        "uat_ops1_autofilter.xlsx",
        "uat_ops1_rows.xlsx",
        "uat_ops1_cols.xlsx",
        "uat_ops1_protect.xlsx",
        "uat_ops1_printarea.xlsx",
        "uat_ops1_comment.xlsx",
    ]
    for _f in _test_files:
        try:
            (Path(r"C:\Temp") / _f).unlink(missing_ok=True)
        except Exception:
            pass

    mcp = McpSession()
    try:
        test_merge(mcp)
        test_freeze_panes(mcp)
        test_auto_filter(mcp)
        test_rows(mcp)
        test_cols(mcp)
        test_protect(mcp)
        test_set_print_area(mcp)
        test_cell_comment(mcp)
    finally:
        mcp.close()

    print(f"\nRESULTS: {PASS_COUNT} PASS | {FAIL_COUNT} FAIL")
    sys.exit(0 if FAIL_COUNT == 0 else 1)
