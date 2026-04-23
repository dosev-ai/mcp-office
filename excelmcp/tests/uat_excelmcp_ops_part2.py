"""UAT User Stories: excelmcp server_ops tools (part 2). Run: python excel/tests/uat_excelmcp_ops_part2.py

Covers the remaining 8 tools from server_ops.py (after the first 8 covered in part 1):
  US-OPS2-01  sheet_properties     (get / set tab color, state, gridlines, zoom)
  US-OPS2-02  data_validation      (get / add list, integer, date constraints)
  US-OPS2-03  validate_formula_syntax  (valid / invalid / empty formula)
  US-OPS2-04  copy_range           (copy values, formulas, edge cases)
  US-OPS2-05  create_table         (create named table, duplicate name, list_tables)
  US-OPS2-06  set_column_visibility  (hide / show single + multiple columns)
  US-OPS2-07  set_workbook_calc_mode  (manual / auto / full_calc_on_load)
  US-OPS2-08  hyperlink            (add URL, read back, remove, blocked scheme)
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
    """Create a workbook (overwriting any existing file), optionally writing initial data to Sheet1 A1."""
    kwargs = {"path": path, "confirm": True, "overwrite": True}
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
# US-OPS2-01  sheet_properties  (get / set)
# ---------------------------------------------------------------------------

def test_sheet_properties(mcp):
    print("\n[US-OPS2-01] sheet_properties (get / set)")
    p = r"C:\Temp\uat_ops2_sheet_props.xlsx"
    _make_workbook(mcp, p, sheets=["Sheet1", "Sheet2"])

    # 01a: get properties of Sheet1 -> success, returns expected keys
    resp = mcp.call("sheet_properties", {"operation": "get", "path": p, "sheet": "Sheet1"})
    check("US-OPS2-01a", resp, "get Sheet1 properties -> success")
    d = parse_result(resp)
    _mc("US-OPS2-01a2", isinstance(d, dict) and "state" in d,
        f"result has 'state' key (got keys: {list(d.keys()) if isinstance(d, dict) else d!r})")

    # 01b: set tab color to red (FF0000) -> success
    resp = mcp.call("sheet_properties", {
        "operation": "set", "path": p, "sheet": "Sheet1",
        "tab_color": "FF0000", "confirm": True,
    })
    check("US-OPS2-01b", resp, "set tab_color=FF0000 -> success")

    # 01c: get back -> tab_color is now set
    resp = mcp.call("sheet_properties", {"operation": "get", "path": p, "sheet": "Sheet1"})
    check("US-OPS2-01c", resp, "get after set tab_color -> success")
    d = parse_result(resp)
    _mc("US-OPS2-01c2", isinstance(d, dict) and d.get("tab_color") is not None,
        f"tab_color is set after update (got {d.get('tab_color') if isinstance(d, dict) else d!r})")

    # 01d: set zoom_scale to 150 -> success
    resp = mcp.call("sheet_properties", {
        "operation": "set", "path": p, "sheet": "Sheet1",
        "zoom_scale": 150, "confirm": True,
    })
    check("US-OPS2-01d", resp, "set zoom_scale=150 -> success")

    # 01e: set show_gridlines=False -> success
    resp = mcp.call("sheet_properties", {
        "operation": "set", "path": p, "sheet": "Sheet1",
        "show_gridlines": False, "confirm": True,
    })
    check("US-OPS2-01e", resp, "set show_gridlines=False -> success")

    # 01f: set without confirm -> write gate error
    resp = mcp.call("sheet_properties", {
        "operation": "set", "path": p, "sheet": "Sheet1",
        "tab_color": "00FF00", "confirm": False,
    })
    check("US-OPS2-01f", resp, "set without confirm=True -> write gate error (expected)", should_fail=True)

    # 01g: get on non-existent sheet -> error
    resp = mcp.call("sheet_properties", {
        "operation": "get", "path": p, "sheet": "NoSuch",
    })
    check("US-OPS2-01g", resp, "get on non-existent sheet -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS2-02  data_validation  (get / add)
# ---------------------------------------------------------------------------

def test_data_validation(mcp):
    print("\n[US-OPS2-02] data_validation (get / add)")
    p = r"C:\Temp\uat_ops2_dv.xlsx"
    _make_workbook(mcp, p, sheets=["Sheet1"], data=[
        ["Name", "Score", "Category", "Date"],
    ])

    # 02a: get on fresh sheet -> success, count=0
    resp = mcp.call("data_validation", {"operation": "get", "path": p, "sheet": "Sheet1"})
    check("US-OPS2-02a", resp, "get validation rules on empty sheet -> success")
    d = parse_result(resp)
    _mc("US-OPS2-02a2", isinstance(d, dict) and d.get("count", -1) == 0,
        f"count=0 on fresh sheet (got {d.get('count') if isinstance(d, dict) else d!r})")

    # 02b: add dropdown list validation to C2:C100 -> success
    resp = mcp.call("data_validation", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "validations": [{
            "address": "C2:C100",
            "validation_type": "list",
            "formula1": '"Alpha,Beta,Gamma"',
            "allow_blank": True,
            "show_error": True,
            "error_title": "Invalid",
            "error_message": "Choose from list",
        }],
        "confirm": True,
    })
    check("US-OPS2-02b", resp, "add dropdown list validation to C2:C100 -> success")

    # 02c: add integer range validation to B2:B100 (1–100) -> success
    resp = mcp.call("data_validation", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "validations": [{
            "address": "B2:B100",
            "validation_type": "whole",
            "operator": "between",
            "formula1": "1",
            "formula2": "100",
            "allow_blank": True,
            "show_error": True,
            "error_message": "Must be 1-100",
        }],
        "confirm": True,
    })
    check("US-OPS2-02c", resp, "add integer range validation B2:B100 -> success")

    # 02d: add date constraint to D2:D100 -> success
    resp = mcp.call("data_validation", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "validations": [{
            "address": "D2:D100",
            "validation_type": "date",
            "operator": "greaterThan",
            "formula1": "2020-01-01",
            "allow_blank": True,
            "show_error": True,
            "error_message": "Date must be after 2020-01-01",
        }],
        "confirm": True,
    })
    check("US-OPS2-02d", resp, "add date validation D2:D100 -> success")

    # 02e: get after adding -> count >= 3
    resp = mcp.call("data_validation", {"operation": "get", "path": p, "sheet": "Sheet1"})
    check("US-OPS2-02e", resp, "get after adding 3 rules -> success")
    d = parse_result(resp)
    _mc("US-OPS2-02e2", isinstance(d, dict) and d.get("count", 0) >= 3,
        f"count>=3 after adds (got {d.get('count') if isinstance(d, dict) else d!r})")

    # 02f: add without confirm -> write gate error
    resp = mcp.call("data_validation", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "validations": [{"address": "A1", "validation_type": "list", "formula1": '"X,Y"'}],
        "confirm": False,
    })
    check("US-OPS2-02f", resp, "add without confirm=True -> write gate error (expected)", should_fail=True)

    # 02g: add with empty validations list -> error
    resp = mcp.call("data_validation", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "validations": [],
        "confirm": True,
    })
    check("US-OPS2-02g", resp, "add with empty validations list -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS2-03  validate_formula_syntax
# ---------------------------------------------------------------------------

def test_validate_formula_syntax(mcp):
    print("\n[US-OPS2-03] validate_formula_syntax")

    # 03a: valid simple formula -> success, valid=True
    resp = mcp.call("validate_formula_syntax", {"formula": "=SUM(A1:A10)"})
    check("US-OPS2-03a", resp, "validate =SUM(A1:A10) -> success")
    d = parse_result(resp)
    _mc("US-OPS2-03a2", isinstance(d, dict) and d.get("valid") is True,
        f"valid=True for SUM formula (got {d.get('valid') if isinstance(d, dict) else d!r})")

    # 03b: valid nested IF formula -> success, valid=True
    resp = mcp.call("validate_formula_syntax", {"formula": "=IF(A1>0,\"Pos\",IF(A1<0,\"Neg\",\"Zero\"))"})
    check("US-OPS2-03b", resp, "validate nested IF formula -> success")
    d = parse_result(resp)
    _mc("US-OPS2-03b2", isinstance(d, dict) and d.get("valid") is True,
        f"valid=True for nested IF (got {d.get('valid') if isinstance(d, dict) else d!r})")

    # 03c: valid VLOOKUP formula -> success, valid=True
    resp = mcp.call("validate_formula_syntax", {"formula": "=VLOOKUP(A1,C:D,2,FALSE)"})
    check("US-OPS2-03c", resp, "validate VLOOKUP formula -> success")
    d = parse_result(resp)
    _mc("US-OPS2-03c2", isinstance(d, dict) and d.get("valid") is True,
        f"valid=True for VLOOKUP (got {d.get('valid') if isinstance(d, dict) else d!r})")

    # 03d: formula with mismatched parentheses but valid '=' prefix -> valid=True
    # (tool validates '=' prefix only; deep Excel syntax is not inspected)
    resp = mcp.call("validate_formula_syntax", {"formula": "=SUM(A1:A10"})
    check("US-OPS2-03d", resp, "validate unmatched paren formula -> success (valid=True, prefix check only)")
    d = parse_result(resp)
    _mc("US-OPS2-03d2", isinstance(d, dict) and d.get("valid") is True,
        f"valid=True for '='-prefixed formula even with mismatched parens (got {d.get('valid') if isinstance(d, dict) else d!r})")

    # 03e: empty formula string -> is_valid=False or success with validation info
    resp = mcp.call("validate_formula_syntax", {"formula": ""})
    check("US-OPS2-03e", resp, "validate empty formula -> success (tool returns validation result)")
    d = parse_result(resp)
    _mc("US-OPS2-03e2", isinstance(d, dict), f"result is a dict (got {type(d).__name__})")

    # 03f: formula without '=' prefix -> valid=False
    resp = mcp.call("validate_formula_syntax", {"formula": "SUM(A1:A10)"})
    check("US-OPS2-03f", resp, "validate formula without '=' -> success (reports valid=False)")
    d = parse_result(resp)
    _mc("US-OPS2-03f2", isinstance(d, dict) and d.get("valid") is False,
        f"valid=False for formula missing '=' (got {d.get('valid') if isinstance(d, dict) else d!r})")

    # 03g: valid SUMIF formula -> success, valid=True
    resp = mcp.call("validate_formula_syntax", {"formula": "=SUMIF(A:A,\">0\",B:B)"})
    check("US-OPS2-03g", resp, "validate SUMIF formula -> success")
    d = parse_result(resp)
    _mc("US-OPS2-03g2", isinstance(d, dict) and d.get("valid") is True,
        f"valid=True for SUMIF (got {d.get('valid') if isinstance(d, dict) else d!r})")


# ---------------------------------------------------------------------------
# US-OPS2-04  copy_range
# ---------------------------------------------------------------------------

def test_copy_range(mcp):
    print("\n[US-OPS2-04] copy_range")
    p = r"C:\Temp\uat_ops2_copy_range.xlsx"
    _make_workbook(mcp, p, sheets=["Sheet1"], data=[
        ["Alpha", "Beta", "Gamma"],
        [10, 20, 30],
        [100, 200, 300],
    ])

    # 04a: copy A1:C3 to E1 -> success
    resp = mcp.call("copy_range", {
        "path": p, "sheet": "Sheet1",
        "src_range": "A1:C3", "dst_range": "E1",
        "confirm": True,
    })
    check("US-OPS2-04a", resp, "copy A1:C3 to E1 -> success")

    # 04b: verify E1 has the copied value 'Alpha'
    resp = mcp.call("cell", {"operation": "read", "path": p, "sheet": "Sheet1", "address": "E1"})
    check("US-OPS2-04b", resp, "read E1 after copy -> success")
    d = parse_result(resp)
    _mc("US-OPS2-04b2", isinstance(d, dict) and d.get("value") == "Alpha",
        f"E1 == 'Alpha' after copy (got {d.get('value') if isinstance(d, dict) else d!r})")

    # 04c: verify E2 has numeric value 10
    resp = mcp.call("cell", {"operation": "read", "path": p, "sheet": "Sheet1", "address": "E2"})
    check("US-OPS2-04c", resp, "read E2 after copy -> success")
    d = parse_result(resp)
    _mc("US-OPS2-04c2", isinstance(d, dict) and d.get("value") == 10,
        f"E2 == 10 after copy (got {d.get('value') if isinstance(d, dict) else d!r})")

    # 04d: copy a formula cell — write formula to A5 first, then copy to E5
    mcp.call("range_io", {
        "operation": "write", "path": p, "sheet": "Sheet1",
        "address": "A5", "values": [["=SUM(A2:A3)"]], "confirm": True,
    })
    resp = mcp.call("copy_range", {
        "path": p, "sheet": "Sheet1",
        "src_range": "A5:A5", "dst_range": "E5",
        "confirm": True,
    })
    check("US-OPS2-04d", resp, "copy formula cell A5 to E5 -> success")

    # 04e: copy single cell A1 to J1 -> success
    resp = mcp.call("copy_range", {
        "path": p, "sheet": "Sheet1",
        "src_range": "A1:A1", "dst_range": "J1",
        "confirm": True,
    })
    check("US-OPS2-04e", resp, "copy single cell A1 to J1 -> success")

    # 04f: confirm=False -> write gate error
    resp = mcp.call("copy_range", {
        "path": p, "sheet": "Sheet1",
        "src_range": "A1:C3", "dst_range": "A10",
        "confirm": False,
    })
    check("US-OPS2-04f", resp, "copy_range confirm=False -> write gate error (expected)", should_fail=True)

    # 04g: non-existent sheet -> error
    resp = mcp.call("copy_range", {
        "path": p, "sheet": "NoSheet",
        "src_range": "A1:C3", "dst_range": "E1",
        "confirm": True,
    })
    check("US-OPS2-04g", resp, "copy_range non-existent sheet -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS2-05  create_table
# ---------------------------------------------------------------------------

def test_create_table(mcp):
    print("\n[US-OPS2-05] create_table")
    p = r"C:\Temp\uat_ops2_table.xlsx"
    _make_workbook(mcp, p, sheets=["Sheet1", "Sheet2"], data=[
        ["Item", "Qty", "Price"],
        ["Widget", 10, 5.99],
        ["Gadget", 5, 12.50],
        ["Doohickey", 20, 2.25],
    ])

    # 05a: create table tbl_Items over A1:C4 -> success
    resp = mcp.call("create_table", {
        "path": p, "sheet": "Sheet1",
        "table_range": "A1:C4",
        "table_name": "tbl_Items",
        "confirm": True,
    })
    check("US-OPS2-05a", resp, "create tbl_Items over A1:C4 -> success")

    # 05b: list_tables -> tbl_Items appears
    resp = mcp.call("list_tables", {"path": p, "sheet": "Sheet1"})
    check("US-OPS2-05b", resp, "list_tables after create -> success")
    d = parse_result(resp)
    tables = d if isinstance(d, list) else d.get("tables", []) if isinstance(d, dict) else []
    names = [t.get("name") or t.get("display_name") for t in tables if isinstance(t, dict)]
    _mc("US-OPS2-05b2", "tbl_Items" in names,
        f"tbl_Items found in list_tables (got {names})")

    # 05c: create second table with different style -> success
    mcp.call("range_io", {
        "operation": "write", "path": p, "sheet": "Sheet2",
        "address": "A1",
        "values": [["Month", "Revenue"], ["Jan", 10000], ["Feb", 12000]],
        "confirm": True,
    })
    resp = mcp.call("create_table", {
        "path": p, "sheet": "Sheet2",
        "table_range": "A1:B3",
        "table_name": "tbl_Revenue",
        "style": "TableStyleLight1",
        "confirm": True,
    })
    check("US-OPS2-05c", resp, "create tbl_Revenue with custom style -> success")

    # 05d: list_tables workbook-wide -> both tables present
    resp = mcp.call("list_tables", {"path": p})
    check("US-OPS2-05d", resp, "list_tables workbook-wide -> success")
    d = parse_result(resp)
    tables = d if isinstance(d, list) else d.get("tables", []) if isinstance(d, dict) else []
    names = [t.get("name") or t.get("display_name") for t in tables if isinstance(t, dict)]
    _mc("US-OPS2-05d2", len(names) >= 2,
        f">=2 tables in workbook (got {names})")

    # 05e: duplicate table name -> error
    resp = mcp.call("create_table", {
        "path": p, "sheet": "Sheet1",
        "table_range": "A1:C4",
        "table_name": "tbl_Items",
        "confirm": True,
    })
    check("US-OPS2-05e", resp, "duplicate table name -> error (expected)", should_fail=True)

    # 05f: create without confirm -> write gate error
    resp = mcp.call("create_table", {
        "path": p, "sheet": "Sheet1",
        "table_range": "A1:C4",
        "table_name": "tbl_ShouldFail",
        "confirm": False,
    })
    check("US-OPS2-05f", resp, "create_table confirm=False -> write gate error (expected)", should_fail=True)

    # 05g: non-existent sheet -> error
    resp = mcp.call("create_table", {
        "path": p, "sheet": "NoSheet",
        "table_range": "A1:C4",
        "table_name": "tbl_Ghost",
        "confirm": True,
    })
    check("US-OPS2-05g", resp, "create_table non-existent sheet -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS2-06  set_column_visibility
# ---------------------------------------------------------------------------

def test_set_column_visibility(mcp):
    print("\n[US-OPS2-06] set_column_visibility")
    p = r"C:\Temp\uat_ops2_col_vis.xlsx"
    _make_workbook(mcp, p, data=[
        ["ID", "Name", "Internal", "Score", "Notes"],
        [1, "Alice", "secret1", 90, "good"],
        [2, "Bob", "secret2", 75, "ok"],
    ])

    # 06a: hide column C -> success
    resp = mcp.call("set_column_visibility", {
        "path": p, "sheet": "Sheet1",
        "columns": ["C"], "hidden": True, "confirm": True,
    })
    check("US-OPS2-06a", resp, "hide column C -> success")

    # 06b: hide multiple columns D and E -> success
    resp = mcp.call("set_column_visibility", {
        "path": p, "sheet": "Sheet1",
        "columns": ["D", "E"], "hidden": True, "confirm": True,
    })
    check("US-OPS2-06b", resp, "hide columns D and E -> success")

    # 06c: show column C back -> success
    resp = mcp.call("set_column_visibility", {
        "path": p, "sheet": "Sheet1",
        "columns": ["C"], "hidden": False, "confirm": True,
    })
    check("US-OPS2-06c", resp, "show column C (unhide) -> success")

    # 06d: show D and E back -> success
    resp = mcp.call("set_column_visibility", {
        "path": p, "sheet": "Sheet1",
        "columns": ["D", "E"], "hidden": False, "confirm": True,
    })
    check("US-OPS2-06d", resp, "show columns D and E (unhide) -> success")

    # 06e: hide non-contiguous columns A and C -> success
    resp = mcp.call("set_column_visibility", {
        "path": p, "sheet": "Sheet1",
        "columns": ["A", "C"], "hidden": True, "confirm": True,
    })
    check("US-OPS2-06e", resp, "hide non-contiguous columns A and C -> success")

    # 06f: confirm=False -> write gate error
    resp = mcp.call("set_column_visibility", {
        "path": p, "sheet": "Sheet1",
        "columns": ["B"], "hidden": True, "confirm": False,
    })
    check("US-OPS2-06f", resp, "set_column_visibility confirm=False -> write gate error (expected)", should_fail=True)

    # 06g: non-existent sheet -> error
    resp = mcp.call("set_column_visibility", {
        "path": p, "sheet": "NoSuch",
        "columns": ["A"], "hidden": True, "confirm": True,
    })
    check("US-OPS2-06g", resp, "set_column_visibility non-existent sheet -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS2-07  set_workbook_calc_mode
# ---------------------------------------------------------------------------

def test_set_workbook_calc_mode(mcp):
    print("\n[US-OPS2-07] set_workbook_calc_mode")
    p = r"C:\Temp\uat_ops2_calc_mode.xlsx"
    _make_workbook(mcp, p, data=[["=1+1", "=2*3"]])

    # 07a: set calc_mode=manual -> success
    resp = mcp.call("set_workbook_calc_mode", {
        "path": p, "calc_mode": "manual", "confirm": True,
    })
    check("US-OPS2-07a", resp, "set calc_mode=manual -> success")
    d = parse_result(resp)
    _mc("US-OPS2-07a2", isinstance(d, dict),
        f"result is a dict (got {type(d).__name__})")

    # 07b: set calc_mode=auto -> success
    resp = mcp.call("set_workbook_calc_mode", {
        "path": p, "calc_mode": "auto", "confirm": True,
    })
    check("US-OPS2-07b", resp, "set calc_mode=auto -> success")

    # 07c: set calc_mode=autoNoTable -> success
    resp = mcp.call("set_workbook_calc_mode", {
        "path": p, "calc_mode": "autoNoTable", "confirm": True,
    })
    check("US-OPS2-07c", resp, "set calc_mode=autoNoTable -> success")

    # 07d: set full_calc_on_load=True -> success
    resp = mcp.call("set_workbook_calc_mode", {
        "path": p, "full_calc_on_load": True, "confirm": True,
    })
    check("US-OPS2-07d", resp, "set full_calc_on_load=True -> success")

    # 07e: set both calc_mode=manual and full_calc_on_load=False -> success
    resp = mcp.call("set_workbook_calc_mode", {
        "path": p, "calc_mode": "manual", "full_calc_on_load": False, "confirm": True,
    })
    check("US-OPS2-07e", resp, "set calc_mode=manual + full_calc_on_load=False -> success")

    # 07f: confirm=False -> write gate error
    resp = mcp.call("set_workbook_calc_mode", {
        "path": p, "calc_mode": "auto", "confirm": False,
    })
    check("US-OPS2-07f", resp, "set_workbook_calc_mode confirm=False -> write gate error (expected)", should_fail=True)

    # 07g: invalid calc_mode value -> error
    resp = mcp.call("set_workbook_calc_mode", {
        "path": p, "calc_mode": "turbo", "confirm": True,
    })
    check("US-OPS2-07g", resp, "invalid calc_mode='turbo' -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-OPS2-08  hyperlink  (get / add / remove)
# ---------------------------------------------------------------------------

def test_hyperlink(mcp):
    print("\n[US-OPS2-08] hyperlink (get / add / remove)")
    p = r"C:\Temp\uat_ops2_hyperlinks.xlsx"
    _make_workbook(mcp, p, sheets=["Sheet1", "Sheet2"], data=[
        ["Label", "Link"],
        ["GitHub", ""],
        ["Docs", ""],
        ["Internal", ""],
    ])

    # 08a: get hyperlinks on fresh sheet -> success, count=0
    resp = mcp.call("hyperlink", {"operation": "get", "path": p, "sheet": "Sheet1"})
    check("US-OPS2-08a", resp, "get hyperlinks on empty sheet -> success")
    d = parse_result(resp)
    count = d.get("count", -1) if isinstance(d, dict) else -1
    _mc("US-OPS2-08a2", count == 0,
        f"count=0 on fresh sheet (got {count})")

    # 08b: add URL hyperlink to B2 -> success
    resp = mcp.call("hyperlink", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "links": [{"address": "B2", "url": "https://github.com", "display_text": "GitHub"}],
        "confirm": True,
    })
    check("US-OPS2-08b", resp, "add https URL hyperlink to B2 -> success")

    # 08c: add http hyperlink to B3 -> success
    resp = mcp.call("hyperlink", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "links": [{"address": "B3", "url": "https://docs.example.com", "display_text": "Docs",
                   "tooltip": "Documentation site"}],
        "confirm": True,
    })
    check("US-OPS2-08c", resp, "add https URL hyperlink with tooltip to B3 -> success")

    # 08d: add multiple hyperlinks in one call -> success
    resp = mcp.call("hyperlink", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "links": [
            {"address": "A5", "url": "https://example.com/page1", "display_text": "Page 1"},
            {"address": "A6", "url": "https://example.com/page2", "display_text": "Page 2"},
        ],
        "confirm": True,
    })
    check("US-OPS2-08d", resp, "add two hyperlinks in one call -> success")

    # 08e: get after adding -> count >= 2
    resp = mcp.call("hyperlink", {"operation": "get", "path": p, "sheet": "Sheet1"})
    check("US-OPS2-08e", resp, "get hyperlinks after adds -> success")
    d = parse_result(resp)
    count = d.get("count", 0) if isinstance(d, dict) else 0
    _mc("US-OPS2-08e2", count >= 2,
        f"count>=2 hyperlinks after adding (got {count})")

    # 08f: blocked scheme (javascript:) -> error
    resp = mcp.call("hyperlink", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "links": [{"address": "B4", "url": "javascript:alert(1)", "display_text": "XSS"}],
        "confirm": True,
    })
    check("US-OPS2-08f", resp, "add javascript: URL -> blocked (expected)", should_fail=True)

    # 08g: remove hyperlink from B2 -> success
    resp = mcp.call("hyperlink", {
        "operation": "remove", "path": p, "sheet": "Sheet1",
        "address": "B2", "confirm": True,
    })
    check("US-OPS2-08g", resp, "remove hyperlink from B2 -> success")

    # 08h: confirm=False for add -> write gate error
    resp = mcp.call("hyperlink", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "links": [{"address": "C1", "url": "https://example.com"}],
        "confirm": False,
    })
    check("US-OPS2-08h", resp, "add hyperlink confirm=False -> write gate error (expected)", should_fail=True)

    # 08i: add with empty links list -> error
    resp = mcp.call("hyperlink", {
        "operation": "add", "path": p, "sheet": "Sheet1",
        "links": [],
        "confirm": True,
    })
    check("US-OPS2-08i", resp, "add with empty links list -> error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    mcp = McpSession()
    try:
        test_sheet_properties(mcp)
        test_data_validation(mcp)
        test_validate_formula_syntax(mcp)
        test_copy_range(mcp)
        test_create_table(mcp)
        test_set_column_visibility(mcp)
        test_set_workbook_calc_mode(mcp)
        test_hyperlink(mcp)
    finally:
        mcp.close()

    total = PASS_COUNT + FAIL_COUNT
    print(f"\nRESULTS: {PASS_COUNT}/{total} passed, {FAIL_COUNT} failed")
    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
