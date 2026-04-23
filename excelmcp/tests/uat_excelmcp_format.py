"""UAT User Stories: excelmcp server_format tools (all 8). Run: python excel/tests/uat_excelmcp_format.py"""
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
    FAIL_COUNT += (not ok)
    print(f"  {sym} {tc_id}: {desc}")


def _make_wb(mcp, path, sheets=None):
    """Helper: create workbook at path, optionally with named sheets."""
    args = {"path": path, "confirm": True}
    if sheets:
        args["sheets"] = sheets
    mcp.call("create_workbook", args)


def _write_range(mcp, path, sheet, address, values):
    """Helper: write a 2-D range."""
    mcp.call("range_io", {"operation": "write", "path": path, "sheet": sheet,
                          "address": address, "values": values, "confirm": True})


def _read_cell(mcp, path, sheet, address):
    """Helper: read a single cell value, return the value field."""
    resp = mcp.call("cell", {"operation": "read", "path": path, "sheet": sheet, "address": address})
    d = parse_result(resp)
    return d.get("value") if isinstance(d, dict) else None


# ---------------------------------------------------------------------------
# US-FMT-01  apply_number_format
# ---------------------------------------------------------------------------

def test_apply_number_format(mcp):
    print("\n[US-FMT-01] apply_number_format")
    p = r"C:\Temp\uat_fmt_numfmt.xlsx"
    _make_wb(mcp, p)
    _write_range(mcp, p, "Sheet1", "A1", [[12345.678], [0.42], [42], [1234.5], [44927]])

    # 01a: apply "0.00" to single cell A1 -> success
    resp = mcp.call("apply_number_format", {
        "path": p, "sheet": "Sheet1", "address": "A1",
        "format_code": "0.00", "confirm": True
    })
    check("US-FMT-01a", resp, "apply_number_format '0.00' to A1 -> success")

    # 01b: apply "#,##0" to A2 (integer with thousands separator) -> success
    resp = mcp.call("apply_number_format", {
        "path": p, "sheet": "Sheet1", "address": "A2",
        "format_code": "#,##0", "confirm": True
    })
    check("US-FMT-01b", resp, "apply_number_format '#,##0' to A2 -> success")

    # 01c: apply "0%" percentage format to A3 -> success
    resp = mcp.call("apply_number_format", {
        "path": p, "sheet": "Sheet1", "address": "A3",
        "format_code": "0%", "confirm": True
    })
    check("US-FMT-01c", resp, "apply_number_format '0%' percentage to A3 -> success")

    # 01d: apply "$#,##0.00" currency format to A4 -> success
    resp = mcp.call("apply_number_format", {
        "path": p, "sheet": "Sheet1", "address": "A4",
        "format_code": "$#,##0.00", "confirm": True
    })
    check("US-FMT-01d", resp, "apply_number_format '$#,##0.00' currency to A4 -> success")

    # 01e: apply "YYYY-MM-DD" date format to A5 -> success
    resp = mcp.call("apply_number_format", {
        "path": p, "sheet": "Sheet1", "address": "A5",
        "format_code": "YYYY-MM-DD", "confirm": True
    })
    check("US-FMT-01e", resp, "apply_number_format 'YYYY-MM-DD' date format to A5 -> success")

    # 01f: apply to a range A1:A5 (bulk) -> success + result is dict
    resp = mcp.call("apply_number_format", {
        "path": p, "sheet": "Sheet1", "address": "A1:A5",
        "format_code": "#,##0.00", "confirm": True
    })
    check("US-FMT-01f", resp, "apply_number_format '#,##0.00' to range A1:A5 -> success")
    d = parse_result(resp)
    _mc("US-FMT-01f2", isinstance(d, dict), f"result is a dict (got {type(d).__name__})")

    # 01g: confirm=False -> write gate error
    resp = mcp.call("apply_number_format", {
        "path": p, "sheet": "Sheet1", "address": "A1",
        "format_code": "0.00", "confirm": False
    })
    check("US-FMT-01g", resp, "confirm=False -> write gate error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-FMT-02  apply_style  (server_format.py: ranges=[...] list param)
# ---------------------------------------------------------------------------

def test_apply_style(mcp):
    print("\n[US-FMT-02] apply_style")
    p = r"C:\Temp\uat_fmt_style.xlsx"
    _make_wb(mcp, p)
    _write_range(mcp, p, "Sheet1", "A1", [["Header"], ["Data1"], ["Data2"]])

    # 02a: bold=True on single cell A1 -> success
    resp = mcp.call("apply_style", {
        "path": p, "sheet": "Sheet1", "ranges": ["A1"],
        "bold": True, "confirm": True
    })
    check("US-FMT-02a", resp, "apply_style bold=True on A1 -> success")

    # 02b: italic=True on A2 -> success
    resp = mcp.call("apply_style", {
        "path": p, "sheet": "Sheet1", "ranges": ["A2"],
        "italic": True, "confirm": True
    })
    check("US-FMT-02b", resp, "apply_style italic=True on A2 -> success")

    # 02c: font_color='FF0000' (red) on A1 -> success
    resp = mcp.call("apply_style", {
        "path": p, "sheet": "Sheet1", "ranges": ["A1"],
        "font_color": "FF0000", "confirm": True
    })
    check("US-FMT-02c", resp, "apply_style font_color='FF0000' (red) on A1 -> success")

    # 02d: bg_color='FFFF00' (yellow fill) on A1 -> success
    resp = mcp.call("apply_style", {
        "path": p, "sheet": "Sheet1", "ranges": ["A1"],
        "bg_color": "FFFF00", "confirm": True
    })
    check("US-FMT-02d", resp, "apply_style bg_color='FFFF00' (yellow) on A1 -> success")

    # 02e: font_size=14 on A1 -> success
    resp = mcp.call("apply_style", {
        "path": p, "sheet": "Sheet1", "ranges": ["A1"],
        "font_size": 14, "confirm": True
    })
    check("US-FMT-02e", resp, "apply_style font_size=14 on A1 -> success")

    # 02f: combined bold + italic + font_color + bg_color in one call on range A1:A3
    resp = mcp.call("apply_style", {
        "path": p, "sheet": "Sheet1", "ranges": ["A1:A3"],
        "bold": True, "italic": True, "font_color": "FFFFFF", "bg_color": "336699",
        "confirm": True
    })
    check("US-FMT-02f", resp, "apply_style bold+italic+font_color+bg_color combined on A1:A3 -> success")

    # 02g: multiple disjoint ranges ['A1', 'A3'] in one call -> success
    resp = mcp.call("apply_style", {
        "path": p, "sheet": "Sheet1", "ranges": ["A1", "A3"],
        "bold": True, "confirm": True
    })
    check("US-FMT-02g", resp, "apply_style bold=True on multiple ranges ['A1','A3'] -> success")


# ---------------------------------------------------------------------------
# US-FMT-03  apply_alignment
# ---------------------------------------------------------------------------

def test_apply_alignment(mcp):
    print("\n[US-FMT-03] apply_alignment")
    p = r"C:\Temp\uat_fmt_align.xlsx"
    _make_wb(mcp, p)
    _write_range(mcp, p, "Sheet1", "A1",
                 [["CenteredText"], ["LeftText"], ["RightText"],
                  ["TopAligned"], ["BottomAligned"], ["Wrap this long label text here"]])

    # 03a: horizontal="center" on A1 -> success
    resp = mcp.call("apply_alignment", {
        "path": p, "sheet": "Sheet1", "address": "A1",
        "horizontal": "center", "confirm": True
    })
    check("US-FMT-03a", resp, "apply_alignment horizontal='center' on A1 -> success")

    # 03b: horizontal="left" on A2 -> success
    resp = mcp.call("apply_alignment", {
        "path": p, "sheet": "Sheet1", "address": "A2",
        "horizontal": "left", "confirm": True
    })
    check("US-FMT-03b", resp, "apply_alignment horizontal='left' on A2 -> success")

    # 03c: horizontal="right" on A3 -> success
    resp = mcp.call("apply_alignment", {
        "path": p, "sheet": "Sheet1", "address": "A3",
        "horizontal": "right", "confirm": True
    })
    check("US-FMT-03c", resp, "apply_alignment horizontal='right' on A3 -> success")

    # 03d: vertical="top" on A4 -> success
    resp = mcp.call("apply_alignment", {
        "path": p, "sheet": "Sheet1", "address": "A4",
        "vertical": "top", "confirm": True
    })
    check("US-FMT-03d", resp, "apply_alignment vertical='top' on A4 -> success")

    # 03e: vertical="bottom" on A5 -> success
    resp = mcp.call("apply_alignment", {
        "path": p, "sheet": "Sheet1", "address": "A5",
        "vertical": "bottom", "confirm": True
    })
    check("US-FMT-03e", resp, "apply_alignment vertical='bottom' on A5 -> success")

    # 03f: wrap_text=True on A6 -> success
    resp = mcp.call("apply_alignment", {
        "path": p, "sheet": "Sheet1", "address": "A6",
        "wrap_text": True, "confirm": True
    })
    check("US-FMT-03f", resp, "apply_alignment wrap_text=True on A6 -> success")

    # 03g: confirm=False -> write gate error
    resp = mcp.call("apply_alignment", {
        "path": p, "sheet": "Sheet1", "address": "A1",
        "horizontal": "center", "confirm": False
    })
    check("US-FMT-03g", resp, "confirm=False -> write gate error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-FMT-04  add_border
# ---------------------------------------------------------------------------

def test_add_border(mcp):
    print("\n[US-FMT-04] add_border")
    p = r"C:\Temp\uat_fmt_border.xlsx"
    _make_wb(mcp, p)
    _write_range(mcp, p, "Sheet1", "A1",
                 [["ThinCell"], ["MediumCell"], ["ThickCell"],
                  ["DashedCell"], ["ColoredBorder"], ["RangeCell"]])

    # 04a: thin border (default border_style) on A1 -> success
    resp = mcp.call("add_border", {
        "path": p, "sheet": "Sheet1", "address": "A1",
        "confirm": True
    })
    check("US-FMT-04a", resp, "add_border default thin on A1 -> success")

    # 04b: border_style='medium' on A2 -> success
    resp = mcp.call("add_border", {
        "path": p, "sheet": "Sheet1", "address": "A2",
        "border_style": "medium", "confirm": True
    })
    check("US-FMT-04b", resp, "add_border border_style='medium' on A2 -> success")

    # 04c: border_style='thick' on A3 -> success
    resp = mcp.call("add_border", {
        "path": p, "sheet": "Sheet1", "address": "A3",
        "border_style": "thick", "confirm": True
    })
    check("US-FMT-04c", resp, "add_border border_style='thick' on A3 -> success")

    # 04d: border_style='dashed' on A4 -> success
    resp = mcp.call("add_border", {
        "path": p, "sheet": "Sheet1", "address": "A4",
        "border_style": "dashed", "confirm": True
    })
    check("US-FMT-04d", resp, "add_border border_style='dashed' on A4 -> success")

    # 04e: thin border with color='0000FF' (blue) on A5 -> success
    resp = mcp.call("add_border", {
        "path": p, "sheet": "Sheet1", "address": "A5",
        "border_style": "thin", "color": "0000FF", "confirm": True
    })
    check("US-FMT-04e", resp, "add_border thin color='0000FF' (blue) on A5 -> success")

    # 04f: border on range A1:A6 (all sides for every cell) -> success
    resp = mcp.call("add_border", {
        "path": p, "sheet": "Sheet1", "address": "A1:A6",
        "border_style": "thin", "confirm": True
    })
    check("US-FMT-04f", resp, "add_border thin on range A1:A6 -> success")

    # 04g: confirm=False -> write gate error
    resp = mcp.call("add_border", {
        "path": p, "sheet": "Sheet1", "address": "A1",
        "confirm": False
    })
    check("US-FMT-04g", resp, "confirm=False -> write gate error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-FMT-05  apply_format_to_sheet_list
# ---------------------------------------------------------------------------

def test_apply_format_to_sheet_list(mcp):
    print("\n[US-FMT-05] apply_format_to_sheet_list")
    p = r"C:\Temp\uat_fmt_sheetlist.xlsx"
    _make_wb(mcp, p, sheets=["Sheet1", "Sheet2", "Sheet3"])
    _write_range(mcp, p, "Sheet1", "A1", [["Header1", "Header2"]])
    _write_range(mcp, p, "Sheet2", "A1", [["H1", "H2"]])
    _write_range(mcp, p, "Sheet3", "A1", [["X", "Y"]])

    # 05a: apply bold to A1:B1 across all sheets (sheets=None) -> success
    resp = mcp.call("apply_format_to_sheet_list", {
        "path": p, "address": "A1:B1",
        "style": {"bold": True}, "confirm": True
    })
    check("US-FMT-05a", resp, "apply_format_to_sheet_list bold A1:B1 across all sheets -> success")
    d = parse_result(resp)
    _mc("US-FMT-05a2", isinstance(d, dict), f"result is a dict (got {type(d).__name__})")

    # 05b: apply italic to explicit sheet list ['Sheet1', 'Sheet2'] only
    resp = mcp.call("apply_format_to_sheet_list", {
        "path": p, "sheets": ["Sheet1", "Sheet2"],
        "address": "A1", "style": {"italic": True}, "confirm": True
    })
    check("US-FMT-05b", resp, "apply_format_to_sheet_list italic on ['Sheet1','Sheet2'] -> success")

    # 05c: apply number_format across all sheets
    resp = mcp.call("apply_format_to_sheet_list", {
        "path": p, "address": "B1",
        "number_format": "#,##0.00", "confirm": True
    })
    check("US-FMT-05c", resp, "apply_format_to_sheet_list number_format '#,##0.00' on B1 -> success")

    # 05d: apply alignment horizontal="center" across all sheets
    resp = mcp.call("apply_format_to_sheet_list", {
        "path": p, "address": "A1:B1",
        "alignment": {"horizontal": "center"}, "confirm": True
    })
    check("US-FMT-05d", resp, "apply_format_to_sheet_list alignment center on A1:B1 -> success")

    # 05e: apply column_widths -> success
    resp = mcp.call("apply_format_to_sheet_list", {
        "path": p,
        "column_widths": {"A": 20, "B": 15}, "confirm": True
    })
    check("US-FMT-05e", resp, "apply_format_to_sheet_list column_widths A=20, B=15 -> success")

    # 05f: apply row_heights -> success
    resp = mcp.call("apply_format_to_sheet_list", {
        "path": p,
        "row_heights": {"1": 30}, "confirm": True
    })
    check("US-FMT-05f", resp, "apply_format_to_sheet_list row_heights row 1=30 -> success")

    # 05g: confirm=False -> write gate error
    resp = mcp.call("apply_format_to_sheet_list", {
        "path": p, "address": "A1",
        "style": {"bold": True}, "confirm": False
    })
    check("US-FMT-05g", resp, "confirm=False -> write gate error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-FMT-06  copy_range_format
# ---------------------------------------------------------------------------

def test_copy_range_format(mcp):
    print("\n[US-FMT-06] copy_range_format")
    p = r"C:\Temp\uat_fmt_copyfmt.xlsx"
    _make_wb(mcp, p, sheets=["Src", "Dst"])
    _write_range(mcp, p, "Src", "A1", [["BoldData"], ["ItalicData"], ["NumFmt"]])
    _write_range(mcp, p, "Dst", "A1", [["Plain1"], ["Plain2"], ["Plain3"]])

    # Set up source formatting
    mcp.call("apply_style", {"path": p, "sheet": "Src", "ranges": ["A1"],
                              "bold": True, "confirm": True})
    mcp.call("apply_style", {"path": p, "sheet": "Src", "ranges": ["A2"],
                              "italic": True, "font_color": "FF0000", "confirm": True})
    mcp.call("apply_number_format", {"path": p, "sheet": "Src", "address": "A3",
                                     "format_code": "#,##0.00", "confirm": True})

    # 06a: copy bold format from Src!A1 to Dst!A1 -> success
    resp = mcp.call("copy_range_format", {
        "path": p, "src_sheet": "Src", "src_address": "A1",
        "dst_sheet": "Dst", "dst_address": "A1", "confirm": True
    })
    check("US-FMT-06a", resp, "copy_range_format Src!A1 (bold) -> Dst!A1 -> success")

    # 06b: copy italic+red color from Src!A2 to Dst!A2 -> success
    resp = mcp.call("copy_range_format", {
        "path": p, "src_sheet": "Src", "src_address": "A2",
        "dst_sheet": "Dst", "dst_address": "A2", "confirm": True
    })
    check("US-FMT-06b", resp, "copy_range_format Src!A2 (italic+color) -> Dst!A2 -> success")

    # 06c: copy with copy_border=True (add border to source first)
    mcp.call("add_border", {"path": p, "sheet": "Src", "address": "A1",
                             "border_style": "medium", "confirm": True})
    resp = mcp.call("copy_range_format", {
        "path": p, "src_sheet": "Src", "src_address": "A1",
        "dst_sheet": "Dst", "dst_address": "B1",
        "copy_border": True, "confirm": True
    })
    check("US-FMT-06c", resp, "copy_range_format with copy_border=True -> success")

    # 06d: same-sheet copy (Src!A1 -> Src!C1) -> success
    resp = mcp.call("copy_range_format", {
        "path": p, "src_sheet": "Src", "src_address": "A1",
        "dst_sheet": "Src", "dst_address": "C1", "confirm": True
    })
    check("US-FMT-06d", resp, "copy_range_format same sheet Src!A1 -> Src!C1 -> success")

    # 06e: copy a range A1:A2 -> Dst!B2 (auto-expand dst) -> success
    resp = mcp.call("copy_range_format", {
        "path": p, "src_sheet": "Src", "src_address": "A1:A2",
        "dst_sheet": "Dst", "dst_address": "B2", "confirm": True
    })
    check("US-FMT-06e", resp, "copy_range_format range Src!A1:A2 -> Dst!B2 (auto-expand) -> success")

    # 06f: non-existent src_sheet -> error
    resp = mcp.call("copy_range_format", {
        "path": p, "src_sheet": "NoSuchSheet", "src_address": "A1",
        "dst_sheet": "Dst", "dst_address": "A1", "confirm": True
    })
    check("US-FMT-06f", resp, "non-existent src_sheet -> error (expected)", should_fail=True)

    # 06g: confirm=False -> write gate error
    resp = mcp.call("copy_range_format", {
        "path": p, "src_sheet": "Src", "src_address": "A1",
        "dst_sheet": "Dst", "dst_address": "A1", "confirm": False
    })
    check("US-FMT-06g", resp, "confirm=False -> write gate error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-FMT-07  write_range_with_format
# ---------------------------------------------------------------------------

def test_write_range_with_format(mcp):
    print("\n[US-FMT-07] write_range_with_format")
    p = r"C:\Temp\uat_fmt_wrf.xlsx"
    _make_wb(mcp, p)

    # 07a: header row with bold style -> values written, call succeeds
    resp = mcp.call("write_range_with_format", {
        "path": p, "sheet": "Sheet1", "address": "A1",
        "data": [["Product", "Revenue", "Cost"]],
        "style": {"bold": True},
        "confirm": True
    })
    check("US-FMT-07a", resp, "write_range_with_format header row bold -> success")
    val = _read_cell(mcp, p, "Sheet1", "A1")
    _mc("US-FMT-07a2", val == "Product", f"A1='Product' written correctly (got {val!r})")

    # 07b: data rows with number_format '#,##0' -> values and format applied
    resp = mcp.call("write_range_with_format", {
        "path": p, "sheet": "Sheet1", "address": "B2",
        "data": [[100000], [250000], [75000]],
        "number_format": "#,##0",
        "confirm": True
    })
    check("US-FMT-07b", resp, "write_range_with_format data rows with number_format '#,##0' -> success")
    val2 = _read_cell(mcp, p, "Sheet1", "B2")
    _mc("US-FMT-07b2", val2 == 100000, f"B2=100000 written correctly (got {val2!r})")

    # 07c: write with alignment horizontal="center" -> success
    resp = mcp.call("write_range_with_format", {
        "path": p, "sheet": "Sheet1", "address": "A5",
        "data": [["Centered Label"], ["Another Row"]],
        "alignment": {"horizontal": "center"},
        "confirm": True
    })
    check("US-FMT-07c", resp, "write_range_with_format with alignment center -> success")

    # 07d: multi-column combined style (bold + bg_color) -> success + values correct
    resp = mcp.call("write_range_with_format", {
        "path": p, "sheet": "Sheet1", "address": "D1",
        "data": [["Name", "Score"], ["Alice", 95], ["Bob", 87]],
        "style": {"bold": True, "bg_color": "D9E1F2"},
        "confirm": True
    })
    check("US-FMT-07d", resp, "write_range_with_format multi-col bold+bg_color -> success")
    val3 = _read_cell(mcp, p, "Sheet1", "D1")
    _mc("US-FMT-07d2", val3 == "Name", f"D1='Name' written correctly (got {val3!r})")

    # 07e: mixed numeric, string, and None values -> success
    resp = mcp.call("write_range_with_format", {
        "path": p, "sheet": "Sheet1", "address": "A10",
        "data": [["ItemA", 42, 3.14], ["ItemB", None, "N/A"]],
        "confirm": True
    })
    check("US-FMT-07e", resp, "write_range_with_format mixed numeric+string+None -> success")

    # 07f: write with wrap_text alignment -> success
    resp = mcp.call("write_range_with_format", {
        "path": p, "sheet": "Sheet1", "address": "G1",
        "data": [["This text is long enough to demonstrate wrap_text behavior"]],
        "alignment": {"wrap_text": True},
        "confirm": True
    })
    check("US-FMT-07f", resp, "write_range_with_format with wrap_text=True alignment -> success")

    # 07g: confirm=False -> write gate error
    resp = mcp.call("write_range_with_format", {
        "path": p, "sheet": "Sheet1", "address": "A1",
        "data": [["ShouldFail"]],
        "confirm": False
    })
    check("US-FMT-07g", resp, "confirm=False -> write gate error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# US-FMT-08  apply_conditional_format
# ---------------------------------------------------------------------------

def test_apply_conditional_format(mcp):
    print("\n[US-FMT-08] apply_conditional_format")
    p = r"C:\Temp\uat_fmt_cf.xlsx"
    _make_wb(mcp, p)
    _write_range(mcp, p, "Sheet1", "A1", [[10], [85], [50], [20], [95], [30]])

    # 08a: cell_is greaterThan 80 -> bold + red font
    resp = mcp.call("apply_conditional_format", {
        "path": p, "sheet": "Sheet1", "address": "A1:A6",
        "rule_type": "cell_is",
        "condition": "greaterThan 80",
        "format_spec": {"bold": True, "font_color": "FF0000"},
        "confirm": True
    })
    check("US-FMT-08a", resp, "apply_conditional_format cell_is greaterThan 80 bold+red -> success")

    # 08b: cell_is lessThan 25 -> yellow background
    resp = mcp.call("apply_conditional_format", {
        "path": p, "sheet": "Sheet1", "address": "A1:A6",
        "rule_type": "cell_is",
        "condition": "lessThan 25",
        "format_spec": {"bg_color": "FFFF00"},
        "confirm": True
    })
    check("US-FMT-08b", resp, "apply_conditional_format cell_is lessThan 25 yellow bg -> success")

    # 08c: cell_is between 40 60 -> green background
    resp = mcp.call("apply_conditional_format", {
        "path": p, "sheet": "Sheet1", "address": "A1:A6",
        "rule_type": "cell_is",
        "condition": "between 40 60",
        "format_spec": {"bg_color": "92D050"},
        "confirm": True
    })
    check("US-FMT-08c", resp, "apply_conditional_format cell_is between 40 60 green bg -> success")

    # 08d: cell_is equal 50 -> italic blue font
    resp = mcp.call("apply_conditional_format", {
        "path": p, "sheet": "Sheet1", "address": "A1:A6",
        "rule_type": "cell_is",
        "condition": "equal 50",
        "format_spec": {"italic": True, "font_color": "0000FF"},
        "confirm": True
    })
    check("US-FMT-08d", resp, "apply_conditional_format cell_is equal 50 italic+blue -> success")

    # 08e: formula rule type -> orange fill when cell > average
    resp = mcp.call("apply_conditional_format", {
        "path": p, "sheet": "Sheet1", "address": "A1:A6",
        "rule_type": "formula",
        "condition": "",
        "format_spec": {"bg_color": "FFC000"},
        "formula": "=A1>AVERAGE($A$1:$A$6)",
        "confirm": True
    })
    check("US-FMT-08e", resp, "apply_conditional_format formula =A1>AVERAGE orange bg -> success")

    # 08f: apply to single cell with greaterThanOrEqual condition -> success
    resp = mcp.call("apply_conditional_format", {
        "path": p, "sheet": "Sheet1", "address": "A1",
        "rule_type": "cell_is",
        "condition": "greaterThanOrEqual 10",
        "format_spec": {"bold": True},
        "confirm": True
    })
    check("US-FMT-08f", resp, "apply_conditional_format single cell greaterThanOrEqual 10 -> success")

    # 08g: confirm=False -> write gate error
    resp = mcp.call("apply_conditional_format", {
        "path": p, "sheet": "Sheet1", "address": "A1:A6",
        "rule_type": "cell_is",
        "condition": "greaterThan 50",
        "format_spec": {"bold": True},
        "confirm": False
    })
    check("US-FMT-08g", resp, "confirm=False -> write gate error (expected)", should_fail=True)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=" * 60)
    print("excelmcp UAT: server_format tools (all 8)")
    print("=" * 60)

    Path(r"C:\Temp").mkdir(parents=True, exist_ok=True)

    _test_files = [
        "uat_fmt_numfmt.xlsx",
        "uat_fmt_style.xlsx",
        "uat_fmt_align.xlsx",
        "uat_fmt_border.xlsx",
        "uat_fmt_sheetlist.xlsx",
        "uat_fmt_copyfmt.xlsx",
        "uat_fmt_wrf.xlsx",
        "uat_fmt_cf.xlsx",
    ]
    for _f in _test_files:
        try:
            (Path(r"C:\Temp") / _f).unlink(missing_ok=True)
        except Exception:
            pass

    mcp = McpSession()
    try:
        test_apply_number_format(mcp)
        test_apply_style(mcp)
        test_apply_alignment(mcp)
        test_add_border(mcp)
        test_apply_format_to_sheet_list(mcp)
        test_copy_range_format(mcp)
        test_write_range_with_format(mcp)
        test_apply_conditional_format(mcp)
    finally:
        mcp.close()

    print(f"\nRESULTS: {PASS_COUNT} PASS | {FAIL_COUNT} FAIL")
    sys.exit(0 if FAIL_COUNT == 0 else 1)
