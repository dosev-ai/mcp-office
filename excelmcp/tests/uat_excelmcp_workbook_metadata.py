"""UAT for workbook_metadata tool and metadata resource.

Run:
    python excelmcp/tests/uat_excelmcp_workbook_metadata.py
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from urllib.parse import quote


PYTHON = sys.executable
REPO = Path(__file__).resolve().parent.parent.parent
SERVER_CMD = [PYTHON, "-m", "excelmcp.server"]
ENV = {
    **os.environ,
    "PYTHONPATH": str(REPO / "excel" / "src"),
    "EXCEL_ENABLE_WRITE": "true",
    "EXCEL_ALLOWLIST_ROOTS": r"C:\Temp",
}

PASS_COUNT = 0
FAIL_COUNT = 0


class McpSession:
    def __init__(self) -> None:
        self.proc = subprocess.Popen(
            SERVER_CMD,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            env=ENV,
            cwd=str(REPO),
        )
        self._id = 0
        self.rpc("initialize", {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "uat", "version": "1"},
        }, req_id=0)
        self.notify("notifications/initialized", {})

    def rpc(self, method: str, params: dict, req_id: int | None = None) -> dict:
        if req_id is None:
            self._id += 1
            req_id = self._id
        msg = json.dumps({
            "jsonrpc": "2.0",
            "id": req_id,
            "method": method,
            "params": params,
        }) + "\n"
        assert self.proc.stdin is not None
        self.proc.stdin.write(msg.encode("utf-8"))
        self.proc.stdin.flush()
        assert self.proc.stdout is not None
        raw = self.proc.stdout.readline()
        return json.loads(raw) if raw else {}

    def notify(self, method: str, params: dict) -> None:
        msg = json.dumps({"jsonrpc": "2.0", "method": method, "params": params}) + "\n"
        assert self.proc.stdin is not None
        self.proc.stdin.write(msg.encode("utf-8"))
        self.proc.stdin.flush()

    def call_tool(self, name: str, arguments: dict) -> dict:
        return self.rpc("tools/call", {"name": name, "arguments": arguments})

    def close(self) -> None:
        try:
            assert self.proc.stdin is not None
            self.proc.stdin.close()
        except Exception:
            pass
        self.proc.wait(timeout=5)


def parse_tool_result(resp: dict):
    result = resp.get("result", {})
    if isinstance(result, dict) and "structuredContent" in result:
        sc = result["structuredContent"]
        if isinstance(sc, dict) and "result" in sc:
            return sc["result"]
    if isinstance(result, dict) and "content" in result:
        items = result["content"]
        if items and items[0].get("type") == "text":
            return json.loads(items[0]["text"])
    return result


def is_error(resp: dict) -> bool:
    if "error" in resp:
        return True
    result = resp.get("result", {})
    return isinstance(result, dict) and result.get("isError") is True


def tc(tc_id: str, condition: bool, desc: str, detail: str = "") -> None:
    global PASS_COUNT, FAIL_COUNT
    ok = bool(condition)
    PASS_COUNT += int(ok)
    FAIL_COUNT += int(not ok)
    status = "PASS" if ok else "FAIL"
    print(f"  {status} {tc_id}: {desc}")
    if detail and not ok:
        print(f"       {detail}")


def main() -> int:
    print("Workbook metadata UAT")
    print(f"Repo: {REPO}")
    session = McpSession()
    path = Path(r"C:\Temp\uat workbook metadata.xlsx")
    if path.exists():
        path.unlink()
    try:
        resp = session.rpc("tools/list", {})
        tools = {t["name"] for t in resp.get("result", {}).get("tools", [])}
        tc("WM-01", "workbook_metadata" in tools, "tools/list contains workbook_metadata", str(sorted(tools)))

        resp = session.call_tool("create_workbook", {"path": str(path), "confirm": True})
        tc("WM-02", not is_error(resp), "create_workbook succeeds for metadata fixture", str(resp))

        resp = session.call_tool("workbook_metadata", {"operation": "read", "path": str(path)})
        data = parse_tool_result(resp)
        tc("WM-03", not is_error(resp) and data.get("present") is False, "read before write returns present=False", str(data))

        payload = {"owner": "uat", "tags": {"slice": "metadata"}}
        resp = session.call_tool("workbook_metadata", {
            "operation": "write",
            "path": str(path),
            "payload": payload,
            "confirm": True,
        })
        data = parse_tool_result(resp)
        tc("WM-04", not is_error(resp) and data.get("ok") is True and data.get("created") is True, "write stores workbook metadata", str(data))

        resp = session.call_tool("workbook_metadata", {"operation": "read", "path": str(path)})
        data = parse_tool_result(resp)
        tc("WM-05", not is_error(resp) and data.get("metadata") == payload, "read after write round-trips payload", str(data))
        tc("WM-06", data.get("visibility") == "hidden" and data.get("version") == "v1", "read contract returns hidden visibility and v1", str(data))

        resp = session.rpc("resources/templates/list", {})
        resources = resp.get("result", {}).get("resourceTemplates", [])
        tc(
            "WM-07",
            any("metadata" in r.get("uriTemplate", "") or "metadata" in r.get("name", "") for r in resources),
            "resources/templates/list exposes metadata resource template",
            str(resources),
        )

        encoded_path = quote(str(path), safe="")
        uri = f"excelmcp://workbook/{encoded_path}/metadata"
        resp = session.rpc("resources/read", {"uri": uri})
        contents = resp.get("result", {}).get("contents", [])
        payload_text = contents[0].get("text", "") if contents else ""
        parsed = json.loads(payload_text) if payload_text else {}
        tc("WM-08", parsed.get("metadata") == payload, "resources/read returns stored metadata via encoded path", str(parsed))

    finally:
        session.close()

    print(f"\nTotal: {PASS_COUNT} PASS / {FAIL_COUNT} FAIL")
    return 0 if FAIL_COUNT == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
