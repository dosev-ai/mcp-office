# Quick Start

Get excelmcp running in under 10 minutes.

---

## Path A — Agent-first (recommended)

This is the fastest path. You configure your MCP client, point it at excelmcp, and the tools are available immediately — no separate server process to manage.

### Prerequisites

- Windows 10 or 11
- Python 3.11 or later
- [Claude Desktop](https://claude.ai/download) **or** [VS Code with GitHub Copilot](https://marketplace.visualstudio.com/items?itemName=GitHub.copilot)
- Microsoft Excel (required for COM tools: styling, PDF export, live workbook operations)

### Step 1 — Clone and install

```bash
git clone https://github.com/dosev-ai/mcp-office.git
cd mcp-office
python -m venv .venv
.venv\Scripts\activate
pip install -e ./excelmcp
```

### Step 2 — Configure Claude Desktop

Open `%APPDATA%\Claude\claude_desktop_config.json` and add the `excel-excelmcp` server:

```json
{
  "mcpServers": {
    "excel-excelmcp": {
      "command": "C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe",
      "args": ["-m", "excelmcp.server"],
      "env": {
        "EXCEL_ALLOWED_DIRS": "C:\\path\\to\\your\\files",
        "EXCEL_ENABLE_WRITE": "true"
      }
    }
  }
}
```

Replace:
- `C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe` — the full path to the Python executable inside the `.venv` you created
- `C:\\path\\to\\your\\files` — the directory (or directories, semicolon-separated) where your Excel files live

Restart Claude Desktop after saving.

### Step 2 (alternative) — Configure VS Code Copilot

Create or edit `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "excel-excelmcp": {
      "type": "stdio",
      "command": "C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe",
      "args": ["-m", "excelmcp.server"],
      "env": {
        "EXCEL_ALLOWED_DIRS": "C:\\path\\to\\your\\files",
        "EXCEL_ENABLE_WRITE": "true"
      }
    }
  }
}
```

### Step 3 — Verify

In Claude Desktop or VS Code Copilot, ask:

```
Call capabilities() on excel-excelmcp
```

You should see the full tool list (60+ tools). If you do, you’re ready.

Next: **[Your First Excel MCP Workflow →](first-workflow.md)**

---

## Path B — Manual install

If you want to run the server manually (for testing, scripting, or a custom MCP client):

```bash
# Clone and install (same as above)
git clone https://github.com/dosev-ai/mcp-office.git
cd mcp-office
python -m venv .venv
.venv\Scripts\activate
pip install -e ./excelmcp

# Set required environment variables
set EXCEL_ALLOWED_DIRS=C:\path\to\your\files
set EXCEL_ENABLE_WRITE=true

# Run the server (stdio mode)
python -m excelmcp.server
```

The server speaks JSON-RPC 2.0 over stdin/stdout. Send an `initialize` message to confirm it’s alive:

```bash
echo {"jsonrpc":"2.0","method":"initialize","id":1,"params":{}} | python -m excelmcp.server
```

You should receive a `result` response within a few seconds.

---

## Environment variables

| Variable | Required | Description |
|---|---|---|
| `EXCEL_ALLOWED_DIRS` | Yes | Semicolon-separated list of directories excelmcp can access. Files outside this list are rejected. |
| `EXCEL_ENABLE_WRITE` | No | Set to `true` to enable write operations. Default: read-only. |
| `EXCEL_MAX_RANGE_CELLS` | No | Maximum cells per write operation. Default: 10000. |
| `EXCEL_SESSION_TIMEOUT` | No | COM session timeout in seconds. Default: 30. |

---

## Troubleshooting

**`EXCEL_ALLOWED_DIRS` error**
The file you’re trying to access is outside the allowed directories. Add its parent folder to `EXCEL_ALLOWED_DIRS` in your MCP config.

**`EXCEL_ENABLE_WRITE` error**
Write operations are disabled by default. Add `"EXCEL_ENABLE_WRITE": "true"` to the `env` block in your MCP config and restart your client.

**COM errors / Excel not responding**
Make sure Excel is installed and not blocked on a dialog box. If you’re working with OneDrive-synced files, copy the file to a local path (e.g. `C:\Temp`) before running write operations — COM writes to OneDrive-synced paths can hang.

**`capabilities()` returns no tools**
The server started but the tool list is empty. Check that the `command` path in your config points to the correct `.venv` Python executable and that `pip install -e ./excelmcp` completed without errors.

**First run friction?**
Open a [First Run Report](https://github.com/dosev-ai/mcp-office/issues/new?template=first_run_report.yml). These directly shape the next release.
