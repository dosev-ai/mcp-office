# Quick Start

> **Packages covered:** [excelmcp](#excelmcp) · [pptmcp](#pptmcp) · [wordmcp](#wordmcp)

Get any MCP Office package running in under 10 minutes. Each section is self-contained — jump to the package you want.

---

## excelmcp

### Path A — Agent-first (recommended)

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
        "EXCEL_ALLOWLIST_ROOTS": "C:\\path\\to\\your\\files",
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
        "EXCEL_ALLOWLIST_ROOTS": "C:\\path\\to\\your\\files",
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
set EXCEL_ALLOWLIST_ROOTS=C:\path\to\your\files
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
| `EXCEL_ALLOWLIST_ROOTS` | Yes | Semicolon-separated list of directories excelmcp can access. Files outside this list are rejected. |
| `EXCEL_ENABLE_WRITE` | No | Set to `true` to enable write operations. Default: read-only. |
| `EXCEL_MAX_RANGE_CELLS` | No | Maximum cells per write operation. Default: 10000. |
| `EXCEL_SESSION_TIMEOUT` | No | COM session timeout in seconds. Default: 30. |

---

## Troubleshooting

**`EXCEL_ALLOWLIST_ROOTS` error**
The file you’re trying to access is outside the allowed directories. Add its parent folder to `EXCEL_ALLOWLIST_ROOTS` in your MCP config.

**`EXCEL_ENABLE_WRITE` error**
Write operations are disabled by default. Add `"EXCEL_ENABLE_WRITE": "true"` to the `env` block in your MCP config and restart your client.

**COM errors / Excel not responding**
Make sure Excel is installed and not blocked on a dialog box. If you’re working with cloud-synced files, copy the file to a local path (e.g. `C:\Temp`) before running write operations — COM writes to cloud-synced paths can hang.

**`capabilities()` returns no tools**
The server started but the tool list is empty. Check that the `command` path in your config points to the correct `.venv` Python executable and that `pip install -e ./excelmcp` completed without errors.

**First run friction?**
Open a [First Run Report](https://github.com/dosev-ai/mcp-office/issues/new?template=first_run_report.yml). These directly shape the next release.

---

## pptmcp

Build, edit, review, and export PowerPoint presentations. Includes the Output Contract framework for machine-verifiable slide specs.

### Prerequisites

- Windows 10 or 11
- Python 3.11 or later
- [Claude Desktop](https://claude.ai/download) **or** [VS Code with GitHub Copilot](https://marketplace.visualstudio.com/items?itemName=GitHub.copilot)
- Microsoft PowerPoint (required for COM tools: PDF export, slide show, chart recalculation). Read-only and python-pptx tools work without PowerPoint.

### Step 1 — Clone and install

pptmcp depends on the `shared` library. Install both:

```bash
git clone https://github.com/dosev-ai/mcp-office.git
cd mcp-office
python -m venv .venv
.venv\Scripts\activate
pip install -e ./shared
pip install -e ./pptmcp
```

### Step 2 — Configure Claude Desktop

Open `%APPDATA%\Claude\claude_desktop_config.json` and add the `powerpoint-pptmcp` server:

```json
{
  "mcpServers": {
    "powerpoint-pptmcp": {
      "command": "C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe",
      "args": ["-m", "pptmcp.server"],
      "env": {
        "PPT_ALLOWLIST_ROOTS": "C:\\path\\to\\your\\files",
        "PPT_ENABLE_WRITE": "true"
      }
    }
  }
}
```

Replace:
- `C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe` — full path to the Python executable inside the `.venv` you created
- `C:\\path\\to\\your\\files` — directory where your PowerPoint files live

Restart Claude Desktop after saving.

### Step 2 (alternative) — Configure VS Code Copilot

```json
{
  "servers": {
    "powerpoint-pptmcp": {
      "type": "stdio",
      "command": "C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe",
      "args": ["-m", "pptmcp.server"],
      "env": {
        "PPT_ALLOWLIST_ROOTS": "C:\\path\\to\\your\\files",
        "PPT_ENABLE_WRITE": "true"
      }
    }
  }
}
```

### Step 3 — Verify

In Claude Desktop or VS Code Copilot, ask:

```
Call capabilities() on powerpoint-pptmcp
```

You should see 48 tools. If you do, you're ready.

### Environment variables

| Variable | Required | Description |
|---|---|---|
| `PPT_ALLOWLIST_ROOTS` | Yes | Semicolon-separated list of directories pptmcp can access. |
| `PPT_ENABLE_WRITE` | No | Set to `true` to enable write/delete operations. Default: read-only. |
| `PPT_MAX_SLIDES` | No | Maximum slides returned per call. Default: 50. |

### Troubleshooting

**`PPT_ALLOWLIST_ROOTS` error**
The file is outside the allowed directories. Add its parent folder to `PPT_ALLOWLIST_ROOTS`.

**COM errors / PowerPoint not responding**
Make sure PowerPoint is installed and not blocked on a dialog. Copy cloud-synced files to a local path (e.g. `C:\Temp`) before write operations.

**Shape coordinates look wrong**
pptmcp shape positions (`left`, `top`, `width`, `height`) are in **inches**, not EMU. Passing EMU values corrupts the file. Example: `left=1.0` means 1 inch from the left edge.

---

## wordmcp

Template assembly, tracked-changes support, and structural QA for Word documents. No live Word required for read operations and python-docx tools; COM tools require Word installed.

### Prerequisites

- Windows 10 or 11
- Python 3.11 or later
- [Claude Desktop](https://claude.ai/download) **or** [VS Code with GitHub Copilot](https://marketplace.visualstudio.com/items?itemName=GitHub.copilot)
- Microsoft Word (required for COM tools: PDF/HTML export, tracked-changes management). Read-only and python-docx tools work without Word.

### Step 1 — Clone and install

wordmcp has no external library dependencies beyond python-docx — install it directly:

```bash
git clone https://github.com/dosev-ai/mcp-office.git
cd mcp-office
python -m venv .venv
.venv\Scripts\activate
pip install -e ./wordmcp
```

### Step 2 — Configure Claude Desktop

Open `%APPDATA%\Claude\claude_desktop_config.json` and add the `word-wordmcp` server:

```json
{
  "mcpServers": {
    "word-wordmcp": {
      "command": "C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe",
      "args": ["-m", "wordmcp.server"],
      "env": {
        "WORD_ALLOWLIST_ROOTS": "C:\\path\\to\\your\\files",
        "WORD_ENABLE_WRITE": "true"
      }
    }
  }
}
```

Replace:
- `C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe` — full path to the Python executable inside the `.venv` you created
- `C:\\path\\to\\your\\files` — directory where your Word files live

Restart Claude Desktop after saving.

### Step 2 (alternative) — Configure VS Code Copilot

```json
{
  "servers": {
    "word-wordmcp": {
      "type": "stdio",
      "command": "C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe",
      "args": ["-m", "wordmcp.server"],
      "env": {
        "WORD_ALLOWLIST_ROOTS": "C:\\path\\to\\your\\files",
        "WORD_ENABLE_WRITE": "true"
      }
    }
  }
}
```

### Step 3 — Verify

In Claude Desktop or VS Code Copilot, ask:

```
Call capabilities() on word-wordmcp
```

You should see 51 tools. If you do, you're ready.

### Environment variables

| Variable | Required | Description |
|---|---|---|
| `WORD_ALLOWLIST_ROOTS` | Yes | Semicolon-separated list of directories wordmcp can access. |
| `WORD_ENABLE_WRITE` | No | Set to `true` to enable write/delete operations. Default: read-only. |
| `WORD_MAX_TEXT_CHARS` | No | Maximum characters returned by text export tools. Default: 50000. |

### Troubleshooting

**`WORD_ALLOWLIST_ROOTS` error**
The file is outside the allowed directories. Add its parent folder to `WORD_ALLOWLIST_ROOTS`.

**COM tools raise `ToolError: COM not available`**
COM tools (PDF export, tracked-changes) require Windows + Word installed. These tools work only on the machine where Word is installed. python-docx tools (read, write, find/replace, tables) work on any OS.

**PDF export fails silently**
Close any PDF viewer (e.g. Adobe Reader) that has the target output file open before running PDF export. A locked file causes a silent COM failure.

**First run friction?**
Open a [First Run Report](https://github.com/dosev-ai/mcp-office/issues/new?template=first_run_report.yml). These directly shape the next release.

