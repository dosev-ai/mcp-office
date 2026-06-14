# MCP Office

> Local-first, governed MCP servers for Microsoft Office — built for developers who want to treat Office files the way they treat code.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://python.org)
[![CI](https://github.com/dosev-ai/mcp-office/actions/workflows/ci.yml/badge.svg)](https://github.com/dosev-ai/mcp-office/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/mcp-office-excel.svg)](https://pypi.org/project/mcp-office-excel/)
[![UAT](https://img.shields.io/badge/UAT-E2E%20validated%20on%20Windows%2011-brightgreen.svg)](CHANGELOG.md)

---

## What is this?

MCP Office is a suite of [Model Context Protocol](https://modelcontextprotocol.io) servers that expose Microsoft Office capabilities as governed, deterministic tool calls. Each server runs locally on Windows, connects to your MCP client (Claude Desktop, VS Code Copilot, or any MCP-compatible client), and gives you structured control over Office files — without an in-app AI assistant.

**This is not a Copilot replacement.** It is a developer-first execution layer for Office automation.

---

## Packages

### ✅ Available now

| Package | What it does | Install | Tools |
|---|---|---|---|
| [`excelmcp`](excelmcp/) | Read, write, style, validate, and export Excel workbooks | `pip install -e ./excelmcp` | 60+ |
| [`pptmcp`](pptmcp/) | Build, edit, review, and export PowerPoint presentations. Output Contract framework for machine-verifiable slide specs | `pip install -e ./shared && pip install -e ./pptmcp` | 48 |
| [`wordmcp`](wordmcp/) | Template assembly, tracked-changes support, and structural QA for Word documents | `pip install -e ./wordmcp` | 51 |

### 🚧 Coming next

| Package | Status |
|---|---|
| `mailmcp` | In development — Outlook email, calendar, contacts, and MailRepo search |

New packages are added as they complete their proof cycle. See [ROADMAP.md](ROADMAP.md).

---

## Quick start

### Prerequisites

- Windows 10 or 11
- Python 3.11 or later (`python --version`)
- Git (`git --version`)
- [Claude Desktop](https://claude.ai/download) or VS Code with GitHub Copilot
- Microsoft Office (Excel / PowerPoint / Word) — required for COM-backed tools (styling, PDF export, tracked-changes)

### Install

> **PyPI note:** `excelmcp` on PyPI is an unrelated third-party package. Do **not** `pip install excelmcp` — that ships you a stranger's code. The published suite packages are `mcp-office-excel` (PyPI) and source-editable installs below. `pip install excelmcp` is never the right command for this project.

```bash
# Clone the repo
git clone https://github.com/dosev-ai/mcp-office.git
cd mcp-office

# Create one shared venv for all packages
python -m venv .venv
.venv\Scripts\activate

# Install whichever packages you want (each is independent)
pip install -e ./excelmcp
pip install -e ./wordmcp

# pptmcp depends on the shared library — install both
pip install -e ./shared && pip install -e ./pptmcp
```

### Configure Claude Desktop

Open `%APPDATA%\Claude\claude_desktop_config.json` (create it if it doesn't exist) and add the servers you installed:

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
    },
    "powerpoint-pptmcp": {
      "command": "C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe",
      "args": ["-m", "pptmcp.server"],
      "env": {
        "PPT_ALLOWLIST_ROOTS": "C:\\path\\to\\your\\files",
        "PPT_ENABLE_WRITE": "true"
      }
    },
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

Replace `C:\\path\\to\\mcp-office` with the absolute path where you cloned the repo, and `C:\\path\\to\\your\\files` with the directory where your Office files live. Restart Claude Desktop after saving.

### Verify

In Claude Desktop, send:

```
Call capabilities() on excel-excelmcp
Call capabilities() on powerpoint-pptmcp
Call capabilities() on word-wordmcp
```

Each should return a tool list (60+ for Excel, 48 for PowerPoint, 51 for Word). If a server is missing, check the `command` path points to your `.venv` Python executable.

Full per-package guides: [excelmcp/README.md](excelmcp/README.md) · [pptmcp/README.md](pptmcp/README.md) · [wordmcp/README.md](wordmcp/README.md)

Detailed step-by-step: [docs/quickstart.md](docs/quickstart.md)

---

## Requirements

| Requirement | Notes |
|---|---|
| Windows 10 or 11 | COM automation requires Windows |
| Python 3.11+ | `python --version` to confirm |
| Git | For cloning the repo |
| Microsoft Office | Required for COM-dependent tools (styling, PDF export, tracked-changes). Read-only docx/xlsx/pptx tools work without Office. |
| MCP client | [Claude Desktop](https://claude.ai/download) **or** [VS Code with Copilot](https://marketplace.visualstudio.com/items?itemName=GitHub.copilot) |

---

## Architecture

```
Your MCP client (Claude Desktop / VS Code Copilot / other)
        │
        │  MCP stdio protocol
        ↓
  MCP Office servers (local Python processes)
   ├─ excelmcp    — Excel automation (live)
   ├─ pptmcp      — PowerPoint automation (live)
   ├─ wordmcp     — Word automation (live)
   └─ mailmcp     — Outlook + MailRepo (coming)
        │
        │  COM / openpyxl / python-pptx / python-docx
        ↓
  Microsoft Office (local installation)
```

Each server is a standalone `stdio` MCP server. No network calls. No cloud dependency. Your files stay local.

---

## Contributing

This project is in active development. The best way to contribute right now:

1. **Try any package** (excelmcp, pptmcp, wordmcp) and open a [First Run Report](https://github.com/dosev-ai/mcp-office/issues/new?template=first_run_report.yml)
2. **Report bugs** via [GitHub Issues](https://github.com/dosev-ai/mcp-office/issues/new?template=bug_report.yml)
3. **Ask questions or share what you built** in [GitHub Discussions](https://github.com/dosev-ai/mcp-office/discussions)

See [CONTRIBUTING.md](CONTRIBUTING.md) for development guidelines.

---

## License

MIT — see [LICENSE](LICENSE).
