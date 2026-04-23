# MCP Office

> Local-first, governed MCP servers for Microsoft Office — built for developers who want to treat Office files the way they treat code.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://python.org)

---

## What is this?

MCP Office is a suite of [Model Context Protocol](https://modelcontextprotocol.io) servers that expose Microsoft Office capabilities as governed, deterministic tool calls. Each server runs locally on Windows, connects to your MCP client (Claude Desktop, VS Code Copilot, or any MCP-compatible client), and gives you structured control over Office files — without an in-app AI assistant.

**This is not a Copilot replacement.** It is a developer-first execution layer for Office automation.

---

## Packages

### ✅ Available now

| Package | What it does | Install |
|---|---|---|
| [`excelmcp`](excelmcp/) | Read, write, style, validate, and export Excel workbooks via 60+ governed tools | `pip install -e ./excelmcp` |

### 🚧 Coming next

| Package | Status |
|---|---|
| `mailmcp` | In development — Outlook email, calendar, contacts, and MailRepo search |
| `pptmcp` | Roadmap — PowerPoint slide generation, review, and export |
| `wordmcp` | Roadmap — Word document creation and structured editing |

New packages are added as they complete their proof cycle. See [ROADMAP.md](ROADMAP.md).

---

## Quick start (excelmcp)

```bash
# 1. Clone
git clone https://github.com/dosev-ai/mcp-office.git
cd mcp-office

# 2. Create a virtual environment
python -m venv .venv
.venv\Scripts\activate  # Windows

# 3. Install excelmcp
pip install -e ./excelmcp

# 4. Configure your MCP client (see mcp.json.template)
```

Then copy `mcp.json.template` to your MCP client config directory and fill in your paths.

Full walkthrough: [excelmcp/README.md](excelmcp/README.md)

---

## Requirements

- Windows 10/11
- Python 3.11+
- Microsoft Office (for COM-dependent tools: styling, PDF export, live workbook operations)
- An MCP-compatible client: [Claude Desktop](https://claude.ai/download), [VS Code with Copilot](https://marketplace.visualstudio.com/items?itemName=GitHub.copilot)

---

## Architecture

```
Your MCP client (Claude Desktop / VS Code Copilot / other)
        │
        │  MCP stdio protocol
        ↓
  MCP Office servers (local Python processes)
   ├─ excelmcp    — Excel automation
   ├─ mailmcp     — Outlook + MailRepo (coming)
   ├─ pptmcp      — PowerPoint (roadmap)
   └─ wordmcp     — Word (roadmap)
        │
        │  COM / openpyxl / python-docx
        ↓
  Microsoft Office (local installation)
```

Each server is a standalone `stdio` MCP server. No network calls. No cloud dependency. Your files stay local.

---

## Contributing

This project is in active development. The best way to contribute right now:

1. **Try excelmcp** and open a [First Run Report](https://github.com/dosev-ai/mcp-office/issues/new?template=first_run_report.yml)
2. **Report bugs** via [GitHub Issues](https://github.com/dosev-ai/mcp-office/issues/new?template=bug_report.yml)
3. **Ask questions or share what you built** in [GitHub Discussions](https://github.com/dosev-ai/mcp-office/discussions)

See [CONTRIBUTING.md](CONTRIBUTING.md) for development guidelines.

---

## License

MIT — see [LICENSE](LICENSE).
