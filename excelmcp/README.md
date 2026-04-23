# excelmcp

> Excel MCP server — read, write, style, validate, and export Excel workbooks via 60+ governed MCP tools.

Part of the [MCP Office](https://github.com/dosev-ai/mcp-office) suite.

---

## Install

```bash
pip install -e ./excelmcp
```

Or from the repo root:

```bash
pip install -e "./excelmcp[dev]"  # includes test dependencies
```

## Configure

Copy `mcp.json.template` from the repo root and fill in your paths:

```json
{
  "mcpServers": {
    "excel-excelmcp": {
      "command": "C:\\path\\to\\.venv\\Scripts\\python.exe",
      "args": ["-m", "excelmcp.server"],
      "env": {
        "EXCEL_ALLOWED_DIRS": "C:\\path\\to\\your\\files",
        "EXCEL_ENABLE_WRITE": "true"
      }
    }
  }
}
```

**Environment variables:**

| Variable | Required | Description |
|---|---|---|
| `EXCEL_ALLOWED_DIRS` | Yes | Semicolon-separated list of directories excelmcp can access |
| `EXCEL_ENABLE_WRITE` | No | Set to `true` to enable write operations (default: read-only) |
| `EXCEL_MAX_RANGE_CELLS` | No | Max cells per write operation (default: 10000) |

## The flagship workflow

```
structured input → range_io write → apply_style → validate_contract → export_as_pdf
```

Every write requires `confirm=True`. Nothing runs without explicit confirmation.

## Key tools

| Tool | What it does |
|---|---|
| `range_io` | Read or write cell ranges with type validation and formula injection protection |
| `apply_style` | Apply formatting (font, colour, borders, number format) via COM or openpyxl |
| `validate_contract` | Check workbook against its embedded `_MCP_META` schema contract |
| `export_as_pdf` | Export workbook or sheet to PDF via Excel native COM export |
| `bulk_range_write` | Write multiple ranges atomically in one call |
| `capabilities` | List all available tools and current server configuration |

Run `capabilities()` in your MCP client to see the full tool list.

## Requirements

- Windows 10/11
- Python 3.11+
- Microsoft Excel (required for COM-dependent tools: `apply_style`, `export_as_pdf`, live session operations)
- Non-COM tools (read, range_io write, validate_contract) work without Excel installed

## Run tests

```bash
pytest excelmcp/tests/ -q
```

## Troubleshooting

**`EXCEL_ALLOWED_DIRS` error:** The file path must be inside one of your configured allowed directories. Add the directory to `EXCEL_ALLOWED_DIRS` in your MCP config.

**`EXCEL_ENABLE_WRITE` error:** Write operations are disabled by default. Set `EXCEL_ENABLE_WRITE=true` in your MCP config.

**COM errors:** Ensure Microsoft Excel is installed and not already blocking on a dialog. Copy the target file to a local path (not OneDrive-synced) before writing.

**First run friction?** Open a [First Run Report](https://github.com/dosev-ai/mcp-office/issues/new?template=first_run_report.yml) — these directly shape the next release.
