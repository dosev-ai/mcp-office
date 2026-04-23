# Roadmap

MCP Office follows a staged package exposure model. Each package goes public only after its proof cycle clears.

## Current status

| Package | Status | Available |
|---|---|---|
| excelmcp | ✅ Phase 1 lead | Now |
| mailmcp | 🚧 In development | After excelmcp proof cycle |
| pptmcp | 🗓️ Roadmap | TBD |
| wordmcp | 🗓️ Roadmap | TBD |

## What “proof cycle” means

Before a package becomes the public lead:
1. It completes internal delivery and UAT
2. It ships clean into this repo
3. First-run confirmations from external users are collected
4. Friction points are addressed
5. The next package is promoted

## excelmcp — Phase 1 (current)

Flagship workflow: **structured input → generate Excel artifact → validate contract → output**

Core tools:
- `range_io` — read/write cell ranges
- `apply_style` — formatting via COM or openpyxl
- `validate_contract` — check workbook against embedded schema
- `export_as_pdf` — native Excel PDF export via COM
- 60+ additional tools for metadata, charts, named ranges, and more

## mailmcp — coming next

Outlook email, calendar, contacts, and MailRepo full-text search over your mail history. Expanding after excelmcp proof.

## pptmcp, wordmcp — roadmap

In active internal development. Will be added to this repo when their proof cycles are ready.
