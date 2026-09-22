# Roadmap

MCP Office ships as one `mcp-office` distribution with distinct local MCP servers. Public package claims follow verified release evidence, not source presence alone.

## Current status

| Package | Status | Available |
|---|---|---|
| excelmcp | ✅ Available now | Now |
| pptmcp | ✅ Available now | Now |
| wordmcp | ✅ Available now | Now |
| mailmcp | 🚧 Source-integrated; release verification pending | Coming next |

Excel, PowerPoint, and Word are the currently released public suite. MailMCP is integrated into the repository and the 0.8.0 source candidate, but it remains **Coming next** until Windows Outlook UAT, published-artifact verification, and the downstream listing/claim gate pass.

## What “proof cycle” means

Before a package changes to **Available now**:

1. implementation and governed review converge;
2. exact-candidate CI and package-specific UAT pass;
3. the package ships through the approved distribution path;
4. the published artifact is installed and verified in a clean target environment;
5. public listings and documentation are reconciled to the verified artifact.

Source integration by itself is not a release claim.

## Available now

### excelmcp

Structured workbook read/write, formatting, validation, charts, exports, and local Excel automation.

### pptmcp

Build, edit, review, and export PowerPoint presentations with Output Contract support. The current always-registered PowerPoint surface is **51 tools**. Platform-conditional COM tools are reported separately.

### wordmcp

Document assembly, structured editing, tracked changes, review/evidence, and export. The canonical public callable surface is **50 endpoints**.

## mailmcp — coming next

MailMCP provides local Outlook email, calendar, contacts, tasks, and related automation. It is present in the single source distribution, but public availability remains gated on Windows Outlook UAT and published-artifact verification.

The public boundary remains local-first: classic Outlook on Windows is required for Outlook automation; write/send/delete/rules controls remain explicitly gated; source availability must not be described as a published release.
