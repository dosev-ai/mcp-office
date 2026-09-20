# MailMCP — Outlook MCP Server

> **Status: Coming next.** MailMCP is integrated into this repository's single `mcp-office` source distribution, but it is **not yet claimed as available in the published PyPI artifact**. Public availability will be claimed only after Windows Outlook UAT and published-artifact verification pass.

MailMCP is the local Outlook automation server in the MCP Office suite. It exposes Outlook email, folders, calendar, contacts, tasks, categories, meeting requests, and related operations over local MCP stdio.

## Public boundary

- Windows with **classic Microsoft Outlook desktop** is required.
- Local stdio execution only; no hosted or remote-execution claim.
- Public scope is Outlook automation only. MailRepo, retrieval, embeddings, Cortex, Action Worker, and other private/internal integrations are not included.
- Reads are bounded and folder-allowlisted.
- Writes, sends, deletes, and forwarding-rule mutations are disabled by default.
- Mutations require the relevant enable flag and an explicit `confirm=True` call where the operation is confirmation-gated.
- Tests and release evidence must use synthetic or dedicated test-mailbox data only.

## What the source-integrated server can do

The public server exposes Outlook operations across these areas:

- **Mail and folders:** health, accounts, folders, messages, search, attachments, recipients, mailbox statistics, conversation threads, forwarding, and bounded multi-folder search.
- **Compose and message actions:** draft composition/editing, replies, sending, moving, flags, read state, bulk message actions, and forwarding rules.
- **Calendar and meetings:** event listing/detail, meeting drafts, event updates, free/busy checks, and meeting-request handling.
- **Tasks and categories:** task operations, message-task extraction, category operations, and meeting requests.
- **Contacts and folders:** contact listing/detail/search and controlled folder creation.
- **Context:** bounded Outlook mail-context retrieval for agent workflows.

The public source intentionally does not provide private search/retrieval systems or a public registry launcher.

## Source-integrated installation

The repository build now contains MailMCP as the `mailmcp` console entry point. Until the next published artifact is verified, validate it from source:

```bat
git clone https://github.com/dosev-ai/mcp-office.git
cd mcp-office
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -e ".[com]"
```

This installs the consolidated suite plus the Windows COM extra used by Outlook automation.

### Upgrade or uninstall the source build

After pulling a newer source revision, refresh the editable install with:

```bat
python -m pip install -e ".[com]"
```

To remove the consolidated package:

```bat
python -m pip uninstall mcp-office
```

## Safe default configuration

With no mutation flags enabled, MailMCP starts in a restrictive read-oriented posture:

| Variable | Default | Purpose |
|---|---:|---|
| `OUTLOOK_ALLOWLIST_FOLDERS` | `Inbox,Contacts` | Folders the server may access. Add `Drafts` or `Calendar` only when those workflows are needed. |
| `OUTLOOK_MAX_ITEMS` | `50` | Maximum items returned by bounded reads. |
| `OUTLOOK_MAX_BODY_CHARS` | `4000` | Maximum body characters returned. |
| `OUTLOOK_ATTACHMENT_MAX_MB` | `10` | Attachment size ceiling. |
| `OUTLOOK_REDACT_MODE` | `none` | Response redaction: `none`, `emails`, or `emails+domains`. For shared demos, `emails` is a safer starting point. |
| `OUTLOOK_ALLOWLIST_DOMAINS` | empty | Optional recipient-domain allowlist. Recommended before enabling compose/send workflows. |
| `OUTLOOK_ENABLE_WRITE` | `false` | Enables non-send write operations when their confirmation gate also passes. |
| `OUTLOOK_ENABLE_SEND` | `false` | Enables sending when the send confirmation and recipient policy also pass. |
| `OUTLOOK_ENABLE_DELETE` | `false` | Enables deletion when the delete confirmation gate also passes. |
| `OUTLOOK_ENABLE_RULES` | `false` | Enables forwarding-rule mutations. |
| `OUTLOOK_FOLDER_CACHE_TTL_SECONDS` | `300` | Folder-cache lifetime. |
| `OUTLOOK_GET_MESSAGE_TIMEOUT_SECS` | `30` | Single-message COM wait, clamped to 5–300 seconds. |

Per-account restrictions can be added with `OUTLOOK_ACCOUNT_1_*` through `OUTLOOK_ACCOUNT_10_*`, for example `OUTLOOK_ACCOUNT_1_EMAIL`, `OUTLOOK_ACCOUNT_1_ALLOWLIST_FOLDERS`, `OUTLOOK_ACCOUNT_1_ALLOWLIST_DOMAINS`, and the corresponding `ENABLE_*`, `MAX_*`, and `REDACT_MODE` values.

Configuration reloads may tighten startup restrictions but cannot widen them. Restart the server with the intended startup policy if you need a broader approved scope.

## Claude Desktop

Open `%APPDATA%\Claude\claude_desktop_config.json` and add:

```json
{
  "mcpServers": {
    "mail-mailmcp": {
      "command": "C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe",
      "args": ["-m", "mailmcp.server"],
      "env": {
        "OUTLOOK_ALLOWLIST_FOLDERS": "Inbox,Contacts",
        "OUTLOOK_REDACT_MODE": "emails"
      }
    }
  }
}
```

Restart Claude Desktop after changing the configuration.

## Claude Code

For a project-scoped Claude Code configuration, add the same local stdio server to `.mcp.json`:

```json
{
  "mcpServers": {
    "mail-mailmcp": {
      "command": "C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe",
      "args": ["-m", "mailmcp.server"],
      "env": {
        "OUTLOOK_ALLOWLIST_FOLDERS": "Inbox,Contacts",
        "OUTLOOK_REDACT_MODE": "emails"
      }
    }
  }
}
```

Alternatively, add the same project-scoped stdio server with the Claude Code CLI:

```bat
claude mcp add --scope project ^
  --env OUTLOOK_ALLOWLIST_FOLDERS=Inbox,Contacts ^
  --env OUTLOOK_REDACT_MODE=emails ^
  --transport stdio mail-mailmcp -- ^
  C:\path\to\mcp-office\.venv\Scripts\python.exe -m mailmcp.server
```

Check the stored server with `claude mcp get mail-mailmcp`. Project-scoped `.mcp.json` servers require the normal Claude Code approval/trust step on first use.

Keep mutation flags unset for the first run. Add only the folders and capabilities required by the workflow you are testing.

## VS Code with GitHub Copilot

Create or edit `.vscode/mcp.json`:

```json
{
  "servers": {
    "mail-mailmcp": {
      "type": "stdio",
      "command": "C:\\path\\to\\mcp-office\\.venv\\Scripts\\python.exe",
      "args": ["-m", "mailmcp.server"],
      "env": {
        "PYTHONPATH": "C:\\path\\to\\mcp-office\\mailmcp\\src",
        "OUTLOOK_ALLOWLIST_FOLDERS": "Inbox,Contacts",
        "OUTLOOK_REDACT_MODE": "emails"
      }
    }
  }
}
```

VS Code also supports user-level MCP configuration; the workspace example above keeps the MailMCP configuration explicit and easy to remove.

## Verify the source-integrated server

Start with read-only checks:

```text
Call outlook_health on mail-mailmcp.
Call outlook_list_accounts on mail-mailmcp.
Call outlook_list_folders on mail-mailmcp.
```

A healthy server should connect to classic Outlook and return only the policy-bounded account/folder data.

You can also start the stdio server directly:

```bat
.venv\Scripts\python.exe -m mailmcp.server
```

The process waits on stdin/stdout for MCP traffic; that is expected.

## Enabling a write workflow

Enable only the minimum capability needed.

For example, a dedicated synthetic draft/send test may use:

```text
OUTLOOK_ALLOWLIST_FOLDERS=Inbox,Contacts,Drafts
OUTLOOK_ALLOWLIST_DOMAINS=example.test
OUTLOOK_ENABLE_WRITE=true
OUTLOOK_ENABLE_SEND=true
OUTLOOK_REDACT_MODE=emails
```

The environment flags do not bypass tool-level confirmation. Draft editing/sending is restricted to verified unsent Drafts items, and recipient/domain checks are re-evaluated before protected operations.

Calendar workflows similarly require `Calendar` in `OUTLOOK_ALLOWLIST_FOLDERS`. Enabling a write flag never implicitly grants folder access.

## Privacy and safety notes

- MailMCP runs as a local stdio process and talks to local Outlook COM. Your MCP client or model may have its own data-handling behavior; MailMCP does not change that client boundary.
- Explicit account scoping requires verifiable Outlook DeliveryStore ownership. Display names alone are not treated as ownership evidence.
- Returned values can redact email addresses/domains according to `OUTLOOK_REDACT_MODE`; opaque Outlook EntryIDs remain unchanged.
- Recipient allowlists fail closed. Disjoint global and per-account domain policies are denied rather than treated as unrestricted.
- Attachment inputs are restricted to approved file types and local roots; dangerous executable/script extensions are rejected.
- Folder errors, scan caps, and search timeouts are surfaced as partial/error state rather than silently returning an unfiltered result.
- A running Outlook COM call cannot always be forcibly interrupted. If Outlook is stalled, open it manually and resolve any blocking dialog before retrying.

## Troubleshooting

**The server starts but Outlook calls fail**

Confirm you are on Windows, classic Outlook desktop is installed, and the profile opens normally. New Outlook does not provide the classic COM automation surface used here.

**`ModuleNotFoundError: mailmcp`**

Refresh the consolidated source install with `python -m pip install -e ".[com]"` from the repository root, then confirm the MCP client uses that virtual environment.

**`No module named pythoncom` / `win32com`**

Install the consolidated COM extra with `python -m pip install -e ".[com]"` from the repository root, or otherwise ensure `pywin32>=306` is installed in the same virtual environment used by the MCP client.

**A folder is denied**

Add the folder explicitly to `OUTLOOK_ALLOWLIST_FOLDERS` and restart. Default access is only `Inbox,Contacts`.

**A write/send/delete/rule operation is denied**

Check both layers: the relevant `OUTLOOK_ENABLE_*` variable must be `true`, and the tool's confirmation/policy requirements must also pass.

**A config reload refuses a wider policy**

That is intentional. Reload can tighten policy but cannot widen the startup boundary. Stop the server, change the startup environment, and start a fresh process.

**Account-scoped operation cannot resolve the mailbox**

MailMCP requires a uniquely attributable DeliveryStore for explicit account scope. A matching Outlook display label is not sufficient.

## Current release path

The remaining public-release sequence is intentionally short:

1. run synthetic Windows + classic Outlook UAT from a fresh consolidated installation;
2. publish and verify the artifact;
3. only then change public listings from **Coming next** to **Available**.

Until those gates pass, this README documents the source-integrated MailMCP server, not a released MailMCP availability claim.
