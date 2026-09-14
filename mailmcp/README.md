# MailMCP

MailMCP is the local Outlook automation server in the MCP Office suite. This migration candidate is **not yet a public release**; the public suite should continue to describe MailMCP as **Coming next** until release UAT and published-artifact verification pass.

## Boundary

- Windows with classic Microsoft Outlook desktop is required.
- Local stdio execution only; no hosted or remote-execution claim.
- Public scope is Outlook automation only. Private repository, retrieval, embedding, and internal governance integrations are not included.
- Reads are folder-allowlisted and bounded.
- Writes, sends, deletes, and forwarding-rule mutations are disabled by default and use explicit configuration plus confirmation gates.
- Tests and release evidence must use synthetic or dedicated test-mailbox data only.

Full install, client configuration, examples, troubleshooting, privacy notes, and capability documentation are completed in the downstream documentation gate before release.
