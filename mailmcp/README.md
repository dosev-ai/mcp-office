# MailMCP

MailMCP is the local Outlook automation server in the MCP Office suite. This migration candidate is **not yet a public release**; the public suite should continue to describe MailMCP as **Coming next** until release UAT and published-artifact verification pass.

## Boundary

- Windows with classic Microsoft Outlook desktop is required.
- Local stdio execution only; no hosted or remote-execution claim.
- Public scope is Outlook automation only. Private repository, retrieval, embedding, and internal governance integrations are not included.
- Reads are folder-allowlisted and bounded.
- Writes, sends, deletes, and forwarding-rule mutations are disabled by default and use explicit configuration plus confirmation gates.
- Tests and release evidence must use synthetic or dedicated test-mailbox data only.

## Candidate safety and configuration notes

The default folder allowlist remains `Inbox,Contacts`. Draft editing and sending require explicitly adding `Drafts` to `OUTLOOK_ALLOWLIST_FOLDERS`; calendar reads and updates require explicitly adding `Calendar`. Enabling a mutation or send flag does not implicitly grant folder access. The relevant enable flag and `confirm=True` are still required. For a dedicated synthetic compose/edit/send test, configure `Inbox,Contacts,Drafts`, enable non-send writes and sending, and restrict recipient domains to the test mailbox domain before starting the server.

Draft edit/send operations accept only unsent MailItems in the verified default Drafts folder. Contact-by-ID reads require the verified default Contacts folder. Explicit account scoping requires a verifiable account DeliveryStore ID; display names are not ownership evidence. Soft deletion refuses to purge an item already in Deleted Items; permanent deletion must be requested explicitly.

Configuration reload cannot widen the startup folder or domain scope of an existing account profile. Disjoint global and per-account recipient-domain lists are denied, not converted to unrestricted access. Redaction applies to returned contact and task-extraction text, while opaque EntryIDs remain unchanged.

Multi-folder search waits for bounded folder results before selecting the globally newest results. Timeouts, folder errors, and scan caps are reported as partial results. The timeout bounds the caller's wait; a running Outlook COM call cannot be forcibly interrupted. Do not repeatedly retry a stalled Outlook profile. A failed Outlook filter raises an error rather than returning an unfiltered list.

No registry launcher is supplied at this migration gate. Launcher configuration will follow verified root-distribution installation and release UAT; this candidate does not claim that an installed public package already contains MailMCP.

Full install, client configuration, examples, troubleshooting, privacy notes, and capability documentation are completed in the downstream documentation gate before release.
