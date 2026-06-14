# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| Latest (`main`) | ✅ |
| Older releases | ❌ — upgrade to latest |

## Scope

MCP Office servers are **local-first** tools that run entirely on your Windows machine. They communicate with your MCP client over `stdio` — no network calls, no cloud dependency. Accordingly:

- **In scope**: memory-safety bugs, path traversal bypasses, formula injection, allowlist bypass, mutation-gate bypass (the `ENABLE_WRITE`/`confirm` gates), stdout protocol leaks that could affect clients.
- **Out of scope**: hosted infrastructure attacks (no hosted tier exists), social engineering, physical-access attacks.

## Reporting a Vulnerability

Please **do not open a public GitHub issue** for security vulnerabilities.

Report privately via [GitHub Security Advisories](https://github.com/dosev-ai/mcp-office/security/advisories/new). Include:

1. Which package (`excelmcp`, `pptmcp`, `wordmcp`, `mailmcp`, `shared`)
2. A minimal reproduction (env vars, MCP call, input)
3. The observed vs. expected behaviour
4. Your assessment of impact (data exposure, arbitrary file write, etc.)

We aim to acknowledge reports within **48 hours** and provide a resolution timeline within **7 days**.

## Security Design

Each server enforces two independent layers:

- **Allowlist gate** — file paths and Outlook folders are validated against `*_ALLOWLIST_ROOTS` / `*_ALLOWLIST_FOLDERS` env vars before any COM or file I/O.
- **Mutation gate** — every write/send/delete tool requires `*_ENABLE_WRITE=true` (env) **and** `confirm=True` (parameter). Neither gate alone is sufficient.

See [`CLAUDE.md`](CLAUDE.md) for the full security invariant specification.
