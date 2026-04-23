# Contributing to MCP Office

Thanks for your interest. Here’s how to help right now.

## Best ways to contribute (Phase 1)

### 1. Try excelmcp and report your first-run experience
Open a [First Run Report](https://github.com/dosev-ai/mcp-office/issues/new?template=first_run_report.yml). These directly shape the next iteration.

### 2. Report bugs
Use the [Bug Report template](https://github.com/dosev-ai/mcp-office/issues/new?template=bug_report.yml). Include your OS, Office version, package version, and steps to reproduce.

### 3. Ask questions and share what you built
[GitHub Discussions](https://github.com/dosev-ai/mcp-office/discussions) is the place for questions, ideas, and show-and-tell.

### 4. Submit a feature request
Use the [Feature Request template](https://github.com/dosev-ai/mcp-office/issues/new?template=feature_request.yml).

## Development setup

```bash
git clone https://github.com/dosev-ai/mcp-office.git
cd mcp-office
python -m venv .venv
.venv\Scripts\activate
pip install -e "./excelmcp[dev]"
```

Run the tests:
```bash
pytest excelmcp/tests/ -q
```

Lint:
```bash
ruff check excelmcp/
```

## Pull request guidelines

- One logical change per PR
- Tests required for any new tool or behaviour change
- Ruff clean before submitting
- PR description must include: what changed, why, and how to test it

## Code of conduct

Be direct, be constructive, be kind. We’re building something useful — focus on the work.
