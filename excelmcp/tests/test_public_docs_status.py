from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
FIRST_WORKFLOW = ROOT / "docs" / "first-workflow.md"
README = ROOT / "README.md"


def test_first_workflow_matches_public_suite_availability():
    first_workflow = FIRST_WORKFLOW.read_text(encoding="utf-8")
    readme = README.read_text(encoding="utf-8")

    assert "Excel, PowerPoint, and Word are available" in first_workflow
    assert "Mail / Outlook remains the next package" in first_workflow
    assert "Excel is the lead package. Others are coming." not in first_workflow

    assert "### ✅ Available now" in readme
    assert "[`excelmcp`]" not in readme  # guard accidental escaped markdown in the canonical table
    assert "[`pptmcp`]" not in readme
    assert "[`wordmcp`]" not in readme
    assert "### 🚧 Coming next" in readme
    assert "[`mailmcp`]" not in readme
