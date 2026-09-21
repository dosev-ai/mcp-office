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

    available, coming_next = readme.split("### 🚧 Coming next", 1)
    assert "### ✅ Available now" in available
    assert all(package in available for package in ("excelmcp", "pptmcp", "wordmcp"))
    assert "mailmcp" in coming_next
