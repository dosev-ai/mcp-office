from pathlib import Path

from pptmcp.presentation_pptx import capabilities

ROOT = Path(__file__).resolve().parents[2]


def test_public_powerpoint_count_matches_executable_capabilities():
    count = len(capabilities()["tools"])
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    quickstart = (ROOT / "docs" / "quickstart.md").read_text(encoding="utf-8")
    package_readme = (ROOT / "pptmcp" / "README.md").read_text(encoding="utf-8")

    assert f"| `pip install mcp-office` | {count} |" in readme.split("| [`pptmcp`](pptmcp/) |", 1)[1].split("\n", 1)[0]
    assert f"You should see {count} PowerPoint tools." in quickstart
    assert f"**{count} always-registered tools**" in package_readme
