from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
FIRST_WORKFLOW = ROOT / "docs" / "first-workflow.md"
README = ROOT / "README.md"
QUICKSTART = ROOT / "docs" / "quickstart.md"
ROADMAP = ROOT / "ROADMAP.md"
PPT_README = ROOT / "pptmcp" / "README.md"
WORD_README = ROOT / "wordmcp" / "README.md"
SERVER_JSON = ROOT / "server.json"
PPT_PYPROJECT = ROOT / "pptmcp" / "pyproject.toml"
WORD_PYPROJECT = ROOT / "wordmcp" / "pyproject.toml"


def test_first_workflow_matches_public_suite_availability():
    first_workflow = FIRST_WORKFLOW.read_text(encoding="utf-8")
    readme = README.read_text(encoding="utf-8")

    assert "Excel, PowerPoint, and Word are available" in first_workflow
    assert "Mail / Outlook remains the next package" in first_workflow
    assert "Excel is the lead package. Others are coming." not in first_workflow

    available, after_coming_next = readme.split("### 🚧 Coming next", 1)
    coming_next = after_coming_next.split("\n---\n", 1)[0]
    assert "### ✅ Available now" in available
    assert all(package in available for package in ("excelmcp", "pptmcp", "wordmcp"))
    assert "| [`mailmcp`](mailmcp/) |" in coming_next


def test_current_public_tool_counts_and_roadmap_are_synchronized():
    readme = README.read_text(encoding="utf-8")
    quickstart = QUICKSTART.read_text(encoding="utf-8")
    roadmap = ROADMAP.read_text(encoding="utf-8")
    ppt_readme = PPT_README.read_text(encoding="utf-8")
    word_readme = WORD_README.read_text(encoding="utf-8")
    server_json = SERVER_JSON.read_text(encoding="utf-8")
    ppt_pyproject = PPT_PYPROJECT.read_text(encoding="utf-8")
    word_pyproject = WORD_PYPROJECT.read_text(encoding="utf-8")

    assert "| [`excelmcp`](excelmcp/) | Read, write, style, validate, and export Excel workbooks | `pip install mcp-office` | 65 |" in readme
    assert "| [`pptmcp`](pptmcp/) | Build, edit, review, and export PowerPoint presentations. Output Contract framework for machine-verifiable slide specs | `pip install mcp-office` | 51 |" in readme
    assert "| [`wordmcp`](wordmcp/) | Template assembly, tracked-changes support, and structural QA for Word documents | `pip install mcp-office` | 50 |" in readme
    assert "| [`mailmcp`](mailmcp/) | Integrated into the current source distribution; Windows Outlook UAT and published-artifact verification pending |" in readme
    assert "Each should return a tool list (65 for Excel, 51 for PowerPoint, 50 for Word)." in readme

    assert "You should see 51 PowerPoint tools." in quickstart
    assert "You should see 50 tools." in quickstart
    assert "You should see 48 tools." not in quickstart
    assert "You should see 51 tools. If you do, you're ready." not in quickstart

    assert "**51 always-registered tools**" in ppt_readme
    assert "2 platform-conditional COM-only tools" in ppt_readme
    assert "`total_tools`" not in ppt_readme
    assert "**50 canonical callable endpoints**" in word_readme
    assert "48 always-registered tools" not in ppt_readme
    assert "**51 tools**" not in word_readme

    assert "51 tools" in server_json
    assert "50 canonical callable endpoints" in server_json
    assert "51 tools" in ppt_pyproject
    assert "50 canonical callable endpoints" in word_pyproject

    assert "| pptmcp | ✅ Available now | Now |" in roadmap
    assert "| wordmcp | ✅ Available now | Now |" in roadmap
    assert "| mailmcp | 🚧 Source-integrated; release verification pending | Coming next |" in roadmap
    assert "pptmcp | 🗓️ Roadmap" not in roadmap
    assert "wordmcp | 🗓️ Roadmap" not in roadmap
