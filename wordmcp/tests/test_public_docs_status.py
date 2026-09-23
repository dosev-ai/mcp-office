from pathlib import Path

from wordmcp import server as srv

ROOT = Path(__file__).resolve().parents[2]


def _call_tool(tool):
    fn = getattr(tool, "fn", tool)
    return fn()


def test_public_word_count_matches_executable_capabilities():
    count = len(_call_tool(srv.capabilities)["tools"])
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    quickstart = (ROOT / "docs" / "quickstart.md").read_text(encoding="utf-8")
    package_readme = (ROOT / "wordmcp" / "README.md").read_text(encoding="utf-8")

    assert f"| `pip install mcp-office` | {count} |" in readme.split("| [`wordmcp`](wordmcp/) |", 1)[1].split("\n", 1)[0]
    assert f"You should see {count} tools." in quickstart
    assert f"**{count} canonical callable endpoints**" in package_readme
