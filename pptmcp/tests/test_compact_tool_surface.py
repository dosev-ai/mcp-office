"""Compact pptmcp tool surface tests."""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

from pptmcp.presentation_pptx import capabilities

pytestmark = pytest.mark.enable_socket


def test_capabilities_document_compact_surface() -> None:
    caps = capabilities()
    compact = caps["compact_tool_surface"]

    assert compact["enabled_by_env"] == "PPT_COMPACT_TOOL_SURFACE"
    assert compact["tool_count"] == 27  # 24 original + 3 table mutation tools
    assert {"slide", "shape", "export", "declare_slide_contract"}.issubset(compact["tools"])
    assert {"add_slide", "set_table_style", "export_deck"}.isdisjoint(compact["tools"])
    assert caps["tool_bundles"]["render_check_iterate"] == [
        "declare_slide_contract",
        "validate_contract",
        "check_presentation_against_contract",
        "export",
        "review_slide_export",
        "produce_evidence_bundle",
    ]


def test_compact_server_registration_is_dispatcher_first() -> None:
    env = os.environ.copy()
    env["PPT_COMPACT_TOOL_SURFACE"] = "true"
    repo_root = Path(__file__).resolve().parents[2]
    src_paths = [
        str(repo_root / "shared" / "src"),
        str(repo_root / "powerpoint" / "src"),
    ]
    env["PYTHONPATH"] = os.pathsep.join(
        src_paths + ([env["PYTHONPATH"]] if env.get("PYTHONPATH") else [])
    )

    script = r"""
import json
from pptmcp import server

mcp = server.mcp
if hasattr(mcp, "_tool_manager") and hasattr(mcp._tool_manager, "_tools"):
    names = set(mcp._tool_manager._tools)
elif hasattr(mcp, "_tools"):
    names = set(mcp._tools)
elif hasattr(mcp, "_local_provider") and hasattr(mcp._local_provider, "_components"):
    names = {
        key.split(":")[1].split("@")[0]
        for key in mcp._local_provider._components
        if key.startswith("tool:")
    }
else:
    raise RuntimeError("Cannot discover FastMCP tools")

print(json.dumps(sorted(names)))
"""
    result = subprocess.run(
        [sys.executable, "-c", script],
        check=True,
        capture_output=True,
        text=True,
        env=env,
    )
    names = set(json.loads(result.stdout))

    assert len(names) == 27
    assert {"capabilities", "slide", "shape", "export", "add_content", "set_format", "declare_slide_contract"}.issubset(names)
    assert {
        "add_slide",
        "delete_slide",
        "set_table_style",
        "export_slide",
        "export_deck",
        "save",
        "create_presentation",
    }.isdisjoint(names)
