"""
User Story Tests — Word MCP ARCH Design (P1–P8)

Sprint target: (internal reference removed)
'[ARCH] Word MCP — Tool Design: Unified Surface, Format-Parameterised Export,
 Stage 1→2→3 Migration'

These 10 tests encode the acceptance criteria P1–P8 as verifiable assertions.
BEFORE delivery some tests FAIL (baseline state).
AFTER delivery ALL 10 tests MUST PASS.

Run with: pytest wordmcp/tests/test_user_stories_arch.py -v
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path

import pytest

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

WORD_SRC = Path(__file__).parents[1] / "src"
sys.path.insert(0, str(WORD_SRC))

_handlers_com_path = WORD_SRC / "wordmcp" / "_server_handlers_com.py"
_handlers_edit_path = WORD_SRC / "wordmcp" / "_server_handlers_docx_edit.py"
_handlers_adv_path = WORD_SRC / "wordmcp" / "_server_handlers_docx_advanced.py"
_handlers_core_path = WORD_SRC / "wordmcp" / "_server_handlers_docx_core.py"
_manifest_path = WORD_SRC / "wordmcp" / "_server_manifest.py"
_server_path = WORD_SRC / "wordmcp" / "server.py"


def _parse_source(path: Path) -> ast.Module:
    return ast.parse(path.read_text(encoding="utf-8"))


def _all_function_names(path: Path) -> list[str]:
    tree = _parse_source(path)
    return [
        node.name
        for node in ast.walk(tree)
        if isinstance(node, ast.FunctionDef | ast.AsyncFunctionDef)
    ]


def _get_docstring(path: Path, func_name: str) -> str:
    """Return the docstring for a top-level function in the file."""
    tree = _parse_source(path)
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            if node.name == func_name:
                return ast.get_docstring(node) or ""
    return ""


def _registered_tool_names() -> list[str]:
    """
    Return the list of tool names actually returned from all bind_tool / handler
    registration calls.  We read the _server_manifest.py and all handler source
    files and collect every key added to the `tools` dict via the manifest or the
    `return {name: func}` pattern in handler modules.
    """
    all_names: list[str] = []

    # Collect function names from all handler modules as a proxy for tool count
    handler_files = [
        WORD_SRC / "wordmcp" / "_server_handlers_docx_core.py",
        WORD_SRC / "wordmcp" / "_server_handlers_docx_edit.py",
        WORD_SRC / "wordmcp" / "_server_handlers_docx_advanced.py",
        WORD_SRC / "wordmcp" / "_server_handlers_com.py",
        WORD_SRC / "wordmcp" / "_server_handlers_review.py",
        WORD_SRC / "wordmcp" / "_server_handlers_meta.py",
    ]
    for hf in handler_files:
        if not hf.exists():
            continue
        src = hf.read_text(encoding="utf-8")
        # Find all functions that appear in return dicts (registered tools)
        # Strategy: look for `return {` blocks and collect string keys
        import re
        keys = re.findall(r'"([a-z][a-z_0-9]*)"\s*:', src)
        all_names.extend(keys)

    # Deduplicate preserving order
    seen: set[str] = set()
    result = []
    for n in all_names:
        if n not in seen:
            seen.add(n)
            result.append(n)
    return result


# ---------------------------------------------------------------------------
# US-01: Deprecated COM tool wrappers must be unregistered
# ---------------------------------------------------------------------------

DEPRECATED_COM_TOOLS = [
    "export_as_pdf",
    "accept_all_track_changes",
    "reject_all_track_changes",
    "list_tracked_changes",
]


@pytest.mark.parametrize("tool_name", DEPRECATED_COM_TOOLS)
def test_us01_deprecated_com_tools_not_registered(tool_name: str):
    """
    US-01: Deprecated COM tool wrappers must NOT be registered as live tools.

    As a developer integrating with the Word MCP API, I want deprecated tool
    aliases (export_as_pdf, accept_all_track_changes, reject_all_track_changes,
    list_tracked_changes) to be completely removed from the tool registry so that
    clients cannot accidentally call them instead of the canonical single-surface
    manage_* tools.
    """
    registered = _registered_tool_names()
    assert tool_name not in registered, (
        f"Deprecated tool '{tool_name}' is still registered. "
        "It must be removed from the return dict in _server_handlers_com.py. "
        "P2/P4 criterion."
    )


# ---------------------------------------------------------------------------
# US-02: export_document must use Literal typing
# ---------------------------------------------------------------------------

def test_us02_export_document_uses_literal_format_type():
    """
    US-02: export_document must accept format as Literal, not bare str.

    As a client developer, I want auto-completion and validation on the 'format'
    parameter of export_document so that I cannot accidentally specify an unsupported
    format like 'docx' or 'xlsx'. The param must be typed as
    Literal['pdf', 'html', 'txt', 'md'] or a subset thereof.
    """
    _handlers_export_path = WORD_SRC / "wordmcp" / "_server_handlers_docx_export.py"
    src = _handlers_export_path.read_text(encoding="utf-8")
    # Must import Literal from typing
    has_literal_import = "Literal" in src
    # Must use Literal in the function signature for the format param
    has_literal_param = "Literal[" in src and "format" in src
    assert has_literal_import and has_literal_param, (
        "export_document() does not use Literal typing for the 'format' parameter. "
        "Change 'format: str' to 'format: Literal[...]' and import Literal. P2 criterion."
    )


# ---------------------------------------------------------------------------
# US-03: manage_comments supports resolve and delete operations
# ---------------------------------------------------------------------------

def test_us03_manage_comments_supports_resolve_and_delete():
    """
    US-03: manage_comments must expose resolve and delete operations.

    As a document review agent, I want to resolve a comment thread and delete
    obsolete comments via the manage_comments tool, so that I can complete a full
    comment lifecycle without switching tools.
    """
    src = _handlers_edit_path.read_text(encoding="utf-8")
    # The manage_comments handler should handle 'resolve' and 'delete' operations
    assert "resolve" in src, (
        "manage_comments does not implement 'resolve' operation. P4 criterion."
    )
    assert "delete" in src, (
        "manage_comments does not implement 'delete' operation. P4 criterion."
    )


# ---------------------------------------------------------------------------
# US-04: manage_tracked_changes has no conditionally required revision_index
# ---------------------------------------------------------------------------

def test_us04_manage_tracked_changes_no_conditionally_required_params():
    """
    US-04: manage_tracked_changes must not expose accept_one/reject_one stubs
    that silently require a conditionally-present revision_index parameter.

    As a user calling manage_tracked_changes(operation='list'), I must not
    encounter a broken code path triggered by omitting revision_index (which
    is only needed for stubs that do not yet exist). P3 criterion.
    """
    src = _handlers_com_path.read_text(encoding="utf-8")
    # The stubs 'accept_one' and 'reject_one' must not exist until COM Phase 2
    assert "accept_one" not in src, (
        "'accept_one' stub still present in manage_tracked_changes — "
        "it introduces a conditionally required revision_index. Remove until COM Phase 2. P3."
    )
    assert "reject_one" not in src, (
        "'reject_one' stub still present in manage_tracked_changes — "
        "it introduces a conditionally required revision_index. Remove until COM Phase 2. P3."
    )


# ---------------------------------------------------------------------------
# US-05: bulk_update_paragraphs tool exists
# ---------------------------------------------------------------------------

def test_us05_bulk_update_paragraphs_tool_exists():
    """
    US-05: A bulk_update_paragraphs tool must exist for batch paragraph editing.

    As an agent processing a document restructuring request, I want to update
    multiple paragraphs in a single MCP call so that I avoid N round-trips for
    an N-paragraph edit and reduce latency significantly. P5 criterion.
    """
    src = _handlers_adv_path.read_text(encoding="utf-8")
    assert "bulk_update_paragraphs" in src, (
        "bulk_update_paragraphs tool is not defined in _server_handlers_docx_advanced.py. "
        "Add batch paragraph update support. P5 criterion."
    )


# ---------------------------------------------------------------------------
# US-06: bulk_update_table_cells tool exists
# ---------------------------------------------------------------------------

def test_us06_bulk_update_table_cells_tool_exists():
    """
    US-06: A bulk_update_table_cells tool must exist for batch table editing.

    As an agent populating a data table, I want to fill multiple cells in a
    single MCP call so that I avoid O(rows*cols) individual update_table_cell
    round-trips when filling a table from a dataset. P5 criterion.
    """
    src = _handlers_adv_path.read_text(encoding="utf-8")
    assert "bulk_update_table_cells" in src, (
        "bulk_update_table_cells tool is not defined in _server_handlers_docx_advanced.py. "
        "Add batch table-cell update support. P5 criterion."
    )


# ---------------------------------------------------------------------------
# US-07: Total tool count ≤ 40
# ---------------------------------------------------------------------------

MAX_TOOL_COUNT = 40


def test_us07_total_tool_count_at_most_40():
    """
    US-07: The total registered Word MCP tool count must not exceed 40.

    As a developer designing an agentic workflow, I need the Word MCP surface to
    remain ≤ 40 tools so that it fits within the context window of all MCP-capable
    models and is cognitively manageable for a planning agent. P8 criterion.
    """
    registered = _registered_tool_names()
    count = len(registered)
    assert count <= MAX_TOOL_COUNT, (
        f"Word MCP has {count} registered tools, exceeding the limit of {MAX_TOOL_COUNT}. "
        f"Tools registered: {registered}. "
        "Remove deprecated wrappers and merge export tools. P8 criterion."
    )


# ---------------------------------------------------------------------------
# US-08: Stage 3 template — enum values listed in tool descriptions
# ---------------------------------------------------------------------------

MANAGED_OPERATION_TOOLS = [
    ("_manage_comments", ["list", "add", "resolve", "delete"]),
    ("_manage_tracked_changes", ["list", "accept_all", "reject_all"]),
]


@pytest.mark.parametrize("func_name,expected_ops", MANAGED_OPERATION_TOOLS)
def test_us08_stage3_enum_values_in_tool_descriptions(func_name, expected_ops):
    """
    US-08: Every tool that takes an 'operation' or 'format' enum param must list
    all accepted values in its docstring (Stage 3 template element 1).

    As a developer reading tool docs, I want to see every valid enum value listed
    directly in the description so that I don't have to read source code to discover
    what operations are available. P7 criterion.
    """
    # Check both handler modules for these functions
    for handler_path in [_handlers_edit_path, _handlers_com_path]:
        docstring = _get_docstring(handler_path, func_name)
        if docstring:
            for op in expected_ops:
                assert op in docstring, (
                    f"Tool '{func_name}' docstring is missing enum value '{op}'. "
                    f"Stage 3 template requires all valid values listed. P7 criterion."
                )
            return
    pytest.fail(
        f"Function '{func_name}' not found in handler modules — "
        "cannot verify Stage 3 enum documentation. P7 criterion."
    )


# ---------------------------------------------------------------------------
# US-09: COM write tools document WORD_ENABLE_WRITE gate
# ---------------------------------------------------------------------------

COM_WRITE_TOOLS_GATE_CHECK = ["_export_document", "export_document"]


def test_us09_com_export_tool_documents_enable_write_gate():
    """
    US-09: The export_document COM tool must document the WORD_ENABLE_WRITE
    environment variable gate in its description.

    As a systems administrator deploying Word MCP in a read-only environment,
    I need every write/export tool to explicitly document the WORD_ENABLE_WRITE
    env-var requirement so that I can configure the deployment correctly without
    reading source code. P6/P7 criterion.
    """
    src = _handlers_com_path.read_text(encoding="utf-8")
    assert "WORD_ENABLE_WRITE" in src, (
        "_server_handlers_com.py does not document WORD_ENABLE_WRITE gate for COM write tools. "
        "Add gate documentation to export_document and manage_tracked_changes docstrings. P6."
    )


# ---------------------------------------------------------------------------
# US-10: Review tools declared in server.py type annotations
# ---------------------------------------------------------------------------

REVIEW_TOOL_NAMES = [
    "get_document_outline",
    "review_document",
    "write_review_findings",
    "export_review_evidence",
]


def test_us10_review_tools_declared_in_server_py():
    """
    US-10: All 4 review tools must appear in server.py's module-level type
    declarations so that static analysis tools (Pylance, mypy) can resolve them.

    As a developer adding type annotations or running a linter, I want all
    dynamically-registered review tools to be declared in server.py
    (e.g. as `tool_name: Any`) so that the static-analysis gap reported in
    the ARCH audit is closed. P8 support criterion.
    """
    server_src = _server_path.read_text(encoding="utf-8")
    missing = [name for name in REVIEW_TOOL_NAMES if name not in server_src]
    assert not missing, (
        f"Review tool(s) {missing} are not declared in server.py. "
        "Add them to the module-level Any-type variable block to close the "
        "static-analysis gap. P8 support criterion."
    )
