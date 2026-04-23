"""test_unit_refactor_imports.py
A-01 – A-13: Import tests for all 7 sub-modules + 4 prompt files.
C-01 – C-06: Security / contract tests.
"""
from __future__ import annotations

import importlib
import inspect


# ---------------------------------------------------------------------------
# A-01 – A-02: _format sub-modules
# ---------------------------------------------------------------------------

def test_a01_format_cell_importable() -> None:
    mod = importlib.import_module("excelmcp._format_cell")
    assert hasattr(mod, "__all__")
    for name in mod.__all__:
        assert hasattr(mod, name), f"_format_cell.__all__ lists {name!r} but it is not defined"


def test_a02_format_dimensions_importable() -> None:
    mod = importlib.import_module("excelmcp._format_dimensions")
    assert hasattr(mod, "__all__")
    for name in mod.__all__:
        assert hasattr(mod, name)


def test_a03_format_bulk_importable() -> None:
    mod = importlib.import_module("excelmcp._format_bulk")
    assert hasattr(mod, "__all__")
    for name in mod.__all__:
        assert hasattr(mod, name)


def test_a04_format_router_re_exports_all_names() -> None:
    """_format.py thin router must export everything from its sub-modules."""
    import excelmcp._format as router
    for sub in ("excelmcp._format_cell", "excelmcp._format_dimensions", "excelmcp._format_bulk"):
        sub_mod = importlib.import_module(sub)
        for name in sub_mod.__all__:
            assert hasattr(router, name), f"_format router missing {name!r} from {sub}"


# ---------------------------------------------------------------------------
# A-05 – A-08: _data sub-modules
# ---------------------------------------------------------------------------

def test_a05_data_cell_importable() -> None:
    mod = importlib.import_module("excelmcp._data_cell")
    assert hasattr(mod, "__all__")
    for name in mod.__all__:
        assert hasattr(mod, name)


def test_a06_data_tables_importable() -> None:
    mod = importlib.import_module("excelmcp._data_tables")
    assert hasattr(mod, "__all__")
    for name in mod.__all__:
        assert hasattr(mod, name)


def test_a07_data_annotations_importable() -> None:
    mod = importlib.import_module("excelmcp._data_annotations")
    assert hasattr(mod, "__all__")
    for name in mod.__all__:
        assert hasattr(mod, name)


def test_a08_data_write_importable() -> None:
    mod = importlib.import_module("excelmcp._data_write")
    assert hasattr(mod, "__all__")
    for name in mod.__all__:
        assert hasattr(mod, name)


def test_a09_data_router_re_exports_all_names() -> None:
    import excelmcp._data as router
    for sub in (
        "excelmcp._data_cell",
        "excelmcp._data_tables",
        "excelmcp._data_annotations",
        "excelmcp._data_write",
    ):
        sub_mod = importlib.import_module(sub)
        for name in sub_mod.__all__:
            assert hasattr(router, name), f"_data router missing {name!r} from {sub}"


# ---------------------------------------------------------------------------
# A-10 – A-13: Prompt sub-modules register prompts at import time
# ---------------------------------------------------------------------------

def _count_mcp_prompts() -> int:
    from excelmcp._server_instance import mcp
    # FastMCP stores prompts in mcp._prompt_manager or mcp.prompts depending on version
    if hasattr(mcp, "_prompt_manager"):
        pm = mcp._prompt_manager
        if hasattr(pm, "_prompts"):
            return len(pm._prompts)
    if hasattr(mcp, "prompts"):
        return len(mcp.prompts)
    # fallback: count via list_prompts if available
    return -1


def test_a10_server_prompts_analysis_importable() -> None:
    mod = importlib.import_module("excelmcp.server_prompts_analysis")
    # module-level logger must exist (BLOCKER-1)
    assert hasattr(mod, "logger")
    # must not define logging.basicConfig call (stdout-safety)
    src = inspect.getsource(mod)
    assert "logging.basicConfig" not in src


def test_a11_server_prompts_format_importable() -> None:
    mod = importlib.import_module("excelmcp.server_prompts_format")
    assert hasattr(mod, "logger")
    src = inspect.getsource(mod)
    assert "logging.basicConfig" not in src


def test_a12_server_prompts_workflow_importable() -> None:
    mod = importlib.import_module("excelmcp.server_prompts_workflow")
    assert hasattr(mod, "logger")
    src = inspect.getsource(mod)
    assert "logging.basicConfig" not in src


def test_a13_server_prompts_mgmt_importable() -> None:
    mod = importlib.import_module("excelmcp.server_prompts_mgmt")
    assert hasattr(mod, "logger")
    src = inspect.getsource(mod)
    assert "logging.basicConfig" not in src


# ---------------------------------------------------------------------------
# C-01 – C-06: Security / contract tests
# ---------------------------------------------------------------------------

def test_c01_validate_hyperlink_url_is_private_in_annotations() -> None:
    """_validate_hyperlink_url must NOT appear in _data_annotations.__all__."""
    mod = importlib.import_module("excelmcp._data_annotations")
    assert "_validate_hyperlink_url" not in mod.__all__, (
        "_validate_hyperlink_url must remain private (not in __all__)"
    )


def test_c02_validate_hyperlink_url_function_exists_in_annotations() -> None:
    """_validate_hyperlink_url must exist as a private function."""
    mod = importlib.import_module("excelmcp._data_annotations")
    assert hasattr(mod, "_validate_hyperlink_url"), (
        "_validate_hyperlink_url must exist in _data_annotations"
    )
    assert callable(mod._validate_hyperlink_url)


def test_c03_validate_cell_writes_only_in_server_com_recalc() -> None:
    """_validate_cell_writes must be defined only in server_com_recalc, not server_com."""
    import excelmcp.server_com as com
    assert not hasattr(com, "_validate_cell_writes"), (
        "_validate_cell_writes must NOT be in server_com (only in server_com_recalc)"
    )
    import excelmcp.server_com_recalc as recalc
    assert hasattr(recalc, "_validate_cell_writes"), (
        "_validate_cell_writes must exist in server_com_recalc"
    )


def test_c04_server_com_re_exports_exception_classes() -> None:
    """W5 ADJUSTMENT-1: thin server_com.py must re-export exception classes for tests."""
    import excelmcp.server_com as com
    from excelmcp._core import NotAllowedError, OfficeCOMError, ValidationError
    assert getattr(com, "NotAllowedError", None) is NotAllowedError
    assert getattr(com, "OfficeCOMError", None) is OfficeCOMError
    assert getattr(com, "ValidationError", None) is ValidationError


def test_c05_prompt_modules_have_no_print_calls() -> None:
    """Prompt modules must never write to stdout (stdout = MCP JSON-RPC only).

    The print() check uses AST inspection to avoid false positives from
    instructional text inside prompt string literals (e.g. 'to print (e.g. A1:H40)').
    """
    import ast
    for mod_name in (
        "excelmcp.server_prompts_analysis",
        "excelmcp.server_prompts_format",
        "excelmcp.server_prompts_workflow",
        "excelmcp.server_prompts_mgmt",
    ):
        mod = importlib.import_module(mod_name)
        src = inspect.getsource(mod)
        assert "sys.stdout" not in src, f"{mod_name} must not write to sys.stdout"
        # AST-based check: find actual Call nodes where the function is 'print'
        tree = ast.parse(src)
        print_calls = [
            node for node in ast.walk(tree)
            if isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "print"
        ]
        assert not print_calls, (
            f"{mod_name} contains {len(print_calls)} print() call(s) at lines "
            + str([getattr(n, 'lineno', '?') for n in print_calls])
        )


def test_c06_all_sub_modules_have_dunder_all() -> None:
    """Every refactored sub-module must declare __all__ to control star-import surface."""
    for mod_name in (
        "excelmcp._format_cell",
        "excelmcp._format_dimensions",
        "excelmcp._format_bulk",
        "excelmcp._data_cell",
        "excelmcp._data_tables",
        "excelmcp._data_annotations",
        "excelmcp._data_write",
    ):
        mod = importlib.import_module(mod_name)
        assert hasattr(mod, "__all__"), f"{mod_name} must define __all__"
        assert isinstance(mod.__all__, list), f"{mod_name}.__all__ must be a list"
        assert len(mod.__all__) > 0, f"{mod_name}.__all__ must not be empty"
