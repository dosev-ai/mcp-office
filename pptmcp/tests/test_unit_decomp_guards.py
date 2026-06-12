"""Decomposition guard tests — verify import resolution and invariants after pptmcp Red-file splits."""
from __future__ import annotations
import importlib
import inspect
import pathlib
import pytest

_SRC = pathlib.Path(__file__).parents[1] / "src" / "pptmcp"


def _src_lines(filename: str) -> int:
    return len((_SRC / filename).read_text(encoding="utf-8").splitlines())


# ─── Group 1: _review_detectors.py ───────────────────────────────────────────

def test_review_detectors_importable():
    importlib.import_module("pptmcp._review_detectors")


@pytest.mark.parametrize("name", [
    "_detect_png_issues", "_detect_text_overflow_heuristic", "_detect_shapes_outside_slide",
    "_detect_table_content_clipping", "_detect_unfinished_placeholder",
    "_detect_low_font_size", "_detect_sparse_slide",
])
def test_review_detectors_exports_detect_function(name):
    m = importlib.import_module("pptmcp._review_detectors")
    assert callable(getattr(m, name)), f"_review_detectors missing {name}"


@pytest.mark.parametrize("name", [
    "DEFAULT_TEXT", "_EMU_PER_INCH", "_PT_HIGH", "_PT_MEDIUM",
    "_SPARSE_FILL_THRESHOLD", "_SPARSE_COUNT_THRESHOLD", "_TEXT_TRUNCATION_LIMIT",
])
def test_review_detectors_exports_constant(name):
    m = importlib.import_module("pptmcp._review_detectors")
    assert hasattr(m, name), f"_review_detectors missing constant {name}"


# ─── Group 2: _review_export.py ──────────────────────────────────────────────

def test_review_export_importable():
    importlib.import_module("pptmcp._review_export")


def test_review_export_exports_review_slide_export():
    m = importlib.import_module("pptmcp._review_export")
    assert callable(m.review_slide_export)


# ─── Group 3: review_pptx.py shim ────────────────────────────────────────────

def test_review_pptx_shim_exports_review_slide_export():
    from pptmcp.review_pptx import review_slide_export  # noqa: F401
    assert callable(review_slide_export)


@pytest.mark.parametrize("name", [
    "_detect_table_content_clipping", "_detect_shapes_outside_slide",
    "_detect_text_overflow_heuristic", "_detect_low_font_size", "_detect_sparse_slide",
    "_detect_unfinished_placeholder", "_detect_png_issues", "DEFAULT_TEXT",
])
def test_review_pptx_shim_exports_symbol(name):
    import pptmcp.review_pptx as m
    assert hasattr(m, name), f"review_pptx shim missing {name}"


# ─── Group 4: _shapes_helpers.py ─────────────────────────────────────────────

def test_shapes_helpers_importable():
    importlib.import_module("pptmcp._shapes_helpers")


@pytest.mark.parametrize("name", [
    "_normalize_color_hex", "_apply_cell_text_style", "_apply_cell_fill",
    "_is_title_placeholder", "_is_empty_placeholder", "_shapes_overlap",
    "_remove_empty_overlapping_content_placeholders", "_calculate_overflow_risk",
    "_first_overlap_for_shape", "_EMPTY_PLACEHOLDER_PATTERNS",
])
def test_shapes_helpers_exports_symbol(name):
    m = importlib.import_module("pptmcp._shapes_helpers")
    assert hasattr(m, name), f"_shapes_helpers missing {name}"


# ─── Group 5: _shapes_text.py ────────────────────────────────────────────────

def test_shapes_text_importable():
    importlib.import_module("pptmcp._shapes_text")


@pytest.mark.parametrize("name", [
    "_A_R", "_A_BR", "_A_T", "_A_RPR", "_A_PPR", "_XML_SPACE",
    "_collect_para_text", "_rebuild_para_text",
    "set_speaker_notes", "edit_text_placeholder", "replace_slide_text",
    "set_text_format", "set_paragraph_format",
])
def test_shapes_text_exports_symbol(name):
    m = importlib.import_module("pptmcp._shapes_text")
    assert hasattr(m, name), f"_shapes_text missing {name}"


# ─── Group 6: _shapes_geometry.py ────────────────────────────────────────────

def test_shapes_geometry_importable():
    importlib.import_module("pptmcp._shapes_geometry")


@pytest.mark.parametrize("name", [
    "_style_table_shape", "add_textbox", "add_shape", "add_table_to_slide",
    "set_table_style", "set_shape_fill", "set_shape_geometry",
    "delete_shape", "insert_image", "detect_overlapping_shapes",
])
def test_shapes_geometry_exports_symbol(name):
    m = importlib.import_module("pptmcp._shapes_geometry")
    assert hasattr(m, name), f"_shapes_geometry missing {name}"


# ─── Group 6b: _shapes_overlap.py ───────────────────────────────────────────

def test_shapes_overlap_importable():
    importlib.import_module("pptmcp._shapes_overlap")


def test_shapes_overlap_exports_detect_overlapping_shapes():
    m = importlib.import_module("pptmcp._shapes_overlap")
    assert callable(m.detect_overlapping_shapes)


def test_shapes_geometry_still_exports_detect_overlapping_shapes():
    m = importlib.import_module("pptmcp._shapes_geometry")
    assert callable(m.detect_overlapping_shapes), "backwards-compat re-export missing"


# ─── Group 7: _shapes_links.py ───────────────────────────────────────────────

def test_shapes_links_importable():
    importlib.import_module("pptmcp._shapes_links")


@pytest.mark.parametrize("name", ["add_hyperlink", "manage_hyperlinks", "remove_empty_placeholders"])
def test_shapes_links_exports_symbol(name):
    m = importlib.import_module("pptmcp._shapes_links")
    assert hasattr(m, name), f"_shapes_links missing {name}"


def test_shapes_links_no_local_is_empty_placeholder_shadow():
    m = importlib.import_module("pptmcp._shapes_links")
    src = inspect.getsource(m.remove_empty_placeholders)
    assert "def _is_empty_placeholder" not in src, "Local shadow still present — BLOCKER-1 not resolved"


# ─── Group 8: shapes_pptx.py shim ────────────────────────────────────────────

@pytest.mark.parametrize("name", [
    "set_speaker_notes", "edit_text_placeholder", "replace_slide_text",
    "add_textbox", "add_shape", "add_table_to_slide", "set_table_style",
    "set_text_format", "set_paragraph_format", "set_shape_fill", "set_shape_geometry",
    "delete_shape", "add_hyperlink", "manage_hyperlinks", "insert_image",
    "remove_empty_placeholders", "detect_overlapping_shapes",
    "_calculate_overflow_risk", "_style_table_shape", "_EMPTY_PLACEHOLDER_PATTERNS",
])
def test_shapes_pptx_shim_exports_symbol(name):
    import pptmcp.shapes_pptx as m
    assert hasattr(m, name), f"shapes_pptx shim missing {name}"


# ─── Group 9: Layer separation (no xfail — enforce always) ───────────────────

@pytest.mark.parametrize("mod_name", [
    "pptmcp._review_detectors", "pptmcp._review_export",
    "pptmcp._shapes_helpers", "pptmcp._shapes_text",
    "pptmcp._shapes_geometry", "pptmcp._shapes_links",
])
def test_no_mcp_import_in_submodule(mod_name):
    try:
        mod = importlib.import_module(mod_name)
    except ImportError:
        pytest.skip(f"{mod_name} not yet created")
    src = pathlib.Path(mod.__file__).read_text(encoding="utf-8")
    assert "from mcp" not in src, f"{mod_name} must not import from mcp"
    assert "import fastmcp" not in src.lower(), f"{mod_name} must not import fastmcp"


# ─── Group 10: File size guards (no xfail — enforce after creation) ──────────

@pytest.mark.parametrize("filename,max_lines", [
    ("_review_detectors.py", 499),  # Amber after RCI-US-15 geometry-aware overflow
    ("_review_export.py", 349),
    ("_shapes_helpers.py", 349),
    ("_shapes_text.py", 349),
    ("_shapes_geometry.py", 349),  # GREEN after detect_overlapping_shapes extracted to _shapes_overlap
    ("_shapes_overlap.py", 349),
    ("_shapes_links.py", 349),
    ("_server_handlers_shape_ops.py", 499),  # Amber — new module from write.py decomp
])
def test_submodule_file_size_green(filename, max_lines):
    p = _SRC / filename
    if not p.exists():
        pytest.skip(f"{filename} not yet created")
    n = _src_lines(filename)
    assert n <= max_lines, f"{filename} is {n}L — exceeds Green threshold"


# ─── Group 10b: _server_handlers_shape_ops.py ────────────────────────────────

def test_shape_ops_importable():
    importlib.import_module("pptmcp._server_handlers_shape_ops")


@pytest.mark.parametrize("name", ["add_content", "set_format", "set_table_style", "shape"])
def test_shape_ops_exports_symbol(name):
    m = importlib.import_module("pptmcp._server_handlers_shape_ops")
    assert callable(getattr(m, name)), f"_server_handlers_shape_ops missing {name}"
