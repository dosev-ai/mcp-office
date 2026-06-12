"""Capabilities and parameter schema for pptmcp."""
from __future__ import annotations

import inspect
from typing import Any

__all__: list[str] = []

COMPACT_TOOL_SURFACE_ENV = "PPT_COMPACT_TOOL_SURFACE"

COMPACT_TOOL_NAMES: tuple[str, ...] = (
    "capabilities",
    "read_presentation",
    "list_layouts",
    "list_slides",
    "read_slide",
    "list_shapes",
    "get_shape",
    "extract_tables",
    "export_slide_as_text",
    "extract_presentation_text",
    "detect_overlapping_shapes",
    "declare_slide_contract",
    "validate_contract",
    "check_presentation_against_contract",
    "review_slide_export",
    "produce_evidence_bundle",
    "get_presentation_context",
    "add_content",
    "set_format",
    "slide",
    "shape",
    "export",
    "manage_comments",
    "batch_set_text",
    "pptmcp_add_column",
    "pptmcp_edit_table_cell",
    "pptmcp_set_column_width",
)

TOOL_BUNDLES: dict[str, list[str]] = {
    "inspect": [
        "read_presentation",
        "list_layouts",
        "list_slides",
        "read_slide",
        "list_shapes",
        "get_shape",
        "get_presentation_context",
    ],
    "author": ["shape", "set_format", "slide"],
    "table_qa": [
        "extract_tables",
        "export_slide_as_text",
        "extract_presentation_text",
        "detect_overlapping_shapes",
    ],
    "render_check_iterate": [
        "declare_slide_contract",
        "validate_contract",
        "check_presentation_against_contract",
        "export",
        "review_slide_export",
        "produce_evidence_bundle",
    ],
}


# ---------------------------------------------------------------------------
# Introspection helper
# ---------------------------------------------------------------------------

def _param_schema(fn) -> list[dict]:  # type: ignore[type-arg]
    """Return [{name, type, required}] for fn using inspect.signature.

    Skips self, *args, **kwargs.
    type is the annotation string; 'Any' if unannotated.
    required=True when no default value is set.
    """
    sig = inspect.signature(fn)
    result = []
    for name, param in sig.parameters.items():
        if name == "self":
            continue
        if param.kind in (
            inspect.Parameter.VAR_POSITIONAL,
            inspect.Parameter.VAR_KEYWORD,
        ):
            continue
        ann = param.annotation
        if ann is inspect.Parameter.empty:
            type_str = "Any"
        else:
            type_str = str(ann)
            # Normalise <class 'typename'> for runtime-evaluated annotations
            if type_str.startswith("<class '") and type_str.endswith("'>"):
                type_str = type_str[8:-2]
            # Strip surrounding quotes from string-literal annotations
            # e.g. "'int | None'" (from `param: "int | None"` + __future__ annotations)
            if len(type_str) >= 2 and (
                (type_str[0] == "'" and type_str[-1] == "'")
                or (type_str[0] == '"' and type_str[-1] == '"')
            ):
                type_str = type_str[1:-1]
        required = param.default is inspect.Parameter.empty
        result.append({"name": name, "type": type_str, "required": required})
    return result


# ---------------------------------------------------------------------------
# Capabilities manifest
# ---------------------------------------------------------------------------

def capabilities() -> dict:
    """Return metadata about this MCP phase and its available tools.

    Each entry in 'tools' and 'com_tools' is a dict::

        {"tool": "<mcp_tool_name>", "params": [{"name": ..., "type": ..., "required": ...}]}

    'tools' covers all always-registered tools (python-pptx + always-on COM wrappers).
    'com_tools' covers platform-conditional COM tools (Windows + PPT_ENABLE_COM=true).
    """
    # Lazy imports avoid circular-import at module load time.
    # (shapes_pptx, content_pptx, review_pptx, contract_pptx all import from here.)
    import pptmcp.review_pptx as _review  # noqa: PLC0415
    from pptmcp import content_pptx as _content  # noqa: PLC0415
    from pptmcp import contract_pptx as _contract  # noqa: PLC0415
    from pptmcp import presentation_pptx as _prs  # noqa: PLC0415
    from pptmcp import shapes_pptx as _shapes  # noqa: PLC0415
    # presentation_com is importable on all platforms (win32com imported lazily inside fns)
    try:
        from pptmcp import presentation_com as _pcom  # noqa: PLC0415
    except Exception:  # pragma: no cover
        _pcom = None  # type: ignore[assignment]

    def _e(tool_name: str, fn: Any) -> dict:  # type: ignore[return]
        return {"tool": tool_name, "params": _param_schema(fn)}

    # 46 always-registered tools (pptx-layer + always-on COM wrappers + D1-D3 dispatch)
    tools = [
        _e("capabilities", capabilities),
        _e("read_presentation", _prs.read_presentation),
        _e("get_presentation_metadata", _prs.get_presentation_metadata),
        _e("list_slides", _prs.list_slides),
        _e("list_layouts", _prs.list_layouts),
        _e("read_slide", _prs.read_slide),
        _e("read_speaker_notes", _prs.read_speaker_notes),
        _e("list_shapes", _prs.list_shapes),
        _e("get_shape", _prs.get_shape),
        _e("extract_tables", _content.extract_tables),
        _e("extract_images", _content.extract_images),
        _e("create_presentation", _prs.create_presentation),
        _e("add_slide", _prs.add_slide),
        _e("edit_text_placeholder", _shapes.edit_text_placeholder),
        _e("set_speaker_notes", _shapes.set_speaker_notes),
        _e("replace_slide_text", _shapes.replace_slide_text),
        _e("delete_shape", _shapes.delete_shape),
        _e("insert_image", _shapes.insert_image),
        _e("reorder_slides", _prs.reorder_slides),
        _e("delete_slide", _prs.delete_slide),
        _e("export_slide_as_text", _content.export_slide_as_text),
        _e("save", _prs.save),
        _e("extract_presentation_text", _content.extract_presentation_text),
        _e("manage_hyperlinks", _shapes.manage_hyperlinks),
        {"tool": "add_content", "params": [
            {"name": "content_type", "type": "str", "required": True},
            {"name": "path", "type": "str", "required": True},
            {"name": "slide_index", "type": "int", "required": True},
            {"name": "left", "type": "float", "required": True},
            {"name": "top", "type": "float", "required": True},
            {"name": "width", "type": "float", "required": True},
            {"name": "height", "type": "float", "required": True},
            {"name": "text", "type": "str", "required": False},
            {"name": "confirm", "type": "bool", "required": False},
            {"name": "shape_type", "type": "str", "required": False},
            {"name": "rows", "type": "int", "required": False},
            {"name": "cols", "type": "int", "required": False},
            {"name": "data", "type": "list[list[str]] | None", "required": False},
            {"name": "font_size_pt", "type": "float | None", "required": False},
            {"name": "header_font_size_pt", "type": "float | None", "required": False},
            {"name": "header_bold", "type": "bool | None", "required": False},
            {"name": "row_height_pt", "type": "float | None", "required": False},
            {"name": "suppress_content_placeholder", "type": "bool", "required": False},
        ]},
        {"tool": "set_format", "params": [
            {"name": "target", "type": "str", "required": True},
            {"name": "path", "type": "str", "required": True},
            {"name": "slide_index", "type": "int", "required": True},
            {"name": "shape_id", "type": "int", "required": True},
            {"name": "confirm", "type": "bool", "required": False},
            {"name": "bold", "type": "bool | None", "required": False},
            {"name": "italic", "type": "bool | None", "required": False},
            {"name": "font_size_pt", "type": "float | None", "required": False},
            {"name": "font_name", "type": "str | None", "required": False},
            {"name": "color_hex", "type": "str | None", "required": False},
            {"name": "paragraph_index", "type": "int", "required": False},
            {"name": "alignment", "type": "str | None", "required": False},
            {"name": "line_spacing", "type": "float | None", "required": False},
            {"name": "space_before_pt", "type": "float | None", "required": False},
            {"name": "space_after_pt", "type": "float | None", "required": False},
        ]},
        _e("set_table_style", _shapes.set_table_style),
        _e("copy_slide", _prs.copy_slide),
        _e("add_hyperlink", _shapes.add_hyperlink),
        _e("detect_overlapping_shapes", _shapes.detect_overlapping_shapes),
        _e("clear_slide_content", _prs.clear_slide_content),
        _e("apply_slide_layout", _prs.apply_slide_layout),
        _e("remove_empty_placeholders", _shapes.remove_empty_placeholders),
        _e("review_slide_export", _review.review_slide_export),
        _e("declare_slide_contract", _contract.declare_slide_contract),
        _e("validate_contract", _contract.validate_contract),
        _e("check_presentation_against_contract", _contract.check_presentation_against_contract),
        _e("produce_evidence_bundle", _content.produce_evidence_bundle),
        # _server_handlers_acp.py (1 — ACP context tool)
        {"tool": "get_presentation_context", "params": [
            {"name": "path", "type": "str", "required": True},
            {"name": "level", "type": "str", "required": False},
            {"name": "annotations", "type": "list[dict] | None", "required": False},
        ]},
        # Consolidated dispatch tools (D1-D3 Sprint Consolidate-D)
        {"tool": "slide", "params": [
            {"name": "operation", "type": "str", "required": True},
            {"name": "path", "type": "str", "required": True},
            {"name": "confirm", "type": "bool", "required": False},
            {"name": "layout_index", "type": "int", "required": False},
            {"name": "title", "type": "str | None", "required": False},
            {"name": "layout_name", "type": "str | None", "required": False},
            {"name": "slide_index", "type": "int | None", "required": False},
            {"name": "new_order", "type": "list[int] | None", "required": False},
            {"name": "source_path", "type": "str | None", "required": False},
            {"name": "source_slide_index", "type": "int | None", "required": False},
            {"name": "target_path", "type": "str | None", "required": False},
            {"name": "target_slide_index", "type": "int | None", "required": False},
            {"name": "suppress_content_placeholder", "type": "bool", "required": False},
        ]},
        {"tool": "shape", "params": [
            {"name": "operation", "type": "str", "required": True},
            {"name": "path", "type": "str", "required": True},
            {"name": "slide_index", "type": "int", "required": True},
            {"name": "confirm", "type": "bool", "required": False},
            {"name": "left", "type": "float", "required": False},
            {"name": "top", "type": "float", "required": False},
            {"name": "width", "type": "float", "required": False},
            {"name": "height", "type": "float", "required": False},
            {"name": "text", "type": "str", "required": False},
            {"name": "shape_type", "type": "str", "required": False},
            {"name": "rows", "type": "int", "required": False},
            {"name": "cols", "type": "int", "required": False},
            {"name": "data", "type": "list[list[str]] | None", "required": False},
            {"name": "shape_id", "type": "int | None", "required": False},
            {"name": "target", "type": "str", "required": False},
            {"name": "bold", "type": "bool | None", "required": False},
            {"name": "italic", "type": "bool | None", "required": False},
            {"name": "font_size_pt", "type": "float | None", "required": False},
            {"name": "font_name", "type": "str | None", "required": False},
            {"name": "color_hex", "type": "str | None", "required": False},
            {"name": "paragraph_index", "type": "int", "required": False},
            {"name": "alignment", "type": "str | None", "required": False},
            {"name": "line_spacing", "type": "float | None", "required": False},
            {"name": "space_before_pt", "type": "float | None", "required": False},
            {"name": "space_after_pt", "type": "float | None", "required": False},
            {"name": "header_font_size_pt", "type": "float | None", "required": False},
            {"name": "header_bold", "type": "bool | None", "required": False},
            {"name": "row_height_pt", "type": "float | None", "required": False},
            {"name": "suppress_content_placeholder", "type": "bool", "required": False},
            {"name": "header_color_hex", "type": "str | None", "required": False},
            {"name": "fill_color_hex", "type": "str | None", "required": False},
            {"name": "header_fill_color_hex", "type": "str | None", "required": False},
            {"name": "header_row_height_pt", "type": "float | None", "required": False},
            {"name": "new_left_in", "type": "float | None", "required": False},
            {"name": "new_top_in", "type": "float | None", "required": False},
            {"name": "new_width_in", "type": "float | None", "required": False},
            {"name": "new_height_in", "type": "float | None", "required": False},
        ]},
        {"tool": "export", "params": [
            {"name": "scope", "type": "str", "required": True},
            {"name": "path", "type": "str", "required": True},
            {"name": "fmt", "type": "str", "required": False},
            {"name": "slide_index", "type": "int | None", "required": False},
            {"name": "output_path", "type": "str | None", "required": False},
            {"name": "dpi", "type": "int | None", "required": False},
            {"name": "width_px", "type": "int", "required": False},
            {"name": "height_px", "type": "int", "required": False},
            {"name": "return_inline", "type": "bool", "required": False},
            {"name": "response_mode", "type": "str", "required": False},
            {"name": "max_inline_bytes", "type": "int", "required": False},
            {"name": "options", "type": "dict | None", "required": False},
            {"name": "confirm", "type": "bool", "required": False},
        ]},
        # Always-registered COM wrapper tools (return graceful error when COM unavailable)
        {"tool": "export_slide", "params": _param_schema(_pcom.ppt_export_slide) if _pcom else []},
        {"tool": "export_deck", "params": _param_schema(_pcom.ppt_export_deck) if _pcom else []},
        {"tool": "export_changed_slides_only", "params": _param_schema(_pcom.export_changed_slides_only) if _pcom else []},
        # H-003: hardcoded to match server.py wrapper (adds slide_indices/width_px/height_px)
        {"tool": "export_slides_to_stamped_dir", "params": [
            {"name": "path", "type": "str", "required": True},
            {"name": "base_dir", "type": "str", "required": True},
            {"name": "slide_indices", "type": "list[int] | None", "required": False},
            {"name": "width_px", "type": "int", "required": False},
            {"name": "height_px", "type": "int", "required": False},
            {"name": "confirm", "type": "bool", "required": False},
        ]},
        {"tool": "manage_comments", "params": [
            {"name": "path", "type": "str", "required": True},
            {"name": "operation", "type": "str", "required": True},
            {"name": "slide_index", "type": "int | None", "required": False},
            {"name": "comment_id", "type": "int | None", "required": False},
            {"name": "text", "type": "str | None", "required": False},
            {"name": "author", "type": "str", "required": False},
            {"name": "slide_x_emu", "type": "int", "required": False},
            {"name": "slide_y_emu", "type": "int", "required": False},
            {"name": "confirm", "type": "bool", "required": False},
        ]},
        {"tool": "batch_set_text", "params": [
            {"name": "path", "type": "str", "required": True},
            {"name": "slide_index", "type": "int", "required": True},
            {"name": "updates", "type": "list[dict]", "required": True},
            {"name": "confirm", "type": "bool", "required": False},
        ]},
        {"tool": "pptmcp_add_column", "params": [
            {"name": "path", "type": "str", "required": True},
            {"name": "slide_index", "type": "int", "required": True},
            {"name": "shape_id", "type": "int", "required": True},
            {"name": "width_inches", "type": "float", "required": False},
            {"name": "confirm", "type": "bool", "required": False},
        ]},
        {"tool": "pptmcp_edit_table_cell", "params": [
            {"name": "path", "type": "str", "required": True},
            {"name": "slide_index", "type": "int", "required": True},
            {"name": "shape_id", "type": "int", "required": True},
            {"name": "row", "type": "int", "required": True},
            {"name": "col", "type": "int", "required": True},
            {"name": "text", "type": "str", "required": True},
            {"name": "confirm", "type": "bool", "required": False},
        ]},
        {"tool": "pptmcp_set_column_width", "params": [
            {"name": "path", "type": "str", "required": True},
            {"name": "slide_index", "type": "int", "required": True},
            {"name": "shape_id", "type": "int", "required": True},
            {"name": "col", "type": "int", "required": True},
            {"name": "width_inches", "type": "float", "required": True},
            {"name": "confirm", "type": "bool", "required": False},
        ]},
    ]

    # 2 platform-conditional COM tools (registered only when _COM_AVAILABLE on Windows)
    com_tools = [
        {"tool": "run_slide_show", "params": _param_schema(_pcom.run_slide_show) if _pcom else []},
        {"tool": "recalculate_charts", "params": _param_schema(_pcom.recalculate_charts) if _pcom else []},
    ]

    # v2 structured fields \u2014 additive, nested under capabilities_v2 key
    from pptmcp._server_manifest import build_capabilities_v2  # noqa: PLC0415
    from pptmcp._server_compact import compact_surface_enabled, COMPACT_TOOL_NAMES as _CTN  # noqa: PLC0415
    _active = list(_CTN) if compact_surface_enabled() else None
    _v2 = build_capabilities_v2(None, active_tools=_active)

    return {
        "phase": "2.1",
        "backend": "python-pptx",
        "tools": tools,
        "com_tools": com_tools,
        "compact_tool_surface": {
            "enabled_by_env": COMPACT_TOOL_SURFACE_ENV,
            "tool_count": len(COMPACT_TOOL_NAMES),
            "tools": list(COMPACT_TOOL_NAMES),
            "description": (
                "Set PPT_COMPACT_TOOL_SURFACE=true to register a dispatcher-first "
                "agent surface instead of the full backward-compatible tool list."
            ),
        },
        "tool_bundles": TOOL_BUNDLES,
        "governance": {
            "allowlist_roots_env": "PPT_ALLOWLIST_ROOTS",
            "enable_write_env": "PPT_ENABLE_WRITE",
            "enable_com_env": "PPT_ENABLE_COM",
        },
        "recommended_workflow": "Output Contract \u2192 Build \u2192 Export (PNG/PDF) \u2192 Self-Review \u2192 Iterate \u2192 Evidence Bundle",
        "resources": ["ppt://prompts/ppt_render_check_iterate_v1"],
        "prompt_name": "ppt_render_check_iterate_v1",
        "capabilities_v2": _v2,
    }
