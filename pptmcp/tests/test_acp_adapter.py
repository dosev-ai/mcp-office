"""Unit tests for pptmcp ACP adapter (build_presentation_context).

Covers security invariants, ACP level contracts, and field correctness.
All external I/O is mocked — no real presentation files needed.
"""
from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from pptmcp._pptx_acp_adapter import build_presentation_context
from pptmcp.presentation_pptx import ValidationError

# ---------------------------------------------------------------------------
# Shared fake data
# ---------------------------------------------------------------------------

_PATH = "/allowed/deck.pptx"

# patch targets (the module where names are *used*)
_CP = "pptmcp._pptx_acp_adapter._check_path"
_LP = "pptmcp._pptx_acp_adapter._load_prs"
_ST = "pptmcp._pptx_acp_adapter._slide_title"
_VA = "pptmcp._pptx_acp_adapter.validate_acp_annotations"


def _make_fake_resolved() -> MagicMock:
    """Return a mock Path-like object as returned by _check_path."""
    resolved = MagicMock(spec=Path)
    resolved.name = "deck.pptx"
    resolved.stem = "deck"
    stat_obj = MagicMock()
    stat_obj.st_mtime = 1_700_000_000.0  # fixed timestamp
    resolved.stat.return_value = stat_obj
    return resolved


def _make_fake_prs(slide_count: int = 2) -> MagicMock:
    """Return a mock Presentation with *slide_count* slides."""
    prs = MagicMock()
    prs.slides = [MagicMock() for _ in range(slide_count)]
    return prs


def _call(level: str = "focused", annotations=None) -> dict:
    """Call build_presentation_context with all I/O stubbed."""
    with (
        patch(_CP, return_value=_make_fake_resolved()),
        patch(_LP, return_value=_make_fake_prs()),
        patch(_ST, return_value="Slide Title"),
        patch(_VA, return_value=[]),
    ):
        return build_presentation_context(_PATH, level, annotations)


# ---------------------------------------------------------------------------
# L-004 tests
# ---------------------------------------------------------------------------

@pytest.mark.unit
def test_level_index_returns_required_fields():
    result = _call("index")
    for key in ("acp_version", "artifact_type", "tool_name", "artifact_id", "summary", "timestamp"):
        assert key in result, f"Missing required ACPIndex field: {key!r}"


@pytest.mark.unit
def test_level_focused_includes_content():
    result = _call("focused")
    assert "content" in result
    content = result["content"]
    assert "slide_titles" in content or "slide_count" in content


@pytest.mark.unit
def test_level_deep_includes_detail():
    result = _call("deep")
    assert "detail" in result


@pytest.mark.unit
def test_path_check_called_before_file_io():
    """_check_path() must be the first call — ValidationError must propagate."""
    with (
        patch(_CP, side_effect=ValidationError("not allowed")),
        patch(_LP, side_effect=AssertionError("_load_prs must not be called")),
    ):
        with pytest.raises(ValidationError, match="not allowed"):
            build_presentation_context(_PATH, "focused")


@pytest.mark.unit
def test_invalid_annotations_rejected():
    """Annotations with injection-style values must raise ValueError."""
    bad_annotations = [{"key": "=injection"}]
    with (
        patch(_CP, return_value=_make_fake_resolved()),
        patch(_LP, return_value=_make_fake_prs()),
        patch(_ST, return_value="Slide Title"),
        patch(_VA, return_value=["injection detected"]),  # validator flags it
    ):
        with pytest.raises(ValueError, match="Invalid ACP annotations"):
            build_presentation_context(_PATH, "deep", bad_annotations)


@pytest.mark.unit
def test_artifact_type_is_presentation():
    result = _call("index")
    assert result["artifact_type"] == "presentation"


@pytest.mark.unit
def test_acp_version_is_1_0():
    result = _call("index")
    assert result["acp_version"] == "1.0"


@pytest.mark.unit
def test_no_cell_values_in_deep_detail():
    """PPT detail must not expose cell values (defence-in-depth against S-001 analogue)."""
    result = _call("deep")
    assert "cell_values" not in result.get("detail", {})


@pytest.mark.unit
def test_slide_titles_formula_prefix_sanitized():
    """H-003: slide titles with formula-injection prefixes must be prefix-escaped."""
    with (
        patch(_CP, return_value=_make_fake_resolved()),
        patch(_LP, return_value=_make_fake_prs(slide_count=2)),
        patch(_ST, side_effect=["=HYPERLINK(\"http://evil.com\",\"Click\")", "Normal Title"]),
        patch(_VA, return_value=[]),
    ):
        result = build_presentation_context(_PATH, "focused")

    titles = result["content"]["slide_titles"]
    assert titles[0] == "'=HYPERLINK(\"http://evil.com\",\"Click\")", (
        f"Expected prefix-escaped title but got {titles[0]!r}"
    )
    assert titles[1] == "Normal Title", (
        f"Normal title should be unchanged but got {titles[1]!r}"
    )


@pytest.mark.unit
def test_png_paths_symlink_excluded():
    """BLOCKER-3: glob results that resolve outside export_dir are filtered out."""
    from pathlib import Path as RealPath

    export_dir_real = RealPath("/allowed/deck_exports")

    # Safe PNG — resolves INSIDE export_dir
    safe_png = MagicMock()
    safe_png.__str__ = MagicMock(return_value="/allowed/deck_exports/slide_01.png")
    safe_png.resolve.return_value = RealPath("/allowed/deck_exports/slide_01.png")

    # Evil PNG — resolves OUTSIDE export_dir (symlink pointing to /etc/passwd)
    evil_png = MagicMock()
    evil_png.__str__ = MagicMock(return_value="/allowed/deck_exports/evil_link.png")
    evil_png.resolve.return_value = RealPath("/etc/passwd")

    # Fake candidate (the export directory)
    candidate = MagicMock(spec=RealPath)
    candidate.is_dir.return_value = True
    candidate.resolve.return_value = export_dir_real
    candidate.glob.return_value = [evil_png, safe_png]

    # Fake Path("/allowed") that returns candidate when divided by "deck_exports"
    roots_path = MagicMock(spec=RealPath)
    roots_path.__truediv__ = MagicMock(return_value=candidate)

    fake_resolved = _make_fake_resolved()
    fake_resolved.stem = "deck"

    def _fake_path_constructor(*args):
        if args and str(args[0]) == "/allowed":
            return roots_path
        return RealPath(*args)

    with (
        patch(_CP, return_value=fake_resolved),
        patch(_LP, return_value=_make_fake_prs()),
        patch(_ST, return_value="Slide Title"),
        patch(_VA, return_value=[]),
        patch("pptmcp._pptx_acp_adapter.os.getenv", return_value="/allowed"),
        patch("pptmcp._pptx_acp_adapter.Path", side_effect=_fake_path_constructor),
    ):
        result = build_presentation_context(_PATH, "deep")

    png_paths = result.get("detail", {}).get("png_paths", [])
    # The evil symlink (resolves to /etc/passwd) must NOT appear
    assert not any("passwd" in p or "evil_link" in p for p in png_paths), (
        f"Symlink pointing outside export_dir leaked into png_paths: {png_paths}"
    )
    # The safe PNG must appear
    assert any("slide_01" in p for p in png_paths), (
        f"Safe PNG was incorrectly filtered out: {png_paths}"
    )


@pytest.mark.unit
def test_invalid_level_raises_value_error():
    """GAP-8: invalid level string must raise ValueError before any I/O."""
    with pytest.raises(ValueError, match="Invalid level"):
        build_presentation_context(_PATH, "invalid_level")  # type: ignore[arg-type]


@pytest.mark.unit
def test_filename_formula_prefix_sanitized():
    """GAP-6 (security): filename starting with '=' must be prefix-escaped in artifact_id and summary."""
    fake_resolved = _make_fake_resolved()
    fake_resolved.name = "=dangerous.pptx"
    fake_resolved.stem = "=dangerous"
    with (
        patch(_CP, return_value=fake_resolved),
        patch(_LP, return_value=_make_fake_prs()),
        patch(_ST, return_value="Title"),
        patch(_VA, return_value=[]),
    ):
        result = build_presentation_context(_PATH, "index")

    artifact_id = result.get("artifact_id", "")
    summary = result.get("summary", "")
    assert not artifact_id.startswith("deck:="), (
        f"Formula-prefix filename leaked into artifact_id: {artifact_id!r}"
    )
    assert "'=" in artifact_id or "'=" in summary, (
        f"Expected apostrophe prefix-escape in artifact_id or summary. "
        f"artifact_id={artifact_id!r}, summary={summary!r}"
    )


@pytest.mark.unit
def test_slide_title_none_appends_empty_string():
    """GAP-7: when _slide_title returns None, '' is appended (not raised, not sanitized)."""
    with (
        patch(_CP, return_value=_make_fake_resolved()),
        patch(_LP, return_value=_make_fake_prs(slide_count=2)),
        patch(_ST, side_effect=[None, "Normal"]),
        patch(_VA, return_value=[]),
    ):
        result = build_presentation_context(_PATH, "focused")

    titles = result["content"]["slide_titles"]
    assert titles[0] == "", f"Expected empty string for None title, got {titles[0]!r}"
    assert titles[1] == "Normal", f"Expected 'Normal', got {titles[1]!r}"


@pytest.mark.unit
def test_artifact_id_starts_with_deck_prefix():
    result = _call("index")
    assert result["artifact_id"].startswith("deck:")


@pytest.mark.unit
def test_deck_prefix_constant_value():
    _wfr = pytest.importorskip("workflowruntime.contracts.constants")
    ARTIFACT_ID_PREFIX_DECK = _wfr.ARTIFACT_ID_PREFIX_DECK

    assert ARTIFACT_ID_PREFIX_DECK == "deck:"


@pytest.mark.unit
def test_pptmcp_deck_prefix_matches_workflowruntime_constant():
    """W5 parity: _ARTIFACT_ID_PREFIX_DECK (pptmcp) == ARTIFACT_ID_PREFIX_DECK (workflowruntime)."""
    _wfr = pytest.importorskip("workflowruntime.contracts.constants")
    from pptmcp._pptx_acp_adapter import _ARTIFACT_ID_PREFIX_DECK as pptmcp_prefix
    wfr_prefix = _wfr.ARTIFACT_ID_PREFIX_DECK

    assert pptmcp_prefix == wfr_prefix
