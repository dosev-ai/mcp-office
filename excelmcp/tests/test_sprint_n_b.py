"""
Sprint N — Commit B and Commit C tests.

US-N06–N10: bulk hyperlinks, bulk comments, registry checks
US-N11–N14: data_validation unification

No COM required — openpyxl only.
"""
import openpyxl
import pytest

from excelmcp._core import _wb_cache
from excelmcp._core import ValidationError, NotAllowedError

from ._smoke_helpers import _get_capabilities_snapshot, _get_tool_names


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_workbook(tmp_path, monkeypatch, *, name="test.xlsx"):
    """Create a minimal xlsx in tmp_path, configure env, return str path."""
    monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
    monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
    path = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    wb.save(str(path))
    return str(path)
class TestCommitBHyperlinks:
    """bulk_add_hyperlinks: single-link, 500-item cap, gate tests."""

    @pytest.mark.unit
    def test_bulk_add_hyperlinks_single_link(self, tmp_path, monkeypatch):
        """Single link via list-of-1 succeeds (US-N06)."""
        from excelmcp._data import bulk_add_hyperlinks

        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        result = bulk_add_hyperlinks(
            path, "Sheet1",
            links=[{"address": "A1", "url": "https://example.com"}],
            confirm=True,
        )
        assert result.get("ok") is True
        assert result.get("links_added") == 1
        assert result.get("sheet") == "Sheet1"

    @pytest.mark.unit
    def test_bulk_add_hyperlinks_cap_501_raises(self, tmp_path, monkeypatch):
        """501 links raises ValidationError (US-N07)."""
        from excelmcp._data import bulk_add_hyperlinks

        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        links = [{"address": "A1", "url": "https://example.com"} for _ in range(501)]
        with pytest.raises(ValidationError, match="500"):
            bulk_add_hyperlinks(path, "Sheet1", links=links, confirm=True)

    @pytest.mark.unit
    def test_bulk_add_hyperlinks_cap_500_passes(self, tmp_path, monkeypatch):
        """Exactly 500 links passes the cap boundary (US-N07)."""
        from excelmcp._data import bulk_add_hyperlinks
        import openpyxl as _openpyxl

        # Create a workbook with enough rows
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        path = tmp_path / "cap500.xlsx"
        wb_obj = _openpyxl.Workbook()
        wb_obj.active.title = "Sheet1"
        wb_obj.save(str(path))
        _wb_cache.clear()

        links = [{"address": f"A{i + 1}", "url": "https://example.com"} for i in range(500)]
        result = bulk_add_hyperlinks(str(path), "Sheet1", links=links, confirm=True)
        assert result.get("ok") is True
        assert result.get("links_added") == 500

    @pytest.mark.unit
    def test_bulk_add_hyperlinks_write_gate(self, tmp_path, monkeypatch):
        """EXCEL_ENABLE_WRITE unset → NotAllowedError (US-N17)."""
        from excelmcp._data import bulk_add_hyperlinks

        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        monkeypatch.delenv("EXCEL_ENABLE_WRITE", raising=False)
        with pytest.raises(NotAllowedError):
            bulk_add_hyperlinks(
                path, "Sheet1",
                links=[{"address": "A1", "url": "https://example.com"}],
                confirm=True,
            )

    @pytest.mark.unit
    def test_bulk_add_hyperlinks_confirm_gate(self, tmp_path, monkeypatch):
        """confirm=False → ValidationError (US-N18)."""
        from excelmcp._data import bulk_add_hyperlinks

        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        with pytest.raises(ValidationError):
            bulk_add_hyperlinks(
                path, "Sheet1",
                links=[{"address": "A1", "url": "https://example.com"}],
                confirm=False,
            )


# ---------------------------------------------------------------------------
# TestCommitBComments (US-N08, US-N09)
# ---------------------------------------------------------------------------

class TestCommitBComments:
    """bulk_add_comments: sheet key in return dict, single comment, gate tests."""

    @pytest.mark.unit
    def test_bulk_add_comments_sheet_key_present(self, tmp_path, monkeypatch):
        """Result dict has 'sheet' key (US-N08)."""
        from excelmcp._data import bulk_add_comments

        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        result = bulk_add_comments(
            path, "Sheet1",
            comments=[{"address": "A1", "text": "hi", "author": "T"}],
            confirm=True,
        )
        assert "sheet" in result, "bulk_add_comments result must contain 'sheet' key"
        assert result["sheet"] == "Sheet1"

    @pytest.mark.unit
    def test_bulk_add_comments_single_comment(self, tmp_path, monkeypatch):
        """Single comment via list-of-1 succeeds (US-N09)."""
        from excelmcp._data import bulk_add_comments

        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        result = bulk_add_comments(
            path, "Sheet1",
            comments=[{"address": "B2", "text": "Sprint N test"}],
            confirm=True,
        )
        assert result.get("ok") is True
        assert result.get("comments_added") == 1
        assert result.get("skipped") == 0

    @pytest.mark.unit
    def test_bulk_add_comments_write_gate(self, tmp_path, monkeypatch):
        """EXCEL_ENABLE_WRITE unset → NotAllowedError (US-N17)."""
        from excelmcp._data import bulk_add_comments

        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        monkeypatch.delenv("EXCEL_ENABLE_WRITE", raising=False)
        with pytest.raises(NotAllowedError):
            bulk_add_comments(
                path, "Sheet1",
                comments=[{"address": "A1", "text": "x"}],
                confirm=True,
            )


# ---------------------------------------------------------------------------
# TestCommitBToolRegistry (US-N10)
# ---------------------------------------------------------------------------

class TestCommitBToolRegistry:
    """add_hyperlink and add_cell_comment must be absent from the registry."""

    @pytest.mark.smoke
    def test_add_hyperlink_not_registered(self):
        """add_hyperlink (singular) must NOT be in the MCP tool registry (US-N10)."""
        from excelmcp import server

        tool_names = _get_tool_names(server.mcp)
        assert "add_hyperlink" not in tool_names

    @pytest.mark.smoke
    def test_add_cell_comment_not_registered(self):
        """add_cell_comment (singular) must NOT be in the MCP tool registry (US-N10)."""
        from excelmcp import server

        tool_names = _get_tool_names(server.mcp)
        assert "add_cell_comment" not in tool_names

    @pytest.mark.smoke
    def test_bulk_add_hyperlinks_still_registered(self):
        """bulk_add_hyperlinks replaced by hyperlink unified tool (Sprint O Commit C)."""
        from excelmcp import server

        tool_names = _get_tool_names(server.mcp)
        caps = _get_capabilities_snapshot(server.mcp)
        assert "hyperlink" in tool_names
        assert "bulk_add_hyperlinks" not in tool_names
        assert "hyperlink" in caps["tools"]

    @pytest.mark.smoke
    def test_bulk_add_comments_still_registered(self):
        """bulk_add_comments must still be registered."""
        from excelmcp import server

        tool_names = _get_tool_names(server.mcp)
        caps = _get_capabilities_snapshot(server.mcp)
        assert "bulk_add_comments" in tool_names
        assert "bulk_add_comments" in caps["tools"]


# ===========================================================================
# COMMIT C — US-N11–N14 (add_data_validation unification)
# ===========================================================================


class TestCommitCToolRegistry:
    """add_data_validation (singular MCP tool) must be absent; bulk_ present."""

    @pytest.mark.smoke
    def test_add_data_validation_not_registered(self):
        """add_data_validation (singular MCP tool) must NOT be registered (US-N11)."""
        from excelmcp import server

        tool_names = _get_tool_names(server.mcp)
        assert "add_data_validation" not in tool_names

    @pytest.mark.smoke
    def test_bulk_add_data_validation_still_registered(self):
        """data_validation unified tool must be registered; bulk_add_data_validation must NOT (Sprint O Commit B)."""
        from excelmcp import server

        tool_names = _get_tool_names(server.mcp)
        caps = _get_capabilities_snapshot(server.mcp)
        assert "data_validation" in tool_names
        assert "bulk_add_data_validation" not in tool_names
        assert "data_validation" in caps["tools"]


# ---------------------------------------------------------------------------
# TestCommitCDataValidation (US-N12–N14)
# ---------------------------------------------------------------------------

class TestCommitCDataValidation:
    """Unified add_data_validation: single-dict, list, legacy, alias, gates."""

    @pytest.mark.unit
    def test_add_data_validation_single_rule_dict(self, tmp_path, monkeypatch):
        """New single-dict mode returns single-rule format (US-N12)."""
        from excelmcp._advanced import add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        result = add_data_validation(
            path, "Sheet1",
            rules={"address": "A1:A5", "validation_type": "list", "formula1": '"Yes,No"'},
            confirm=True,
        )
        assert result["ok"] is True
        assert result["address"] == "A1:A5"
        assert result["validation_type"] == "list"
        assert "message" in result
        assert "validations_added" not in result

    @pytest.mark.unit
    def test_add_data_validation_multi_rule_list(self, tmp_path, monkeypatch):
        """List mode returns validations_added format (US-N12)."""
        from excelmcp._advanced import add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        result = add_data_validation(
            path, "Sheet1",
            rules=[
                {"address": "A1:A5", "validation_type": "list", "formula1": '"Yes,No"'},
                {"address": "B1:B5", "validation_type": "whole",
                 "formula1": "1", "formula2": "10", "operator": "between"},
            ],
            confirm=True,
        )
        assert result["ok"] is True
        assert result["validations_added"] == 2
        assert "address" not in result
        assert "validation_type" not in result

    @pytest.mark.unit
    def test_add_data_validation_legacy_positional_address(self, tmp_path, monkeypatch):
        """Legacy: 3rd positional str address + 4th positional validation_type (US-N13)."""
        from excelmcp._advanced import add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        result = add_data_validation(
            path, "Sheet1", "C1:C10", "whole",
            formula1="1", formula2="100", operator="between",
            confirm=True,
        )
        assert result["ok"] is True
        assert result["address"] == "C1:C10"
        assert result["validation_type"] == "whole"

    @pytest.mark.unit
    def test_add_data_validation_legacy_kwargs_only(self, tmp_path, monkeypatch):
        """Legacy: address + validation_type as keyword-only args (US-N13)."""
        from excelmcp._advanced import add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        result = add_data_validation(
            path, "Sheet1",
            address="D1:D5",
            validation_type="list",
            formula1='"Alpha,Beta"',
            error_title="Err",
            prompt="Choose",
            confirm=True,
        )
        assert result["ok"] is True
        assert result["address"] == "D1:D5"

    @pytest.mark.unit
    def test_add_data_validation_write_gate_before_path(self, tmp_path, monkeypatch):
        """EXCEL_ENABLE_WRITE unset fires before path allowlist check (US-N14)."""
        from excelmcp._advanced import add_data_validation
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        monkeypatch.delenv("EXCEL_ENABLE_WRITE", raising=False)
        _wb_cache.clear()
        with pytest.raises(NotAllowedError):
            add_data_validation(
                str(tmp_path / "nonexistent.xlsx"), "Sheet1",
                rules={"address": "A1", "validation_type": "whole", "formula1": "1"},
                confirm=True,
            )

    @pytest.mark.unit
    def test_add_data_validation_confirm_gate(self, tmp_path, monkeypatch):
        """confirm=False raises ValidationError (US-N14)."""
        from excelmcp._advanced import add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        with pytest.raises(ValidationError, match="confirm=True"):
            add_data_validation(
                path, "Sheet1",
                rules={"address": "A1", "validation_type": "whole", "formula1": "1"},
                confirm=False,
            )

    @pytest.mark.unit
    def test_bulk_alias_returns_validations_added(self, tmp_path, monkeypatch):
        """bulk_add_data_validation is a thin alias; returns validations_added (US-N12)."""
        from excelmcp._advanced import bulk_add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        result = bulk_add_data_validation(
            path, "Sheet1",
            validations=[{"address": "E1:E3", "validation_type": "list",
                          "formula1": '"P,Q,R"'}],
            confirm=True,
        )
        assert result["ok"] is True
        assert result["validations_added"] == 1
        assert "address" not in result

    @pytest.mark.unit
    def test_add_data_validation_empty_rules_raises(self, tmp_path, monkeypatch):
        """Empty rules list raises ValidationError (US-N12)."""
        from excelmcp._advanced import add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        with pytest.raises(ValidationError):
            add_data_validation(path, "Sheet1", rules=[], confirm=True)

    @pytest.mark.unit
    def test_add_data_validation_cap_501_raises(self, tmp_path, monkeypatch):
        """More than 500 rules raises ValidationError mentioning cap (US-N12)."""
        from excelmcp._advanced import add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        rules = [{"address": "A1", "validation_type": "whole", "formula1": "1"}
                 for _ in range(501)]
        with pytest.raises(ValidationError, match="500"):
            add_data_validation(path, "Sheet1", rules=rules, confirm=True)

    @pytest.mark.unit
    def test_add_data_validation_invalid_type_in_list(self, tmp_path, monkeypatch):
        """Invalid validation_type in list raises ValidationError (US-N12)."""
        from excelmcp._advanced import add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        with pytest.raises(ValidationError, match="[Ii]nvalid validation_type"):
            add_data_validation(
                path, "Sheet1",
                rules=[{"address": "A1", "validation_type": "bogus"}],
                confirm=True,
            )

    @pytest.mark.unit
    def test_add_data_validation_missing_address_in_list(self, tmp_path, monkeypatch):
        """Empty address in list raises ValidationError (US-N12)."""
        from excelmcp._advanced import add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        with pytest.raises(ValidationError, match="address"):
            add_data_validation(
                path, "Sheet1",
                rules=[{"address": "", "validation_type": "whole", "formula1": "1"}],
                confirm=True,
            )

    @pytest.mark.unit
    def test_add_data_validation_list_requires_formula1_in_list_mode(
        self, tmp_path, monkeypatch
    ):
        """list type without formula1 raises ValidationError in list mode (US-N12)."""
        from excelmcp._advanced import add_data_validation
        path = _make_workbook(tmp_path, monkeypatch)
        _wb_cache.clear()
        with pytest.raises(ValidationError, match="formula1"):
            add_data_validation(
                path, "Sheet1",
                rules=[{"address": "A1", "validation_type": "list"}],
                confirm=True,
            )

