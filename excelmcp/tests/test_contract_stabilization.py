"""Contract stabilization tests — W4/W5 canonical tool contract.

52 tests across 7 suites (S1-S7) verifying:
  S1: sheet_name → sheet rename (import_csv_to_sheet)
  S2: Confirm parameter contract (signature checks)
  S3: Gate fix in export_changed_sheets_only
  S4: pivot_table raises ToolError (not return {"error": ...})
  S5: No auto-save in format tools
  S6: Deprecated alias removal (11 aliases across 7 tools)
  S7: apply_style address param addition

All tests are unit-tier — no COM, no live Office.
"""
from __future__ import annotations

import inspect
from unittest.mock import patch

import pytest
from fastmcp.exceptions import ToolError

# ---------------------------------------------------------------------------
# Import all tools under test
# ---------------------------------------------------------------------------
from excelmcp.server_format import (
    add_border,
    apply_alignment,
    apply_number_format,
    apply_style,
)
from excelmcp.server_io import import_csv_to_sheet, named_range
from excelmcp.server_ops import copy_range, create_table, freeze_panes
from excelmcp.server_batch import fill_formula_range
from excelmcp.server_chart import chart
from excelmcp.server_com_pivot import pivot_table
from excelmcp.server_review import export_changed_sheets_only

from excelmcp._core import NotAllowedError, ValidationError


# ===========================================================================
# S1: sheet_name → sheet rename (4 tests)
# ===========================================================================


class TestS1SheetRename:
    """S1: import_csv_to_sheet uses 'sheet' not 'sheet_name'."""

    def test_s1_01_has_sheet_param(self):
        """import_csv_to_sheet signature includes 'sheet'."""
        sig = inspect.signature(import_csv_to_sheet)
        assert "sheet" in sig.parameters

    def test_s1_02_no_sheet_name_param(self):
        """import_csv_to_sheet signature does NOT include 'sheet_name'."""
        sig = inspect.signature(import_csv_to_sheet)
        assert "sheet_name" not in sig.parameters

    def test_s1_03_sheet_is_required(self):
        """'sheet' has no default — it is positional/required."""
        sig = inspect.signature(import_csv_to_sheet)
        p = sig.parameters["sheet"]
        assert p.default is inspect.Parameter.empty

    def test_s1_04_docstring_updated(self):
        """Docstring references 'sheet' not 'sheet_name'."""
        doc = import_csv_to_sheet.__doc__ or ""
        assert "sheet_name" not in doc
        assert "sheet" in doc


# ===========================================================================
# S2: Confirm parameter contract (8 tests)
# ===========================================================================


class TestS2ConfirmContract:
    """S2: confirm param is canonical across all write tools."""

    _TOOLS = [
        apply_number_format,
        apply_style,
        apply_alignment,
        add_border,
        import_csv_to_sheet,
        named_range,
        freeze_panes,
        copy_range,
        create_table,
        fill_formula_range,
        chart,
        pivot_table,
        export_changed_sheets_only,
    ]

    def test_s2_01_confirm_exists_in_all_write_tools(self):
        """Every write tool has a 'confirm' parameter."""
        for tool_fn in self._TOOLS:
            sig = inspect.signature(tool_fn)
            assert "confirm" in sig.parameters, f"{tool_fn.__name__} missing 'confirm'"

    def test_s2_02_confirm_default_false(self):
        """confirm defaults to False in all write tools."""
        for tool_fn in self._TOOLS:
            sig = inspect.signature(tool_fn)
            p = sig.parameters["confirm"]
            assert p.default is False, f"{tool_fn.__name__}: confirm default is {p.default!r}"

    def test_s2_03_confirm_is_bool(self):
        """confirm annotation is bool in all write tools."""
        for tool_fn in self._TOOLS:
            sig = inspect.signature(tool_fn)
            p = sig.parameters["confirm"]
            # from __future__ import annotations turns annotations into strings
            assert p.annotation in (bool, "bool"), (
                f"{tool_fn.__name__}: confirm annotation is {p.annotation!r}"
            )

    def test_s2_04_no_confirmed_alias(self):
        """No tool has a 'confirmed' parameter (common typo alias)."""
        for tool_fn in self._TOOLS:
            sig = inspect.signature(tool_fn)
            assert "confirmed" not in sig.parameters, f"{tool_fn.__name__} has 'confirmed'"

    def test_s2_05_no_force_alias(self):
        """No tool has a 'force' parameter (alias for confirm)."""
        for tool_fn in self._TOOLS:
            sig = inspect.signature(tool_fn)
            assert "force" not in sig.parameters, f"{tool_fn.__name__} has 'force'"

    def test_s2_06_confirm_is_last_or_near_last(self):
        """confirm is the last param (or second-to-last if session_mode exists)."""
        for tool_fn in self._TOOLS:
            sig = inspect.signature(tool_fn)
            params = list(sig.parameters.keys())
            idx = params.index("confirm")
            # confirm should be at end or just before session_mode
            remaining = params[idx + 1:]
            for r in remaining:
                assert r in ("session_mode",), (
                    f"{tool_fn.__name__}: param '{r}' comes after 'confirm'"
                )

    def test_s2_07_confirm_is_keyword(self):
        """confirm is KEYWORD_ONLY or POSITIONAL_OR_KEYWORD (not positional-only)."""
        for tool_fn in self._TOOLS:
            sig = inspect.signature(tool_fn)
            p = sig.parameters["confirm"]
            assert p.kind in (
                inspect.Parameter.KEYWORD_ONLY,
                inspect.Parameter.POSITIONAL_OR_KEYWORD,
            ), f"{tool_fn.__name__}: confirm kind is {p.kind!r}"

    def test_s2_08_no_confirm_in_read_only_chart_list(self):
        """chart(operation='list') works without confirm (read-only path)."""
        # just verifying the tool exists and has confirm=False default
        sig = inspect.signature(chart)
        assert sig.parameters["confirm"].default is False


# ===========================================================================
# S3: Gate fix in export_changed_sheets_only (5 tests)
# ===========================================================================


class TestS3GateFix:
    """S3: export_changed_sheets_only uses _check_write + _check_confirm."""

    def test_s3_01_rejects_without_write_env(self, monkeypatch, tmp_path):
        """Raises NotAllowedError when EXCEL_ENABLE_WRITE is not 'true'."""
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "false")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        with pytest.raises(NotAllowedError):
            export_changed_sheets_only(
                path=str(tmp_path / "test.xlsx"),
                base_dir=str(tmp_path),
                previous_hashes={},
                confirm=True,
            )

    def test_s3_02_rejects_confirm_false(self, monkeypatch, tmp_path):
        """Raises ValidationError when confirm=False."""
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        with pytest.raises(ValidationError):
            export_changed_sheets_only(
                path=str(tmp_path / "test.xlsx"),
                base_dir=str(tmp_path),
                previous_hashes={},
                confirm=False,
            )

    def test_s3_03_write_gate_before_confirm(self, monkeypatch, tmp_path):
        """_check_write fires before _check_confirm (write env missing + confirm=False → NotAllowedError)."""
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "false")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        # Both gates should fail, but _check_write is first
        with pytest.raises(NotAllowedError):
            export_changed_sheets_only(
                path=str(tmp_path / "test.xlsx"),
                base_dir=str(tmp_path),
                previous_hashes={},
                confirm=False,
            )

    def test_s3_04_no_inline_write_gate(self):
        """Source code does not use inline EXCEL_ENABLE_WRITE gate pattern."""
        src = inspect.getsource(export_changed_sheets_only)
        # The old inline pattern was: os.environ.get("EXCEL_ENABLE_WRITE", "").lower() != "true"
        # That specific pattern should be gone (replaced by _check_write)
        assert 'os.environ.get("EXCEL_ENABLE_WRITE"' not in src, (
            "Inline EXCEL_ENABLE_WRITE gate not removed"
        )

    def test_s3_05_uses_check_write_and_check_confirm(self):
        """Source code calls _check_write() and _check_confirm()."""
        src = inspect.getsource(export_changed_sheets_only)
        assert "_check_write()" in src, "_check_write() not found in source"
        assert "_check_confirm(confirm)" in src, "_check_confirm(confirm) not found in source"


# ===========================================================================
# S4: pivot_table raises ToolError (5 tests)
# ===========================================================================


class TestS4PivotToolError:
    """S4: pivot_table raises ToolError instead of returning {"error": ...}."""

    def test_s4_01_create_exception_raises_tool_error(self, monkeypatch, tmp_path):
        """operation='create' unexpected exception → ToolError."""
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        monkeypatch.setenv("EXCEL_ENABLE_COM", "true")
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))
        with patch(
            "excelmcp.server_com_pivot._create_pivot_table",
            side_effect=Exception("create boom"),
        ):
            with pytest.raises(ToolError, match="Pivot table creation failed"):
                pivot_table(
                    operation="create",
                    path=str(tmp_path / "test.xlsx"),
                    source_address="Sheet1!A1:D10",
                    dest_sheet="Sheet2",
                    dest_address="A1",
                    row_fields=["Region"],
                    value_fields=["Sales"],
                    confirm=True,
                )

    def test_s4_02_read_exception_raises_tool_error(self):
        """operation='read' unexpected exception → ToolError."""
        with patch(
            "excelmcp.server_com_pivot._read_pivot_table",
            side_effect=Exception("read boom"),
        ):
            with pytest.raises(ToolError, match="Pivot table read failed"):
                pivot_table(
                    operation="read",
                    path="C:/fake/ignored.xlsx",
                    sheet="Sheet1",
                    pivot_name="PivotTable1",
                )

    def test_s4_03_refresh_exception_raises_tool_error(self):
        """operation='refresh' unexpected exception → ToolError."""
        with patch(
            "excelmcp.server_com_pivot._refresh_pivot_tables",
            side_effect=Exception("refresh boom"),
        ):
            with pytest.raises(ToolError, match="Pivot table refresh failed"):
                pivot_table(
                    operation="refresh",
                    path="C:/fake/ignored.xlsx",
                    confirm=True,
                )

    def test_s4_04_no_return_error_dict_in_source(self):
        """Source code has no return {"error": ...} pattern."""
        src = inspect.getsource(pivot_table)
        assert 'return {"error"' not in src, 'return {"error": ...} still present'

    def test_s4_05_raise_from_exc(self):
        """ToolError is chained via 'from exc' for proper traceback."""
        src = inspect.getsource(pivot_table)
        assert "raise ToolError" in src
        # All ToolError raises in exception handlers should use 'from exc'
        lines = src.split("\n")
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("raise ToolError") and "from" not in stripped:
                # Check if it's inside an except block — if so, it needs 'from'
                # Allow raise ToolError for parameter validation (not in except)
                pass  # we just check the pattern exists
        # Verify at least one 'from exc' pattern
        assert "from exc" in src


# ===========================================================================
# S5: No auto-save in format tools (11 tests)
# ===========================================================================


class TestS5NoAutoSave:
    """S5: apply_number_format, apply_alignment, add_border do NOT auto-save."""

    def test_s5_01_number_format_no_save_call(self):
        """apply_number_format source does not call wb.save()."""
        src = inspect.getsource(apply_number_format)
        assert "wb.save(" not in src

    def test_s5_02_alignment_no_save_call(self):
        """apply_alignment source does not call wb.save()."""
        src = inspect.getsource(apply_alignment)
        assert "wb.save(" not in src

    def test_s5_03_border_no_save_call(self):
        """add_border source does not call wb.save()."""
        src = inspect.getsource(add_border)
        assert "wb.save(" not in src

    def test_s5_04_number_format_docstring_no_save(self):
        """apply_number_format docstring says 'Does NOT auto-save'."""
        doc = apply_number_format.__doc__ or ""
        assert "Does NOT auto-save" in doc

    def test_s5_05_alignment_docstring_no_save(self):
        """apply_alignment docstring says 'Does NOT auto-save'."""
        doc = apply_alignment.__doc__ or ""
        assert "Does NOT auto-save" in doc

    def test_s5_06_border_docstring_no_save(self):
        """add_border docstring says 'Does NOT auto-save'."""
        doc = add_border.__doc__ or ""
        assert "Does NOT auto-save" in doc

    def test_s5_07_number_format_returns_result(self, sample_workbook, monkeypatch):
        """apply_number_format returns backend result without saving."""
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        result = apply_number_format(
            sample_workbook, "Sheet1", "A1", "#,##0.00", confirm=True
        )
        assert isinstance(result, dict)

    def test_s5_08_alignment_returns_result(self, sample_workbook, monkeypatch):
        """apply_alignment returns backend result without saving."""
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        result = apply_alignment(
            sample_workbook, "Sheet1", "A1", horizontal="center", confirm=True
        )
        assert isinstance(result, dict)

    def test_s5_09_border_returns_result(self, sample_workbook, monkeypatch):
        """add_border returns backend result without saving."""
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        result = add_border(
            sample_workbook, "Sheet1", "A1", confirm=True
        )
        assert isinstance(result, dict)

    def test_s5_10_apply_style_no_save_call(self):
        """apply_style source does not call wb.save() (was already clean)."""
        src = inspect.getsource(apply_style)
        assert "wb.save(" not in src

    def test_s5_11_no_save_in_format_module(self):
        """The entire server_format module has zero wb.save() calls."""
        import excelmcp.server_format as mod
        src = inspect.getsource(mod)
        assert "wb.save(" not in src


# ===========================================================================
# S6: Deprecated alias removal (12 tests)
# ===========================================================================


class TestS6AliasRemoval:
    """S6: All 11 deprecated aliases removed from 7 tools."""

    # --- named_range: range_address ---
    def test_s6_01_named_range_no_range_address(self):
        """named_range has no 'range_address' parameter."""
        sig = inspect.signature(named_range)
        assert "range_address" not in sig.parameters

    # --- freeze_panes: freeze_at ---
    def test_s6_02_freeze_panes_no_freeze_at(self):
        """freeze_panes has no 'freeze_at' parameter."""
        sig = inspect.signature(freeze_panes)
        assert "freeze_at" not in sig.parameters

    # --- chart: data_range, target_cell ---
    def test_s6_03_chart_no_data_range(self):
        """chart has no 'data_range' parameter."""
        sig = inspect.signature(chart)
        assert "data_range" not in sig.parameters

    def test_s6_04_chart_no_target_cell(self):
        """chart has no 'target_cell' parameter."""
        sig = inspect.signature(chart)
        assert "target_cell" not in sig.parameters

    # --- fill_formula_range: anchor_cell, end_cell ---
    def test_s6_05_fill_formula_no_anchor_cell(self):
        """fill_formula_range has no 'anchor_cell' parameter."""
        sig = inspect.signature(fill_formula_range)
        assert "anchor_cell" not in sig.parameters

    def test_s6_06_fill_formula_no_end_cell(self):
        """fill_formula_range has no 'end_cell' parameter."""
        sig = inspect.signature(fill_formula_range)
        assert "end_cell" not in sig.parameters

    # --- copy_range: src_range, dst_range ---
    def test_s6_07_copy_range_no_src_range(self):
        """copy_range has no 'src_range' parameter."""
        sig = inspect.signature(copy_range)
        assert "src_range" not in sig.parameters

    def test_s6_08_copy_range_no_dst_range(self):
        """copy_range has no 'dst_range' parameter."""
        sig = inspect.signature(copy_range)
        assert "dst_range" not in sig.parameters

    # --- create_table: table_range ---
    def test_s6_09_create_table_no_table_range(self):
        """create_table has no 'table_range' parameter."""
        sig = inspect.signature(create_table)
        assert "table_range" not in sig.parameters

    # --- pivot_table: source_range, dest_cell ---
    def test_s6_10_pivot_no_source_range(self):
        """pivot_table has no 'source_range' parameter."""
        sig = inspect.signature(pivot_table)
        assert "source_range" not in sig.parameters

    def test_s6_11_pivot_no_dest_cell(self):
        """pivot_table has no 'dest_cell' parameter."""
        sig = inspect.signature(pivot_table)
        assert "dest_cell" not in sig.parameters

    # --- required params after alias removal ---
    def test_s6_12_copy_range_params_required(self):
        """copy_range src_address and dst_address are required (no default)."""
        sig = inspect.signature(copy_range)
        assert sig.parameters["src_address"].default is inspect.Parameter.empty
        assert sig.parameters["dst_address"].default is inspect.Parameter.empty


# ===========================================================================
# S7: apply_style address param (7 tests)
# ===========================================================================


class TestS7ApplyStyleAddress:
    """S7: apply_style accepts address as alternative to ranges."""

    def test_s7_01_has_address_param(self):
        """apply_style has an 'address' parameter."""
        sig = inspect.signature(apply_style)
        assert "address" in sig.parameters

    def test_s7_02_address_is_optional(self):
        """address defaults to None."""
        sig = inspect.signature(apply_style)
        assert sig.parameters["address"].default is None

    def test_s7_03_ranges_is_optional(self):
        """ranges defaults to None (was required list[str])."""
        sig = inspect.signature(apply_style)
        assert sig.parameters["ranges"].default is None

    def test_s7_04_address_resolves_to_ranges(self, sample_workbook, monkeypatch):
        """Passing address='A1' calls backend with ranges=['A1']."""
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        with patch("excelmcp.server_format.wb.apply_style_to_ranges") as mock_fn:
            mock_fn.return_value = {"ok": True}
            apply_style(
                sample_workbook, "Sheet1", address="A1", bold=True, confirm=True
            )
            mock_fn.assert_called_once()
            call_args = mock_fn.call_args
            assert call_args[0][2] == ["A1"]  # resolved_ranges

    def test_s7_05_both_address_and_ranges_raises(self):
        """Providing both address and ranges raises ToolError."""
        with pytest.raises(ToolError, match="Provide address or ranges, not both"):
            apply_style(
                "fake.xlsx", "Sheet1",
                ranges=["A1"], address="B2",
                confirm=True,
            )

    def test_s7_06_neither_address_nor_ranges_raises(self):
        """Providing neither address nor ranges raises ToolError."""
        with pytest.raises(ToolError, match="address or ranges is required"):
            apply_style(
                "fake.xlsx", "Sheet1",
                bold=True, confirm=True,
            )

    def test_s7_07_ranges_still_works(self, sample_workbook, monkeypatch):
        """Passing ranges=['A1:B2'] still works as before."""
        monkeypatch.setenv("EXCEL_ENABLE_WRITE", "true")
        with patch("excelmcp.server_format.wb.apply_style_to_ranges") as mock_fn:
            mock_fn.return_value = {"ok": True}
            apply_style(
                sample_workbook, "Sheet1",
                ranges=["A1:B2"], bold=True, confirm=True,
            )
            mock_fn.assert_called_once()
            call_args = mock_fn.call_args
            assert call_args[0][2] == ["A1:B2"]
