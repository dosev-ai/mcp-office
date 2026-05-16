"""Unit tests for read_summary_routine.read_workbook_range_v1.

CI-safe: no COM, no win32com, no live Office required.
All I/O uses temporary directories and in-memory openpyxl writes.
"""
from __future__ import annotations


import pytest
from openpyxl import Workbook

from excelmcp.read_summary_routine import read_workbook_range_v1


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_workbook(path: str, sheet_name: str, rows: list[list]) -> None:
    """Write a simple workbook with given rows to path."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


class TestAllowlistNotConfigured:
    """GAP-S2-01 T1: EXCEL_ALLOWLIST_ROOTS absent → config_error."""

    def test_allowlist_not_configured(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.delenv("EXCEL_ALLOWLIST_ROOTS", raising=False)
        result = read_workbook_range_v1(
            workbook_path="/some/path/workbook.xlsx",
            sheet_name="Sheet1",
        )
        assert result["error_kind"] == "config_error"
        assert "EXCEL_ALLOWLIST_ROOTS" in result["error"]

    def test_allowlist_empty_string(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", "")
        result = read_workbook_range_v1(
            workbook_path="/some/path/workbook.xlsx",
            sheet_name="Sheet1",
        )
        assert result["error_kind"] == "config_error"


class TestAllowlistViolation:
    """GAP-S2-01 T2: path outside configured roots → allowlist_violation."""

    def test_allowlist_violation(self, monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
        # Configure an allowed root that is a different temp dir
        other_dir = tmp_path / "other"
        other_dir.mkdir()
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(other_dir))

        # Request a path outside the allowed root
        result = read_workbook_range_v1(
            workbook_path=str(tmp_path / "not_in_root.xlsx"),
            sheet_name="Sheet1",
        )
        assert result["error_kind"] == "allowlist_violation"
        assert "not in allowlist" in result["error"]

    def test_path_traversal_blocked(self, monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
        allowed = tmp_path / "allowed"
        allowed.mkdir()
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(allowed))

        # Attempt a traversal path
        result = read_workbook_range_v1(
            workbook_path=str(allowed / ".." / "secret.xlsx"),
            sheet_name="Sheet1",
        )
        # realpath resolves ".." so it ends up outside allowed dir
        assert result["error_kind"] in ("allowlist_violation", "unexpected_error")


class TestReadSuccess:
    """GAP-S2-01 T3: happy-path read returns correct rows/headers."""

    def test_read_success(self, monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))

        wb_path = str(tmp_path / "test.xlsx")
        _make_workbook(
            wb_path,
            "Data",
            [
                ["Name", "Value", "Flag"],
                ["Alpha", 1, True],
                ["Beta", 2, False],
                ["Gamma", 3, True],
            ],
        )

        result = read_workbook_range_v1(
            workbook_path=wb_path,
            sheet_name="Data",
        )

        assert "error" not in result, f"Unexpected error: {result}"
        assert result["headers"] == ["Name", "Value", "Flag"]
        assert result["row_count"] == 3
        assert result["sheet_name"] == "Data"
        assert result["workbook_path"] == wb_path
        assert len(result["rows"]) == 3
        assert result["rows"][0][0] == "Alpha"
        assert result["evidence"]["routine"] == "excel.read_workbook_range_v1"

    def test_read_with_range_ref(self, monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))

        wb_path = str(tmp_path / "ranged.xlsx")
        _make_workbook(
            wb_path,
            "Sheet1",
            [
                ["A", "B", "C", "D"],
                [1, 2, 3, 4],
                [5, 6, 7, 8],
            ],
        )

        result = read_workbook_range_v1(
            workbook_path=wb_path,
            sheet_name="Sheet1",
            range_ref="A1:C3",
        )

        assert "error" not in result, f"Unexpected error: {result}"
        assert result["headers"] == ["A", "B", "C"]
        assert result["row_count"] == 2
        assert result["evidence"]["range_ref"] == "A1:C3"

    def test_empty_sheet(self, monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))

        wb_path = str(tmp_path / "empty.xlsx")
        _make_workbook(wb_path, "Sheet1", [])

        result = read_workbook_range_v1(workbook_path=wb_path, sheet_name="Sheet1")
        assert "error" not in result
        assert result["row_count"] == 0
        assert result["headers"] == []
        assert result["rows"] == []

    def test_missing_sheet(self, monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))

        wb_path = str(tmp_path / "wb.xlsx")
        _make_workbook(wb_path, "Sheet1", [["H1"], [1]])

        result = read_workbook_range_v1(
            workbook_path=wb_path, sheet_name="DoesNotExist"
        )
        assert result["error_kind"] == "sheet_not_found"


class TestMaxRowsLimit:
    """GAP-S2-01 T4: max_rows limits returned data rows."""

    def test_max_rows_limit(self, monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))

        wb_path = str(tmp_path / "big.xlsx")
        # 1 header + 10 data rows
        rows = [["Col1", "Col2"]] + [[f"R{i}", i] for i in range(10)]
        _make_workbook(wb_path, "Sheet1", rows)

        result = read_workbook_range_v1(
            workbook_path=wb_path,
            sheet_name="Sheet1",
            max_rows=3,
        )

        assert "error" not in result, f"Unexpected error: {result}"
        assert result["row_count"] == 3
        assert len(result["rows"]) == 3
        assert result["evidence"]["max_rows"] == 3

    def test_max_rows_larger_than_data(self, monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
        monkeypatch.setenv("EXCEL_ALLOWLIST_ROOTS", str(tmp_path))

        wb_path = str(tmp_path / "small.xlsx")
        rows = [["H"]] + [[i] for i in range(5)]
        _make_workbook(wb_path, "Sheet1", rows)

        result = read_workbook_range_v1(
            workbook_path=wb_path,
            sheet_name="Sheet1",
            max_rows=1000,
        )

        assert "error" not in result
        assert result["row_count"] == 5
