import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from excelmcp.workbook_openpyxl import _wb_cache
from excelmcp._core import _load_wb


class TestFormula2SnapshotRestore:
    """Unit tests for snapshot_formula2_cells and restore_formula2_f_attrs.

    These tests fabricate raw OOXML to verify the snapshot captures Formula2
    cells (aca="1") and that restore re-injects those attributes after an
    openpyxl save strips them.  No Excel COM installation required.
    """

    def _make_xlsx_with_formula2(self, tmp_path: Path, formula2_xml: str) -> Path:
        """Build a minimal xlsx whose sheet1 contains *formula2_xml* verbatim."""
        import zipfile
        import io

        sheet_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<sheetData>"
            + formula2_xml
            + "</sheetData>"
            "</worksheet>"
        )
        ct_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            "</Types>"
        )
        wb_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<sheets>"
            '<sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
            "</sheets>"
            "</workbook>"
        )
        rels_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            "</Relationships>"
        )

        path = tmp_path / "formula2_test.xlsx"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            zout.writestr("[Content_Types].xml", ct_xml)
            zout.writestr("xl/workbook.xml", wb_xml)
            zout.writestr("xl/_rels/workbook.xml.rels", rels_xml)
            zout.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        path.write_bytes(buf.getvalue())
        return path

    def _read_sheet1_xml(self, path: Path) -> str:
        import zipfile
        with zipfile.ZipFile(str(path), "r") as z:
            return z.read("xl/worksheets/sheet1.xml").decode("utf-8")

    @pytest.mark.unit
    def test_snapshot_captures_formula2_cell(self, tmp_path):
        """snapshot_formula2_cells returns aca cell address and f_attrs."""
        from excelmcp._ooxml import snapshot_formula2_cells

        formula2_row = (
            '<row r="7">'
            '<c r="A7" s="1" t="str" cm="1">'
            '<f t="array" aca="1" ref="A7:C9" ca="1">LET(x,FILTER(A1:C5,A1:A5&gt;0,""),'
            'INDEX(x,SEQUENCE(ROWS(x)),{1,2,3}))</f>'
            "<v>hello</v>"
            "</c>"
            "</row>"
        )
        path = self._make_xlsx_with_formula2(tmp_path, formula2_row)

        snap = snapshot_formula2_cells(str(path))

        assert "xl/worksheets/sheet1.xml" in snap
        sheet_snap = snap["xl/worksheets/sheet1.xml"]
        assert "A7" in sheet_snap
        f_attrs = sheet_snap["A7"]
        assert 'aca="1"' in f_attrs
        assert 't="array"' in f_attrs
        assert 'ref="A7:C9"' in f_attrs

    @pytest.mark.unit
    def test_snapshot_ignores_non_formula2_cells(self, tmp_path):
        """snapshot_formula2_cells ignores plain formula cells (no aca=1)."""
        from excelmcp._ooxml import snapshot_formula2_cells

        plain_row = (
            '<row r="1">'
            '<c r="A1"><f>SUM(B1:B5)</f><v>10</v></c>'
            '<c r="B1" t="str"><f t="array" ref="B1:B3">SomeCSE()</f><v>x</v></c>'
            "</row>"
        )
        path = self._make_xlsx_with_formula2(tmp_path, plain_row)

        snap = snapshot_formula2_cells(str(path))

        assert snap == {}

    @pytest.mark.unit
    def test_restore_reinjects_formula2_attrs(self, tmp_path):
        """restore_formula2_f_attrs puts aca=1 attrs back after openpyxl strips them."""
        from excelmcp._ooxml import snapshot_formula2_cells, restore_formula2_f_attrs

        formula2_row = (
            '<row r="1">'
            '<c r="A1" cm="1">'
            '<f t="array" aca="1" ref="A1:C3" ca="1">_xlfn.FILTER(Sheet2!A:C,Sheet2!A:A&gt;0)</f>'
            "<v>foo</v>"
            "</c>"
            "</row>"
        )
        path = self._make_xlsx_with_formula2(tmp_path, formula2_row)

        snap = snapshot_formula2_cells(str(path))
        assert snap, "snapshot should find Formula2 cell"

        import zipfile
        import io as _io
        with open(str(path), "rb") as fh:
            orig = fh.read()
        buf = _io.BytesIO(orig)
        new_buf = _io.BytesIO()
        with zipfile.ZipFile(buf, "r") as zin, zipfile.ZipFile(new_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                if name == "xl/worksheets/sheet1.xml":
                    data = data.replace(b' t="array"', b'').replace(b' aca="1"', b'').replace(b' ca="1"', b'').replace(b' ref="A1:C3"', b'')
                zout.writestr(name, data)
        path.write_bytes(new_buf.getvalue())

        xml_after_strip = self._read_sheet1_xml(path)
        assert 'aca="1"' not in xml_after_strip

        restored = restore_formula2_f_attrs(str(path), snap)
        assert restored == 1

        xml_after_restore = self._read_sheet1_xml(path)
        assert 'aca="1"' in xml_after_restore
        assert 't="array"' in xml_after_restore
        assert 'ref="A1:C3"' in xml_after_restore

    @pytest.mark.unit
    def test_restore_noop_on_empty_snapshots(self, tmp_path):
        """restore_formula2_f_attrs with empty snapshots returns 0 without touching file."""
        from excelmcp._ooxml import restore_formula2_f_attrs

        dummy = tmp_path / "dummy.xlsx"
        dummy.write_bytes(b"placeholder")

        result = restore_formula2_f_attrs(str(dummy), {})
        assert result == 0

    @pytest.mark.unit
    def test_strip_legacy_cse_preserves_formula2(self):
        """_strip_legacy_f_attrs leaves Formula2 cells (aca=1) intact, strips legacy CSE."""
        from excelmcp._ooxml import _strip_legacy_f_attrs

        formula2_cell = '<c r="A7" cm="1"><f t="array" aca="1" ref="A7:K10" ca="1">LET(x,FILTER(a,b),x)</f>'
        result, changed = _strip_legacy_f_attrs(formula2_cell)
        assert 'aca="1"' in result, "Formula2 aca attr must be preserved"
        assert 't="array"' in result, "Formula2 t=array must be preserved"
        assert 'ref="A7:K10"' in result, "Formula2 ref must be preserved"

        cse_cell = '<c r="B2"><f t="array" ref="B2:B5">SUM(A1:A5*2)</f>'
        result_cse, changed_cse = _strip_legacy_f_attrs(cse_cell)
        assert 't="array"' not in result_cse, "Legacy CSE t=array must be stripped"
        assert 'ref=' not in result_cse, "Legacy CSE ref must be stripped"
        assert changed_cse is True


# ── GROUP: _load_wb mtime-based cache invalidation ────────────────────────


class TestWbCacheMtime:
    """Hermetic unit tests for mtime-based eviction in _load_wb.

    Uses mocked Path.stat and openpyxl.load_workbook to avoid touching any
    real filesystem paths.  Each test drives _load_wb itself (not bare-
    Workbook cache injection) so the real mtime-check code path is exercised.
    The autouse _clear_wb_cache fixture + setup_method both clear the cache to
    guarantee isolation.
    """

    FAKE_PATH = Path("C:/fake/workbook.xlsx")

    def setup_method(self):
        _wb_cache.clear()

    @pytest.mark.unit
    @patch("openpyxl.load_workbook")
    @patch("pathlib.Path.stat")
    def test_fresh_load_stores_mtime(self, mock_stat, mock_load_wb):
        """First load: load_workbook called once, (wb, mtime) tuple stored in cache."""
        mock_stat.return_value.st_mtime = 1000.0
        wb = MagicMock()
        mock_load_wb.return_value = wb

        result = _load_wb(self.FAKE_PATH, for_write=False)

        mock_load_wb.assert_called_once()
        assert _wb_cache[(str(self.FAKE_PATH), False)] == (wb, 1000.0)
        assert result is wb

    @pytest.mark.unit
    @patch("openpyxl.load_workbook")
    @patch("pathlib.Path.stat")
    def test_cache_hit_same_mtime_no_reload(self, mock_stat, mock_load_wb):
        """Repeated call with unchanged mtime: cache hit, load_workbook not called again."""
        mock_stat.return_value.st_mtime = 1000.0
        wb = MagicMock()
        mock_load_wb.return_value = wb

        result1 = _load_wb(self.FAKE_PATH, for_write=False)
        result2 = _load_wb(self.FAKE_PATH, for_write=False)

        mock_load_wb.assert_called_once()
        assert result1 is result2

    @pytest.mark.unit
    @patch("openpyxl.load_workbook")
    @patch("pathlib.Path.stat")
    def test_stale_cache_evicted_on_mtime_change(self, mock_stat, mock_load_wb):
        """Mtime change causes stale cache entry to be evicted and workbook reloaded."""
        wb_old = MagicMock(name="wb_old")
        wb_new = MagicMock(name="wb_new")
        mock_load_wb.side_effect = [wb_old, wb_new]
        mock_stat.return_value.st_mtime = 1000.0

        result1 = _load_wb(self.FAKE_PATH, for_write=False)

        mock_stat.return_value.st_mtime = 2000.0

        result2 = _load_wb(self.FAKE_PATH, for_write=False)

        assert mock_load_wb.call_count == 2
        assert result2 is wb_new
        assert result1 is not result2
        assert _wb_cache[(str(self.FAKE_PATH), False)] == (wb_new, 2000.0)

    @pytest.mark.unit
    @patch("openpyxl.load_workbook")
    @patch("pathlib.Path.stat")
    def test_oserror_on_stat_loads_workbook(self, mock_stat, mock_load_wb):
        """OSError from stat: workbook still loaded; cached mtime stored as None."""
        mock_stat.side_effect = OSError("no such file")
        wb = MagicMock()
        mock_load_wb.return_value = wb

        result = _load_wb(self.FAKE_PATH, for_write=False)

        mock_load_wb.assert_called_once()
        assert result is wb
        assert _wb_cache[(str(self.FAKE_PATH), False)] == (wb, None)

    @pytest.mark.unit
    @patch("openpyxl.load_workbook")
    @patch("pathlib.Path.stat")
    def test_write_and_read_modes_tracked_separately(self, mock_stat, mock_load_wb):
        """for_write=False and for_write=True use separate cache keys; data_only differs."""
        mock_stat.return_value.st_mtime = 1000.0
        wb_read = MagicMock(name="wb_read")
        wb_write = MagicMock(name="wb_write")
        mock_load_wb.side_effect = [wb_read, wb_write]

        result_read = _load_wb(self.FAKE_PATH, for_write=False)
        result_write = _load_wb(self.FAKE_PATH, for_write=True)

        assert result_read is wb_read
        assert result_write is wb_write
        assert mock_load_wb.call_count == 2
        assert _wb_cache[(str(self.FAKE_PATH), False)] == (wb_read, 1000.0)
        assert _wb_cache[(str(self.FAKE_PATH), True)] == (wb_write, 1000.0)
        calls = mock_load_wb.call_args_list
        assert calls[0].kwargs["data_only"] is True
        assert calls[1].kwargs["data_only"] is False
