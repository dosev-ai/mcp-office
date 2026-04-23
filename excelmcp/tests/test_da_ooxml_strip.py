"""Unit tests for _strip_legacy_f_attrs and _LEGACY_CSE_F_TAG_RE.

Split from test_da_finalize.py at Sprint E. Pure-function tests — no fixtures, no file I/O.
"""
import pytest


class TestStripLegacyFAttrs:
    """Direct unit tests for ``_strip_legacy_f_attrs`` and
    ``_LEGACY_CSE_F_TAG_RE`` in ``excelmcp._ooxml``.

    These are pure-function tests — no fixtures, no file I/O.
    """

    # ── happy-path: attrs stripped ───────────────────────────────────────

    @pytest.mark.unit
    @pytest.mark.parametrize("inp, expected_tag", [
        (
            '<f ca="1" t="array">formula</f>',
            "<f>formula</f>",
        ),
        (
            '<f ref="A1:B2">formula</f>',
            "<f>formula</f>",
        ),
        (
            '<f ca="1">formula</f>',
            "<f>formula</f>",
        ),
        (
            '<f  t="array"  ref="C3">formula</f>',
            "<f>formula</f>",
        ),
    ])
    def test_strips_cse_attrs(self, inp, expected_tag):
        from excelmcp._ooxml import _strip_legacy_f_attrs
        result, changed = _strip_legacy_f_attrs(inp)
        assert result == expected_tag
        assert changed is True

    # ── no-op: non-CSE tags must be preserved exactly ────────────────────

    @pytest.mark.unit
    @pytest.mark.parametrize("inp", [
        "<f>SUM(A1:A10)</f>",
        '<f si="3">formula</f>',
        '<f t="shared">formula</f>',
        '<f aca="1">formula</f>',
        '<f t="array" aca="1" ref="A1" ca="1">formula</f>',
    ])
    def test_noop_on_safe_tags(self, inp):
        from excelmcp._ooxml import _strip_legacy_f_attrs
        result, changed = _strip_legacy_f_attrs(inp)
        assert result == inp
        assert changed is False

    # ── idempotency ───────────────────────────────────────────────────────

    @pytest.mark.unit
    def test_idempotent(self):
        from excelmcp._ooxml import _strip_legacy_f_attrs
        original = '<f t="array" ref="A1" ca="1">formula</f>'
        once, changed1 = _strip_legacy_f_attrs(original)
        assert changed1 is True
        twice, changed2 = _strip_legacy_f_attrs(once)
        assert once == twice
        assert changed2 is False

    # ── multi-tag: multiple <f> elements in one cell XML chunk ────────────

    @pytest.mark.unit
    def test_strips_multiple_f_tags_in_one_string(self):
        from excelmcp._ooxml import _strip_legacy_f_attrs
        inp = (
            '<c r="A1"><f t="array" ca="1">SUM(A1)</f></c>'
            '<c r="B1"><f aca="1">SUM(B1)</f></c>'
        )
        result, changed = _strip_legacy_f_attrs(inp)
        assert 't="array"' not in result
        assert ' ca="' not in result
        assert 'aca="1"' in result
        assert changed is True

    # ── _LEGACY_CSE_F_TAG_RE regex: direct match tests ───────────────────

    @pytest.mark.unit
    @pytest.mark.parametrize("tag", [
        '<f t="array" aca="1" ref="A1" ca="1">',
        '<f ca="1">',
        '<f aca="1">',
        '<f ref="A1:B5">',
        '<f t="array">',
    ])
    def test_regex_matches_cse_tags(self, tag):
        from excelmcp._ooxml import _LEGACY_CSE_F_TAG_RE
        assert _LEGACY_CSE_F_TAG_RE.search(tag) is not None

    @pytest.mark.unit
    @pytest.mark.parametrize("tag", [
        "<f>",
        '<f si="3">',
        '<f t="shared">',
    ])
    def test_regex_does_not_match_safe_tags(self, tag):
        from excelmcp._ooxml import _LEGACY_CSE_F_TAG_RE
        assert _LEGACY_CSE_F_TAG_RE.search(tag) is None

    # ── boolean flag: was_stripped reflects actual change ─────────────────

    @pytest.mark.unit
    def test_flag_true_when_attrs_stripped(self):
        from excelmcp._ooxml import _strip_legacy_f_attrs
        _, changed = _strip_legacy_f_attrs('<f t="array">SUM()</f>')
        assert changed is True

    @pytest.mark.unit
    def test_flag_false_when_no_attrs_present(self):
        from excelmcp._ooxml import _strip_legacy_f_attrs
        _, changed = _strip_legacy_f_attrs("<f>SUM()</f>")
        assert changed is False


class TestRewriteFormulaBody:
    """Unit tests for ``_rewrite_formula_body`` and ``_is_da_formula``.

    Covers the LET()-wrapped spill-function normalization bug: functions that
    arrive with the _xlfn._xlws. infix must be normalised to _xlfn.FUNC
    (without _xlws.) so Excel does not trigger the repair dialog.
    """

    # ── _is_da_formula: LET detection ─────────────────────────────────────

    @pytest.mark.unit
    def test_is_da_formula_detects_let_wrapping_filter(self):
        """_is_da_formula returns True for LET(src, FILTER(...), src)."""
        from excelmcp._ooxml import _is_da_formula
        assert _is_da_formula("LET(src, FILTER(B1:B10, B1:B10&gt;5), src)") is True

    @pytest.mark.unit
    def test_is_da_formula_detects_let_wrapping_xlookup(self):
        """_is_da_formula returns True for LET()-wrapped XLOOKUP."""
        from excelmcp._ooxml import _is_da_formula
        assert _is_da_formula("LET(r, XLOOKUP(A1, B:B, C:C), r)") is True

    @pytest.mark.unit
    def test_is_da_formula_detects_xlws_prefixed_filter(self):
        """_is_da_formula returns True when FILTER already has _xlfn._xlws. prefix."""
        from excelmcp._ooxml import _is_da_formula
        assert _is_da_formula("_xlfn._xlws.FILTER(A:A, A:A&gt;0)") is True

    @pytest.mark.unit
    def test_is_da_formula_detects_let_with_xlws_prefixed_filter(self):
        """_is_da_formula returns True for LET(src, _xlfn._xlws.FILTER(...), src)."""
        from excelmcp._ooxml import _is_da_formula
        assert _is_da_formula("LET(src, _xlfn._xlws.FILTER(A:A, A:A&gt;0), src)") is True

    # ── _rewrite_formula_body: _xlws. normalisation (BUG 2) ───────────────

    @pytest.mark.unit
    def test_rewrite_normalises_xlws_filter_to_xlfn_filter(self):
        """_xlfn._xlws.FILTER( is rewritten to _xlfn.FILTER(."""
        from excelmcp._ooxml import _rewrite_formula_body
        result = _rewrite_formula_body("_xlfn._xlws.FILTER(A:A, A:A&gt;0)")
        assert "_xlfn.FILTER(" in result
        assert "_xlws." not in result

    @pytest.mark.unit
    def test_rewrite_normalises_xlws_in_let_wrapper(self):
        """LET(src, _xlfn._xlws.FILTER(...), src) → LET(src, _xlfn.FILTER(...), src)."""
        from excelmcp._ooxml import _rewrite_formula_body
        formula = "LET(src, _xlfn._xlws.FILTER(B1:B10, B1:B10&gt;5), src)"
        result = _rewrite_formula_body(formula)
        assert "_xlfn.FILTER(" in result
        assert "_xlws." not in result
        assert "LET(" in result

    @pytest.mark.unit
    def test_rewrite_normalises_multiple_xlws_functions(self):
        """Multiple _xlfn._xlws.FUNC( occurrences are all normalised."""
        from excelmcp._ooxml import _rewrite_formula_body
        formula = "_xlfn._xlws.SORT(_xlfn._xlws.UNIQUE(A1:A10))"
        result = _rewrite_formula_body(formula)
        assert "_xlfn.SORT(" in result
        assert "_xlfn.UNIQUE(" in result
        assert "_xlws." not in result

    @pytest.mark.unit
    def test_rewrite_already_correct_xlfn_unchanged(self):
        """_xlfn.FILTER( (correct form) is left untouched by pass 1."""
        from excelmcp._ooxml import _rewrite_formula_body
        formula = "_xlfn.FILTER(A:A, A:A&gt;0)"
        result = _rewrite_formula_body(formula)
        assert result == formula  # no change needed

    @pytest.mark.unit
    def test_rewrite_adds_xlfn_to_bare_filter(self):
        """Unprefixed FILTER( → _xlfn.FILTER( (pass 2 still works)."""
        from excelmcp._ooxml import _rewrite_formula_body
        result = _rewrite_formula_body("FILTER(A:A, A:A&gt;0)")
        assert result == "_xlfn.FILTER(A:A, A:A&gt;0)"

    @pytest.mark.unit
    def test_rewrite_idempotent_on_xlws_normalised_formula(self):
        """Running _rewrite_formula_body twice gives the same result."""
        from excelmcp._ooxml import _rewrite_formula_body
        formula = "LET(src, _xlfn._xlws.FILTER(B1:B10, B1:B10&gt;5), src)"
        once = _rewrite_formula_body(formula)
        twice = _rewrite_formula_body(once)
        assert once == twice

    # ── patch_dynamic_array_metadata: end-to-end LET+FILTER ───────────────

    @pytest.mark.unit
    def test_patch_processes_let_wrapping_filter_formula(self, tmp_path):
        """patch_dynamic_array_metadata correctly handles LET(src, FILTER(...), src).

        Regression test for the LET-wrapped spill-formula bug: the patcher must
        add cm="1" to the cell, normalise any _xlfn._xlws. prefixes to _xlfn.,
        and leave LET unprefixed (LET is native in Excel 365).
        """
        import zipfile
        from excelmcp._ooxml import patch_dynamic_array_metadata

        # Simulate a file where Excel COM has written _xlfn._xlws.FILTER inside LET
        sheet_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData>'
            '<row r="1"><c r="A1">'
            '<f>LET(src, _xlfn._xlws.FILTER(B1:B10, B1:B10&gt;5), src)</f>'
            '</c></row>'
            '</sheetData>'
            '</worksheet>'
        )
        wb_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
        )
        wb_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1"'
            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
            ' Target="worksheets/sheet1.xml"/>'
            '</Relationships>'
        )
        ct_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels"'
            ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>'
        )

        xlsx_path = tmp_path / "let_filter_xlws.xlsx"
        with zipfile.ZipFile(str(xlsx_path), "w") as zf:
            zf.writestr("[Content_Types].xml", ct_xml)
            zf.writestr("xl/workbook.xml", wb_xml)
            zf.writestr("xl/_rels/workbook.xml.rels", wb_rels_xml)
            zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)

        patched_count = patch_dynamic_array_metadata(str(xlsx_path))

        with zipfile.ZipFile(str(xlsx_path), "r") as zf:
            patched_sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")

        # cm="1" must be added to the cell
        assert 'cm="1"' in patched_sheet, f"cm='1' missing. Sheet XML:\n{patched_sheet}"
        # _xlws. infix must be stripped
        assert "_xlws." not in patched_sheet, f"_xlws. still present. Sheet XML:\n{patched_sheet}"
        # _xlfn.FILTER( must be present (not bare FILTER or _xlws form)
        assert "_xlfn.FILTER(" in patched_sheet, f"_xlfn.FILTER( missing. Sheet XML:\n{patched_sheet}"
        # LET must remain unprefixed
        assert "LET(" in patched_sheet, f"LET( missing. Sheet XML:\n{patched_sheet}"
        assert patched_count >= 1
