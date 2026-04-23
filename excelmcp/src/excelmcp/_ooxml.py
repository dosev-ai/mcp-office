"""
Post-save OOXML patcher for Excel 365 dynamic array / spill formula support.

openpyxl does not emit:
  - the ``cm`` attribute on cells containing dynamic array formulas (LET, FILTER,
    SORT, UNIQUE, SEQUENCE, XLOOKUP, XMATCH, RANDARRAY, …)
  - ``xl/metadata.xml`` with the XLDAPR cell-metadata record
  - the corresponding workbook relationship and Content-Type Override

Without these, Excel's repair dialog removes every spill formula on first open.

``patch_dynamic_array_metadata(path)`` runs as a post-save step and adds all
required structures so the file opens cleanly in Excel 365 with no repair prompt.
"""
from __future__ import annotations

import io
import os
import re
import zipfile

# ── Constants ───────────────────────────────────────────────────────────────

# Functions that make a cell a dynamic-array / spill anchor in Excel 365.
_DA_FUNCS = re.compile(
    r"\b(LET|FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|XLOOKUP|XMATCH"
    r"|RANDARRAY|EXPAND|TAKE|DROP|TOROW|TOCOL|WRAPROWS|WRAPCOLS|MAKEARRAY)\s*\(",
    re.IGNORECASE,
)

# Excel 365 OOXML requires dynamic-array / spill functions to be stored with
# the _xlfn. prefix so the repair engine can validate them.
# Without these prefixes Excel removes ("Removed Records: Formula") the formula.
# NOTE: _xlfn._xlws.* forms (with the _xlws. infix) break when those functions
# are nested inside LET().  The universal-safe form is _xlfn.FUNC (no _xlws.).
# LET is native in Excel 365 and does NOT need any prefix — omit it from this map.
_DA_FUNC_MAP: dict[str, str] = {
    # Spill functions — need _xlfn. prefix (no _xlws. — this is the OOXML-safe universal form)
    "FILTER":    "_xlfn.FILTER",
    "SORT":      "_xlfn.SORT",
    "SORTBY":    "_xlfn.SORTBY",
    "UNIQUE":    "_xlfn.UNIQUE",
    "SEQUENCE":  "_xlfn.SEQUENCE",
    "RANDARRAY": "_xlfn.RANDARRAY",
    "EXPAND":    "_xlfn.EXPAND",
    "TAKE":      "_xlfn.TAKE",
    "DROP":      "_xlfn.DROP",
    "TOROW":     "_xlfn.TOROW",
    "TOCOL":     "_xlfn.TOCOL",
    "WRAPROWS":  "_xlfn.WRAPROWS",
    "WRAPCOLS":  "_xlfn.WRAPCOLS",
    # LET deliberately OMITTED — LET is native in Excel 365 (no _xlfn. prefix needed)
    # Lookup extensions — need _xlfn. prefix
    "XLOOKUP":   "_xlfn.XLOOKUP",
    "XMATCH":    "_xlfn.XMATCH",
    "MAKEARRAY": "_xlfn.MAKEARRAY",
}

# Matches a DA function name NOT already preceded by '.' or '_'
# (which would indicate it already carries _xlfn./_xlws. prefix).
_DA_FUNC_PREFIX_RE = re.compile(
    r"(?<![._])\b(" + "|".join(re.escape(k) for k in _DA_FUNC_MAP) + r")\s*\(",
    re.IGNORECASE,
)

# Matches the _xlfn._xlws.FUNC( form that Excel COM / Excel desktop writes when
# a spill function is authored directly in Excel.  The _xlws. infix is problematic
# when the function is nested inside LET(): Excel's repair engine does not recognise
# the _xlws. form as valid inside LET and triggers the "Removed Records: Formula"
# repair dialog.  The universal-safe OOXML form is _xlfn.FUNC (no _xlws. infix).
_DA_XLWS_PREFIX_RE = re.compile(
    r"_xlfn\._xlws\.(" + "|".join(re.escape(k) for k in _DA_FUNC_MAP) + r")\s*\(",
    re.IGNORECASE,
)

# Detects a <f> opening tag that still carries legacy CSE array attributes
# (t="array", aca, ref, ca).  Files saved by an older patcher version may
# have these; they confuse Excel's repair engine and must be stripped.
#
# Note: COM Formula2 cells also carry these attributes (they have aca="1").
# The patcher strips them here, but _core._save_and_evict snapshotted them
# before the openpyxl save and restore_formula2_f_attrs re-injects them after
# patching, so LET+FILTER Formula2 cells end up with the correct attributes.
_LEGACY_CSE_F_TAG_RE = re.compile(
    r"<f\b([^>]*\b(?:t=\"array\"|aca=|ref=|ca=)[^>]*)>",
    re.IGNORECASE,
)

# Matches a COM Formula2 cell: <c r="XX" [attrs]> [ws] <f [attrs with aca="1"] >
# Group 1: cell address (r="XX" value)
# Group 2: full attribute string inside the <f> opening tag (contains aca="1")
_FORMULA2_CELL_RE = re.compile(
    r'<c\b[^>]*\br="([^"]+)"[^>]*>[^<]*'
    r'<f\b([^>]*\baca="1"[^>]*)>',
    re.IGNORECASE,
)


def _strip_legacy_f_attrs(cell_xml: str) -> tuple[str, bool]:
    """Strip legacy CSE array attributes (t="array", aca, ref, ca) from every
    ``<f …>`` opening tag found in *cell_xml*.  Returns ``(result, changed)``.

    **Exception**: cells with ``aca="1"`` (COM Formula2 cells written via
    `.Formula2` in Excel desktop) are left untouched.  The snapshot/restore
    mechanism in ``_core._save_and_evict`` / ``restore_formula2_f_attrs``
    preserves these cells; stripping them here would destroy their identity.

    Legacy CSE cells (``t="array"`` without ``aca="1"``) are still stripped
    so every DA cell uses the ``cm="1"`` + ``_xlfn.`` prefix approach for
    Excel 365.
    """

    def _cleaner(m: re.Match) -> str:
        attrs = m.group(1)
        # COM Formula2 cells carry aca="1"; the snapshot/restore mechanism in
        # _core._save_and_evict handles them — do NOT strip their attributes here.
        if re.search(r'\baca="1"', attrs):
            return m.group(0)
        attrs = re.sub(r'\s+t="array"', "", attrs, flags=re.IGNORECASE)
        attrs = re.sub(r'\s+aca="[^"]*"', "", attrs, flags=re.IGNORECASE)
        attrs = re.sub(r'\s+ref="[^"]*"', "", attrs, flags=re.IGNORECASE)
        attrs = re.sub(r'\s+ca="[^"]*"', "", attrs, flags=re.IGNORECASE)
        return f"<f{attrs}>"

    healed, _ = _LEGACY_CSE_F_TAG_RE.subn(_cleaner, cell_xml)
    return healed, healed != cell_xml

# Relationship type for sheetMetadata (metadata.xml lives at workbook level).
_REL_TYPE_METADATA = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
)
_CONTENT_TYPE_METADATA = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
)

# The metadata.xml content is always identical regardless of how many spill
# cells exist — all of them share the one XLDAPR entry (cm="1").
_METADATA_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ' xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    "<metadataTypes count=\"1\">"
    '<metadataType name="XLDAPR" minSupportedVersion="120000"'
    ' copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1"'
    ' rowColShift="1" clearFormats="1" clearComments="1" assign="1"'
    ' coerce="1" cellMeta="1"/>'
    "</metadataTypes>"
    '<futureMetadata name="XLDAPR" count="1">'
    "<bk><extLst>"
    '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
    "</ext></extLst></bk>"
    "</futureMetadata>"
    '<cellMetadata count="1">'
    "<bk><rc t=\"1\" v=\"0\"/></bk>"
    "</cellMetadata>"
    "</metadata>"
)


# ── Cell-XML patcher ─────────────────────────────────────────────────────────

def _is_da_formula(formula_text: str) -> bool:
    """Return True if *formula_text* contains a dynamic-array function."""
    return bool(_DA_FUNCS.search(formula_text))


def _rewrite_formula_body(formula_xml_text: str) -> str:
    """Normalise DA function prefixes in the formula body between ``<f>`` and ``</f>``.

    The *formula_xml_text* is the raw content between ``<f>`` and ``</f>``
    in the worksheet XML (already XML-encoded, e.g. ``>`` → ``&gt;``).
    Function names never contain ``<``/``>`` so XML entities don't interfere
    with the word-boundary matching.

    Two-pass normalisation:

    **Pass 1** — strip the ``_xlws.`` infix from ``_xlfn._xlws.FUNC(`` forms.
    The ``_xlfn._xlws.`` form is written by Excel COM / Excel desktop and works
    for standalone spill formulas, but breaks when nested inside ``LET()``
    (Excel repair dialog: "Removed Records: Formula").  The universal-safe form
    is ``_xlfn.FUNC`` (no ``_xlws.`` infix).

    **Pass 2** — add ``_xlfn.`` prefix to any DA function names that are still
    unprefixed after Pass 1 (i.e. names not preceded by ``.`` or ``_``).
    """
    if not formula_xml_text:
        return formula_xml_text

    # Pass 1: _xlfn._xlws.FUNC( → _xlfn.FUNC(
    def _strip_xlws(m: re.Match) -> str:
        func_name = m.group(1).upper()
        return _DA_FUNC_MAP.get(func_name, f"_xlfn.{func_name}") + "("

    formula_xml_text = _DA_XLWS_PREFIX_RE.sub(_strip_xlws, formula_xml_text)

    # Pass 2: add _xlfn. to any remaining unprefixed DA function names
    def _repl(m: re.Match) -> str:
        return _DA_FUNC_MAP[m.group(1).upper()] + "("

    return _DA_FUNC_PREFIX_RE.sub(_repl, formula_xml_text)


def _patch_f_tag(existing_f_attrs: str, anchor: str) -> str:
    """Return the ``<f ...>`` opening tag unchanged.

    Excel 365 dynamic-array formulas do NOT need ``t="array"`` / ``ref`` /
    ``aca`` / ``ca`` on the ``<f>`` element.  Those are legacy CSE
    (Ctrl+Shift+Enter) array-formula attributes.  Adding them to a
    dynamic-array formula confuses Excel's repair engine:

    * ``_xlfn._xlws.*`` functions (FILTER, SORT, SEQUENCE …) are recognised
      as spill-capable, so Excel tolerates the spurious ``t="array"``.
    * ``_xlfn.LET`` and other non-``_xlws.`` functions are *not* recognised
      as unconditionally spill-capable.  When Excel sees ``t="array"`` with
      ``ref="A1"`` on a formula that actually spills to, say, 10 × 3 cells
      it treats the mismatch as a broken CSE array formula and removes it
      ("Removed Records: Formula" repair dialog).

    The correct Excel 365 dynamic-array signature is:
      * ``cm="1"`` on the ``<c>`` element  (set by the caller)
      * ``_xlfn.`` / ``_xlfn._xlws.`` prefixes in the formula body
      * ``xl/metadata.xml`` with XLDAPR
    No special attributes are needed on ``<f>``.
    """
    return f"<f{existing_f_attrs}>"  # pass-through – do not add t/aca/ref/ca


def _add_cm_to_sheet_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    """
    Scan *xml_bytes* (a worksheet XML) for ``<c …><f …>…</f>`` elements that
    contain dynamic-array functions.

    For each matching cell:
    * Adds ``cm="1"`` to the ``<c>`` opening tag.
    * Rewrites formula body to add ``_xlfn.``/``_xlfn._xlws.`` prefixes for
      dynamic-array functions.
    * Does **not** modify ``<f>`` element attributes (no ``t``, ``ref``, etc.).

    Returns ``(patched_bytes, count_of_patched_cells)``.
    """
    text = xml_bytes.decode("utf-8")
    patched_count = 0

    def _patch_cell(m: re.Match) -> str:
        nonlocal patched_count
        cell_open = m.group("cell_open")   # e.g. '<c r="B7" s="23"'
        rest = m.group("rest")             # everything after the opening tag
        formula_body = m.group("formula")  # content of <f>…</f>

        if not _is_da_formula(formula_body):
            return m.group(0)  # unchanged

        # Rewrite formula body to add required _xlfn./_xlfn._xlws. prefixes.
        new_body = _rewrite_formula_body(formula_body)

        if 'cm="' in cell_open:
            # Cell was already tagged cm="1" by a previous patcher run.
            # Still need to:
            #  a) strip legacy CSE t="array"/aca/ref/ca attrs from <f> if an
            #     older patcher version left them behind, AND
            #  b) update formula body prefixes if they changed.
            full = m.group(0)
            healed, legacy_stripped = _strip_legacy_f_attrs(full)
            if legacy_stripped:
                # Count as a patched cell so callers know work was done.
                patched_count += 1
                if new_body != formula_body:
                    healed = healed.replace(
                        ">" + formula_body + "</f>",
                        ">" + new_body + "</f>",
                        1,
                    )
                return healed
            # No legacy attrs — only update formula body if needed.
            if new_body != formula_body:
                return full.replace(
                    ">" + formula_body + "</f>",
                    ">" + new_body + "</f>",
                    1,
                )
            return full

        # 1. Add cm="1" to the <c> opening tag
        patched_count += 1
        patched_cell_open = cell_open + ' cm="1"'

        # 2. Do NOT modify <f> attributes — cm="1" + _xlfn. prefixes +
        #    metadata.xml are sufficient for Excel 365 dynamic arrays.
        #    Adding t="array"/ref/aca/ca would make Excel treat the formula
        #    as a legacy CSE array formula and remove it for LET/XLOOKUP/etc.
        patched_rest = rest

        # Older files can carry legacy CSE attrs (t="array"/aca/ref/ca) on
        # dynamic-array formulas before cm="1" is introduced. Strip them in
        # this first-time cm path too, otherwise Excel can still repair/remove
        # the formula even after metadata patching.
        patched_rest, _ = _strip_legacy_f_attrs(patched_rest)

        # 3. Rewrite formula body (add _xlfn. prefixes)
        if new_body != formula_body:
            patched_rest = patched_rest.replace(
                ">" + formula_body + "</f>",
                ">" + new_body + "</f>",
                1,
            )

        return patched_cell_open + patched_rest

    # Match the <c> opening tag plus the first <f>…</f> inside it.
    # We use a non-greedy match to grab just the first formula element.
    pattern = re.compile(
        r"(?P<cell_open><c\b[^>]*)(?P<rest>[^<]*(?:<[^f/][^>]*>(?:[^<]*</[^>]+>)*[^<]*)*"
        r"<f[^>]*>(?P<formula>[^<]*)</f>)",
        re.DOTALL,
    )
    patched_text = pattern.sub(_patch_cell, text)
    return patched_text.encode("utf-8"), patched_count


# ── Workbook-level XML helpers ───────────────────────────────────────────────

def _ensure_metadata_rel(rels_xml: bytes) -> tuple[bytes, bool]:
    """
    Add the sheetMetadata relationship to workbook.xml.rels if missing.
    Returns ``(bytes, was_already_present)``.
    """
    text = rels_xml.decode("utf-8")
    if "sheetMetadata" in text:
        return rels_xml, True

    # Assign a new rId that doesn't clash — use rId101 as a safe upper value.
    # Find existing IDs and pick max+1.
    existing_ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', text)]
    new_id = (max(existing_ids) + 1) if existing_ids else 101

    new_rel = (
        f'<Relationship Id="rId{new_id}"'
        f' Type="{_REL_TYPE_METADATA}"'
        f' Target="metadata.xml"/>'
    )
    # Insert before </Relationships>
    patched = text.replace("</Relationships>", new_rel + "</Relationships>")
    return patched.encode("utf-8"), False


def _ensure_content_type(ct_xml: bytes) -> bytes:
    """
    Add the metadata.xml Override entry to [Content_Types].xml if missing.
    """
    text = ct_xml.decode("utf-8")
    if "metadata.xml" in text and "sheetMetadata" in text:
        return ct_xml

    new_override = (
        f'<Override PartName="/xl/metadata.xml"'
        f' ContentType="{_CONTENT_TYPE_METADATA}"/>'
    )
    # Insert before </Types>
    patched = text.replace("</Types>", new_override + "</Types>")
    return patched.encode("utf-8")


def _patch_workbook_xml(wb_xml: bytes) -> bytes:
    """
    Upgrade workbook.xml to signal Excel 365 compatibility.

    openpyxl omits ``<fileVersion>`` and writes an old ``calcId`` (124519 =
    Excel 2016).  Excel's repair engine uses these hints to determine whether
    dynamic-array metadata structures are valid.  Aligning them with Excel 365
    values prevents spurious "Removed Records: Formula" repair events.

    Changes applied when needed:
    * Add ``<fileVersion appName="xl" lastEdited="7" …/>`` after the
      ``<workbook>`` opening tag.
    * Replace openpyxl's default ``calcId="124519"`` with ``"191028"``
      (Excel 365).
    * Replace bare ``<workbookPr/>`` with one that carries
      ``defaultThemeVersion="124226"``.
    """
    text = wb_xml.decode("utf-8")
    changed = False

    # --- fileVersion ---
    if "fileVersion" not in text:
        fv = (
            '<fileVersion appName="xl" lastEdited="7"'
            ' lowestEdited="4" rupBuild="29530"/>'
        )
        # Insert right after the <workbook …> opening tag
        text = re.sub(r"(<workbook\b[^>]*>)", r"\1" + fv, text, count=1)
        changed = True

    # --- calcId (openpyxl writes 124519 = Office 2016; 191028 = Office 365) ---
    if 'calcId="124519"' in text:
        text = text.replace('calcId="124519"', 'calcId="191028"')
        changed = True

    # --- workbookPr defaultThemeVersion ---
    # Use a regex to handle both <workbookPr/> (Windows/older openpyxl) and
    # <workbookPr /> (Ubuntu/Linux — openpyxl serialises with a space before />).
    if re.search(r"<workbookPr\s*/>", text) and "defaultThemeVersion" not in text:
        text = re.sub(
            r"<workbookPr\s*/>",
            '<workbookPr defaultThemeVersion="124226"/>',
            text,
            count=1,
        )
        changed = True

    return text.encode("utf-8") if changed else wb_xml


# ── Main entry point ─────────────────────────────────────────────────────────

# ── COM Formula2 snapshot / restore ─────────────────────────────────────────


def snapshot_formula2_cells(path: str) -> dict[str, dict[str, str]]:
    """Pre-save: capture all COM Formula2 cells from the OOXML zip at *path*.

    A COM Formula2 cell is identified by ``aca="1"`` on its ``<f>`` element —
    the Excel 365 dynamic-array marker written by ``.Formula2`` in win32com.

    Returns::

        {sheet_zip_name: {cell_address: f_attrs_string}}

    where ``f_attrs_string`` is the full attribute string of the ``<f>``
    opening tag (e.g. ``' t="array" aca="1" ref="A7:K10" ca="1"'``).

    Returns an empty dict on any I/O or parse error — callers fall back
    safely to the standard ``cm="1"`` patching path.
    """
    result: dict[str, dict[str, str]] = {}
    try:
        with zipfile.ZipFile(path, "r") as zin:
            for name in zin.namelist():
                if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                    continue
                content = zin.read(name).decode("utf-8")
                cells: dict[str, str] = {}
                for m in _FORMULA2_CELL_RE.finditer(content):
                    addr = m.group(1)    # e.g. "A7"
                    f_attrs = m.group(2)  # e.g. ' t="array" aca="1" ref="A7:K10" ca="1"'
                    cells[addr] = f_attrs
                if cells:
                    result[name] = cells
    except Exception:  # pragma: no cover — I/O error guard
        return {}
    return result


def restore_formula2_f_attrs(
    path: str,
    snapshots: dict[str, dict[str, str]],
) -> int:
    """Post-patcher: re-inject COM Formula2 ``<f>`` attributes stripped by openpyxl.

    For each (sheet_zip_name, cell_address) in *snapshots*, locates the
    ``<f>`` element in the worksheet XML and replaces its opening tag with
    the original ``t="array" aca="1" ref="..." ca="1"`` attributes captured
    by :func:`snapshot_formula2_cells` before the openpyxl save.

    The ``cm="1"`` attribute on ``<c>`` (injected by the DA patcher) is
    preserved — this function only modifies the ``<f>`` opening tag.

    Returns the number of cells successfully restored, or 0 on any error.
    """
    if not snapshots:
        return 0

    try:
        with open(path, "rb") as fh:
            original_bytes = fh.read()
    except Exception:  # pragma: no cover
        return 0

    buf_in = io.BytesIO(original_bytes)
    total_restored = 0
    new_entries: dict[str, bytes] = {}

    try:
        with zipfile.ZipFile(buf_in, "r") as zin:
            all_names = zin.namelist()

            for sheet_name, cell_map in snapshots.items():
                if sheet_name not in all_names:
                    continue
                content = zin.read(sheet_name).decode("utf-8")
                changed = False

                for addr, original_f_attrs in cell_map.items():
                    # After patcher, the cell looks like:
                    #   <c r="A7" ... cm="1"><f>formula</f>...
                    # Replace the <f> opening tag with the original one.
                    cell_re = re.compile(
                        r"(<c\b[^>]*\br="
                        + '"'
                        + re.escape(addr)
                        + '"'
                        + r"[^>]*>[^<]*)"
                        r"(<f\b[^>]*>)",
                    )
                    replacement = r"\1" + f"<f{original_f_attrs}>"
                    new_content, n = cell_re.subn(replacement, content, count=1)
                    if n:
                        content = new_content
                        total_restored += 1
                        changed = True

                if changed:
                    new_entries[sheet_name] = content.encode("utf-8")

            if not new_entries:
                return 0

            # Rebuild the zip with the restored sheet XMLs.
            out_buf = io.BytesIO()
            with zipfile.ZipFile(
                out_buf, "w", compression=zipfile.ZIP_DEFLATED
            ) as zout:
                for name in all_names:
                    if name in new_entries:
                        zout.writestr(name, new_entries[name])
                    else:
                        zout.writestr(name, zin.read(name))

    except Exception:  # pragma: no cover — I/O error guard
        return 0

    out_bytes = out_buf.getvalue()
    try:
        with open(path, "wb") as fh:
            fh.write(out_bytes)
    except Exception:  # pragma: no cover
        return 0

    return total_restored


# ── Main entry point ─────────────────────────────────────────────────────────


def patch_dynamic_array_metadata(path: str) -> int:
    """
    Post-process the saved xlsx at *path* to add Excel 365 dynamic-array
    metadata so spill formulas (LET, FILTER, SORT, …) survive file open
    without triggering the Excel repair dialog.

    Returns the total number of cells that were patched.
    Does nothing (returns 0) if no dynamic-array formulas are found.
    """
    path = str(path)
    if not os.path.exists(path):
        return 0

    # Read the entire zip into memory so we can rewrite it atomically.
    with open(path, "rb") as fh:
        original_bytes = fh.read()

    buf = io.BytesIO(original_bytes)
    total_patched = 0
    new_entries: dict[str, bytes] = {}  # name → new bytes for modified entries

    with zipfile.ZipFile(buf, "r") as zin:
        names = zin.namelist()

        # ── 1. Patch worksheet XMLs ──────────────────────────────────────────
        for name in names:
            if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                raw = zin.read(name)
                patched, count = _add_cm_to_sheet_xml(raw)
                if patched != raw:   # any change: new cm OR formula body rewrite
                    new_entries[name] = patched
                    total_patched += count

        # ── 2. If cells were patched, wire up metadata infrastructure ────────
        if total_patched > 0:
            new_entries["xl/metadata.xml"] = _METADATA_XML.encode("utf-8")

            rels_name = "xl/_rels/workbook.xml.rels"
            rels_raw = zin.read(rels_name) if rels_name in names else b"<Relationships/>"
            rels_patched, _ = _ensure_metadata_rel(rels_raw)
            new_entries[rels_name] = rels_patched

            ct_raw = zin.read("[Content_Types].xml")
            new_entries["[Content_Types].xml"] = _ensure_content_type(ct_raw)

        # ── 3. workbook.xml — Excel 365 compatibility flags (always) ─────────
        #    Run even when cells were already patched (e.g. on a second save)
        #    so that fileVersion / calcId are always correct.
        wb_name = "xl/workbook.xml"
        if wb_name in names:
            wb_raw = zin.read(wb_name)
            wb_patched = _patch_workbook_xml(wb_raw)
            if wb_patched is not wb_raw:
                new_entries[wb_name] = wb_patched

        # ── 4. If there is nothing at all to change, short-circuit ───────────
        if not new_entries:
            return 0

        # ── 5. Rebuild the zip ───────────────────────────────────────────────
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name in new_entries:
                    zout.writestr(name, new_entries[name])
                else:
                    zout.writestr(name, zin.read(name))
            # Write newly added entries that weren't in the original
            for name, data in new_entries.items():
                if name not in names:
                    zout.writestr(name, data)

    out_bytes = out_buf.getvalue()
    with open(path, "wb") as fh:
        fh.write(out_bytes)

    return total_patched
