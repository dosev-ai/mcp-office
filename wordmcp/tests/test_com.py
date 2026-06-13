"""COM layer unit tests — groups A-D (Sprint B)."""
from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch

from wordmcp._com import _base as _com_base
from wordmcp.document_com import (
    _check_output_path,
    _word_app_context,
    export_as_pdf,
    accept_all_track_changes,
    reject_all_track_changes,
    list_tracked_changes,
    OfficeCOMError,
    _sanitize_revision_text,
)
from wordmcp.document_docx import ValidationError, NotAllowedError, WordMCPError

pytestmark = pytest.mark.com


# ── GROUP A: _word_app_context ───────────────────────────────────────────────

def test_word_app_context_sets_visible_false():
    mock_app = MagicMock()
    with patch("win32com.client.Dispatch", return_value=mock_app), \
         patch("pythoncom.CoInitialize"), \
         patch("pythoncom.CoUninitialize"):
        with _word_app_context() as app:
            assert app.Visible == False  # noqa: E712 — attribute-set check


def test_word_app_context_sets_display_alerts():
    mock_app = MagicMock()
    with patch("win32com.client.Dispatch", return_value=mock_app), \
         patch("pythoncom.CoInitialize"), \
         patch("pythoncom.CoUninitialize"):
        with _word_app_context() as app:
            assert app.DisplayAlerts == 0


def test_word_app_context_quit_called_on_clean_exit():
    mock_app = MagicMock()
    with patch("win32com.client.Dispatch", return_value=mock_app), \
         patch("pythoncom.CoInitialize"), \
         patch("pythoncom.CoUninitialize"):
        with _word_app_context():
            pass
    mock_app.Quit.assert_called_once()


def test_word_app_context_quit_called_on_exception():
    mock_app = MagicMock()
    with patch("win32com.client.Dispatch", return_value=mock_app), \
         patch("pythoncom.CoInitialize"), \
         patch("pythoncom.CoUninitialize"):
        with pytest.raises(RuntimeError):
            with _word_app_context():
                raise RuntimeError("body error")
    mock_app.Quit.assert_called()


def test_word_app_context_unavailable_raises(monkeypatch):
    monkeypatch.setattr(_com_base, "_COM_AVAILABLE", False)
    with pytest.raises(OfficeCOMError):
        with _word_app_context():
            pass


def test_word_app_context_coinitialize_called():
    mock_app = MagicMock()
    with patch("win32com.client.Dispatch", return_value=mock_app), \
         patch("pythoncom.CoInitialize") as mock_coinit, \
         patch("pythoncom.CoUninitialize"):
        with _word_app_context():
            pass
    mock_coinit.assert_called_once()


# ── GROUP B: export_as_pdf ───────────────────────────────────────────────────

def test_export_as_pdf_happy_path(patch_word_com, com_env, mock_word_doc):
    out = str(com_env.parent / "output.pdf")
    result = export_as_pdf(str(com_env), out, confirm=True)
    assert "output_path" in result
    mock_word_doc.ExportAsFixedFormat.assert_called_once()


def test_export_as_pdf_unc_rejected():
    with pytest.raises(ValidationError):
        _check_output_path(r"\\server\share\out.pdf")


def test_export_as_pdf_null_byte_in_output():
    with pytest.raises(ValidationError):
        _check_output_path("out\x00.pdf")


def test_export_as_pdf_confirm_false_blocked(monkeypatch):
    monkeypatch.setenv("WORD_ENABLE_COM", "true")
    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    with pytest.raises(ValidationError):
        export_as_pdf("fake.docx", "out.pdf", confirm=False)


def test_export_as_pdf_com_disabled(monkeypatch):
    monkeypatch.delenv("WORD_ENABLE_COM", raising=False)
    with pytest.raises(NotAllowedError):
        export_as_pdf("fake.docx", "out.pdf", confirm=True)


def test_export_as_pdf_non_pdf_extension(monkeypatch, tmp_path):
    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    with pytest.raises(ValidationError):
        _check_output_path(str(tmp_path / "out.docx"))


def test_export_as_pdf_com_error_wraps(patch_word_com, com_env, mock_word_doc):
    mock_word_doc.ExportAsFixedFormat.side_effect = OfficeCOMError("export boom")
    out = str(com_env.parent / "output.pdf")
    with pytest.raises(OfficeCOMError):
        export_as_pdf(str(com_env), out, confirm=True)


def test_export_as_pdf_locked_target_gives_actionable_error(
    patch_word_com, com_env, mock_word_doc, monkeypatch
):
    """COM error with 'access' in message -> locked-file OfficeCOMError with actionable text."""
    import wordmcp._com._export as _exp

    fake_exc = RuntimeError("access denied: file locked")
    monkeypatch.setattr(_exp, "_is_com_error", lambda _: True)
    monkeypatch.setattr(_exp, "_com_err_msg", lambda _: "access denied: file locked")
    mock_word_doc.ExportAsFixedFormat.side_effect = fake_exc

    out = str(com_env.parent / "output.pdf")
    with pytest.raises(OfficeCOMError, match="locked by another application"):
        export_as_pdf(str(com_env), out, confirm=True)


def test_export_as_pdf_opens_source_readonly(patch_word_com, com_env, mock_word_doc):
    """Source DOCX must be opened with ReadOnly=True to avoid locking it."""
    out = str(com_env.parent / "output.pdf")
    export_as_pdf(str(com_env), out, confirm=True)
    open_calls = patch_word_com.Documents.Open.call_args_list
    assert len(open_calls) == 1
    _args, kwargs = open_calls[0]
    assert kwargs.get("ReadOnly") is True


# ── GROUP C: accept_all_track_changes ────────────────────────────────────────

def test_accept_all_happy_path(patch_word_com, com_env, mock_word_doc):
    result = accept_all_track_changes(str(com_env), confirm=True)
    assert result == {"status": "ok", "accepted": True}
    mock_word_doc.Revisions.AcceptAll.assert_called_once()
    mock_word_doc.Save.assert_called_once()


def test_accept_all_confirm_false_blocked(monkeypatch):
    monkeypatch.setenv("WORD_ENABLE_COM", "true")
    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    with pytest.raises(ValidationError):
        accept_all_track_changes("fake.docx", confirm=False)


def test_accept_all_com_disabled(monkeypatch):
    monkeypatch.delenv("WORD_ENABLE_COM", raising=False)
    with pytest.raises(NotAllowedError):
        accept_all_track_changes("fake.docx", confirm=True)


def test_accept_all_write_disabled(monkeypatch):
    monkeypatch.setenv("WORD_ENABLE_COM", "true")
    monkeypatch.delenv("WORD_ENABLE_WRITE", raising=False)
    with pytest.raises(NotAllowedError):
        accept_all_track_changes("fake.docx", confirm=True)


def test_accept_all_com_error_wraps(patch_word_com, com_env, mock_word_doc):
    mock_word_doc.Revisions.AcceptAll.side_effect = OfficeCOMError("accept boom")
    with pytest.raises(OfficeCOMError):
        accept_all_track_changes(str(com_env), confirm=True)


def test_accept_all_closes_doc_on_error(patch_word_com, com_env, mock_word_doc):
    mock_word_doc.Revisions.AcceptAll.side_effect = Exception("accept boom")
    with pytest.raises(Exception):
        accept_all_track_changes(str(com_env), confirm=True)
    mock_word_doc.Close.assert_called()


# ── GROUP D: reject_all_track_changes ────────────────────────────────────────

def test_reject_all_happy_path(patch_word_com, com_env, mock_word_doc):
    result = reject_all_track_changes(str(com_env), confirm=True)
    assert result == {"status": "ok", "rejected": True}
    mock_word_doc.Revisions.RejectAll.assert_called_once()
    mock_word_doc.Save.assert_called_once()


def test_reject_all_confirm_false_blocked(monkeypatch):
    monkeypatch.setenv("WORD_ENABLE_COM", "true")
    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    with pytest.raises(ValidationError):
        reject_all_track_changes("fake.docx", confirm=False)


def test_reject_all_com_disabled(monkeypatch):
    monkeypatch.delenv("WORD_ENABLE_COM", raising=False)
    with pytest.raises(NotAllowedError):
        reject_all_track_changes("fake.docx", confirm=True)


def test_reject_all_write_disabled(monkeypatch):
    monkeypatch.setenv("WORD_ENABLE_COM", "true")
    monkeypatch.delenv("WORD_ENABLE_WRITE", raising=False)
    with pytest.raises(NotAllowedError):
        reject_all_track_changes("fake.docx", confirm=True)


def test_reject_all_com_error_wraps(patch_word_com, com_env, mock_word_doc):
    mock_word_doc.Revisions.RejectAll.side_effect = OfficeCOMError("reject boom")
    with pytest.raises(OfficeCOMError):
        reject_all_track_changes(str(com_env), confirm=True)


def test_reject_all_closes_doc_on_error(patch_word_com, com_env, mock_word_doc):
    mock_word_doc.Revisions.RejectAll.side_effect = Exception("reject boom")
    with pytest.raises(Exception):
        reject_all_track_changes(str(com_env), confirm=True)
    mock_word_doc.Close.assert_called()


# ── GROUP E: list_tracked_changes ─────────────────────────────────────────────

def test_list_tracked_happy_path(patch_word_com, com_env, mock_word_doc):
    result = list_tracked_changes(str(com_env))
    assert "revisions" in result
    assert "count" in result
    assert isinstance(result["revisions"], list)
    assert len(result["revisions"]) == 1
    rev = result["revisions"][0]
    for key in ("index", "author", "date", "type", "text"):
        assert key in rev, f"Missing key {key!r} in revision dict"


def test_list_tracked_sanitizes_control_chars(patch_word_com, com_env, mock_word_doc):
    bad_rev = MagicMock()
    bad_rev.Author = "Author"
    bad_rev.Date = "2025-01-01"
    bad_rev.Type = 1
    bad_rev.Range.Text = "\x01\x02 text"
    mock_word_doc.Revisions.__iter__ = MagicMock(return_value=iter([bad_rev]))

    result = list_tracked_changes(str(com_env))
    text = result["revisions"][0]["text"]
    assert "\x01" not in text
    assert "\x02" not in text
    assert "text" in text


def test_list_tracked_truncates_at_120(patch_word_com, com_env, mock_word_doc):
    long_rev = MagicMock()
    long_rev.Author = "Author"
    long_rev.Date = "2025-01-01"
    long_rev.Type = 1
    long_rev.Range.Text = "x" * 200
    mock_word_doc.Revisions.__iter__ = MagicMock(return_value=iter([long_rev]))

    result = list_tracked_changes(str(com_env))
    assert len(result["revisions"][0]["text"]) <= 120


def test_list_tracked_respects_max_items(patch_word_com, com_env, mock_word_doc):
    many_revs = []
    for i in range(5):
        rev = MagicMock()
        rev.Author = f"Author{i}"
        rev.Date = "2025-01-01"
        rev.Type = 1
        rev.Range.Text = f"text {i}"
        many_revs.append(rev)
    mock_word_doc.Revisions.__iter__ = MagicMock(return_value=iter(many_revs))

    result = list_tracked_changes(str(com_env), max_items=2)
    assert len(result["revisions"]) == 2
    assert result["count"] == 2


def test_list_tracked_com_disabled(monkeypatch):
    monkeypatch.delenv("WORD_ENABLE_COM", raising=False)
    with pytest.raises(NotAllowedError):
        list_tracked_changes("fake.docx")


def test_list_tracked_path_not_allowed(monkeypatch, tmp_path):
    monkeypatch.setenv("WORD_ENABLE_COM", "true")
    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    with pytest.raises((ValidationError, NotAllowedError)):
        list_tracked_changes("/outside/the/allowlist/file.docx")


def test_list_tracked_empty_revisions(patch_word_com, com_env, mock_word_doc):
    mock_word_doc.Revisions.__iter__ = MagicMock(return_value=iter([]))
    result = list_tracked_changes(str(com_env))
    assert result == {"revisions": [], "count": 0}


# ── GROUP F: gate enforcement ──────────────────────────────────────────────────

def test_write_gate_blocks_mutations(monkeypatch):
    """All mutation COM tools require WORD_ENABLE_WRITE; list_tracked does not."""
    monkeypatch.setenv("WORD_ENABLE_COM", "true")
    monkeypatch.delenv("WORD_ENABLE_WRITE", raising=False)
    with pytest.raises(NotAllowedError):
        export_as_pdf("fake.docx", "out.pdf", confirm=True)
    with pytest.raises(NotAllowedError):
        accept_all_track_changes("fake.docx", confirm=True)
    with pytest.raises(NotAllowedError):
        reject_all_track_changes("fake.docx", confirm=True)
    # list_tracked_changes must bypass the write gate entirely
    try:
        list_tracked_changes("fake.docx")
    except NotAllowedError as e:
        assert "WORD_ENABLE_WRITE" not in str(e), (
            f"list_tracked_changes must not gate on WORD_ENABLE_WRITE: {e}"
        )
    except Exception:
        pass  # path/allowlist error is expected — write gate was not hit


def test_list_tracked_does_not_require_write_gate(
    monkeypatch, patch_word_com, tmp_path
):
    """list_tracked_changes succeeds with COM enabled but no WORD_ENABLE_WRITE."""
    from docx import Document as _Doc
    monkeypatch.setenv("WORD_ENABLE_COM", "true")
    monkeypatch.delenv("WORD_ENABLE_WRITE", raising=False)
    monkeypatch.setenv("WORD_ALLOWLIST_ROOTS", str(tmp_path))
    docx_file = tmp_path / "ltc_nowrite.docx"
    _Doc().save(str(docx_file))
    result = list_tracked_changes(str(docx_file))
    assert "revisions" in result


def test_list_tracked_does_not_require_confirm(patch_word_com, com_env):
    """list_tracked_changes takes no confirm param and does not raise for omitting it."""
    result = list_tracked_changes(str(com_env))
    assert "revisions" in result
    assert "count" in result


def test_confirm_gate_blocks_mutations(monkeypatch):
    """All mutation COM tools raise ValidationError when confirm=False."""
    monkeypatch.setenv("WORD_ENABLE_COM", "true")
    monkeypatch.setenv("WORD_ENABLE_WRITE", "true")
    with pytest.raises(ValidationError):
        export_as_pdf("fake.docx", "out.pdf", confirm=False)
    with pytest.raises(ValidationError):
        accept_all_track_changes("fake.docx", confirm=False)
    with pytest.raises(ValidationError):
        reject_all_track_changes("fake.docx", confirm=False)


@pytest.mark.parametrize("fn,args,kwargs", [
    pytest.param(export_as_pdf, ("fake.docx", "out.pdf"), {"confirm": True}, id="export_as_pdf"),
    pytest.param(accept_all_track_changes, ("fake.docx",), {"confirm": True}, id="accept_all"),
    pytest.param(reject_all_track_changes, ("fake.docx",), {"confirm": True}, id="reject_all"),
    pytest.param(list_tracked_changes, ("fake.docx",), {}, id="list_tracked"),
])
def test_all_four_require_com_gate(monkeypatch, fn, args, kwargs):
    """All four COM tools raise NotAllowedError when WORD_ENABLE_COM is not set."""
    monkeypatch.delenv("WORD_ENABLE_COM", raising=False)
    with pytest.raises(NotAllowedError):
        fn(*args, **kwargs)


# ── GROUP G: error taxonomy and edge cases ─────────────────────────────────────

def test_office_com_error_subclasses_word_mcp_error():
    assert isinstance(OfficeCOMError("x"), WordMCPError)


def test_office_com_error_subclasses_word_mcp_error_type():
    assert issubclass(OfficeCOMError, WordMCPError)


def test_sanitizer_strips_low_control_chars():
    """\x00, \x01, \x07, \x08 are all stripped by _sanitize_revision_text."""
    result = _sanitize_revision_text("\x00\x01\x07\x08")
    assert result == ""


def test_sanitizer_strips_mid_control_chars():
    """\x0b, \x0e, \x1f, \x7f stripped; \t and \n kept; \r replaced with space."""
    result = _sanitize_revision_text("\x0b\x0e\x1f\x7f")
    assert result == ""

    kept = _sanitize_revision_text("\x09\x0a\x0d")
    assert "\x09" in kept   # tab preserved
    assert "\x0a" in kept   # LF preserved
    assert "\x0d" not in kept  # CR replaced with space
    assert " " in kept      # the space that CR became


def test_max_items_zero_returns_empty(patch_word_com, com_env, mock_word_doc):
    result = list_tracked_changes(str(com_env), max_items=0)
    assert result == {"revisions": [], "count": 0}


def test_com_error_at_close_suppressed(patch_word_com, com_env, mock_word_doc):
    """doc.Close raising must be suppressed — accept_all still returns successfully."""
    mock_word_doc.Close.side_effect = Exception("close failed")
    result = accept_all_track_changes(str(com_env), confirm=True)
    assert result == {"status": "ok", "accepted": True}


def test_check_output_path_rejects_unc_double_slash():
    """UNC paths using forward slashes (//) must be rejected before Path() is called."""
    with pytest.raises(ValidationError):
        _check_output_path("//server/share/out.pdf")
