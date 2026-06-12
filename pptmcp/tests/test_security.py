"""
Security tests for pptmcp — URL scheme validation, write-gate matrix, path traversal.
All tests are @pytest.mark.unit (no live server, no COM required).
"""
import pytest
from pptmcp.presentation_pptx import (
    manage_hyperlinks,
    add_hyperlink,
    add_slide,
    create_presentation,
    copy_slide,
    delete_slide,
    reorder_slides,
    add_shape,
    add_textbox,
    add_table_to_slide,
    insert_image,
    set_text_format,
    set_paragraph_format,
    replace_slide_text,
    edit_text_placeholder,
    set_speaker_notes,
    save,
    NotAllowedError,
    ValidationError,
)


@pytest.mark.unit
class TestManageHyperlinksURLSchemes:
    """G-URL-MH: manage_hyperlinks rejects dangerous URL schemes."""

    def test_javascript_blocked(self, write_env, sample_pptx_file):
        with pytest.raises(ValidationError, match="scheme"):
            manage_hyperlinks(sample_pptx_file, 0, "add", shape_id=1,
                              target_url="javascript:alert(1)", confirm=True)

    def test_data_blocked(self, write_env, sample_pptx_file):
        with pytest.raises(ValidationError, match="scheme"):
            manage_hyperlinks(sample_pptx_file, 0, "add", shape_id=1,
                              target_url="data:text/html,<h1>x</h1>", confirm=True)

    def test_file_blocked(self, write_env, sample_pptx_file):
        with pytest.raises(ValidationError, match="scheme"):
            manage_hyperlinks(sample_pptx_file, 0, "add", shape_id=1,
                              target_url="file:///etc/passwd", confirm=True)

    def test_vbscript_blocked(self, write_env, sample_pptx_file):
        with pytest.raises(ValidationError, match="scheme"):
            manage_hyperlinks(sample_pptx_file, 0, "add", shape_id=1,
                              target_url="vbscript:MsgBox(1)", confirm=True)

    def test_ftp_blocked(self, write_env, sample_pptx_file):
        with pytest.raises(ValidationError, match="scheme"):
            manage_hyperlinks(sample_pptx_file, 0, "add", shape_id=1,
                              target_url="ftp://example.com/file", confirm=True)

    def test_https_allowed(self, write_env, sample_pptx_file):
        # Should not raise ValidationError for scheme (may raise for shape not found — that's OK)
        try:
            manage_hyperlinks(sample_pptx_file, 0, "add", shape_id=1,
                              target_url="https://example.com", confirm=True)
        except ValidationError as e:
            assert "scheme" not in str(e).lower()

    def test_allows_mailto(self, write_env, sample_pptx_file):
        # mailto: is an allowed scheme — must NOT raise scheme ValidationError
        try:
            manage_hyperlinks(sample_pptx_file, 0, "add", shape_id=1,
                              target_url="mailto:user@example.com", confirm=True)
        except ValidationError as e:
            assert "scheme" not in str(e).lower()


@pytest.mark.unit
class TestAddHyperlinkURLSchemes:
    """G-URL-AH: add_hyperlink rejects dangerous URL schemes."""

    def test_javascript_blocked(self, write_env, sample_pptx_file):
        with pytest.raises(ValidationError, match="scheme"):
            add_hyperlink(sample_pptx_file, 0, 0, 0, "javascript:alert(1)", confirm=True)

    def test_data_blocked(self, write_env, sample_pptx_file):
        with pytest.raises(ValidationError, match="scheme"):
            add_hyperlink(sample_pptx_file, 0, 0, 0, "data:text/html,<h1>x</h1>", confirm=True)

    def test_file_blocked(self, write_env, sample_pptx_file):
        with pytest.raises(ValidationError, match="scheme"):
            add_hyperlink(sample_pptx_file, 0, 0, 0, "file:///etc/passwd", confirm=True)

    def test_vbscript_blocked(self, write_env, sample_pptx_file):
        with pytest.raises(ValidationError, match="scheme"):
            add_hyperlink(sample_pptx_file, 0, 0, 0, "vbscript:MsgBox(1)", confirm=True)

    def test_ftp_blocked(self, write_env, sample_pptx_file):
        with pytest.raises(ValidationError, match="scheme"):
            add_hyperlink(sample_pptx_file, 0, 0, 0, "ftp://example.com/file", confirm=True)

    def test_https_allowed(self, write_env, sample_pptx_file):
        try:
            add_hyperlink(sample_pptx_file, 0, 0, 0, "https://example.com", confirm=True)
        except ValidationError as e:
            assert "scheme" not in str(e).lower()

    def test_allows_mailto(self, write_env, sample_pptx_file):
        try:
            add_hyperlink(sample_pptx_file, 0, 0, 0, "mailto:user@example.com", confirm=True)
        except ValidationError as e:
            assert "scheme" not in str(e).lower()


@pytest.mark.unit
class TestWriteGateMatrix:
    """G-GATE: All write tools require PPT_ENABLE_WRITE=true AND confirm=True."""
    # 34 tests: 17 tools × 2 conditions (enable_false, confirm_false)
    # PLUS 2 extra: manage_hyperlinks operation="remove" × 2 (B-5 fix)

    # --- create_presentation ---
    def test_create_presentation_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            create_presentation(str(allowlisted_dir / "new.pptx"), confirm=True)

    def test_create_presentation_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            create_presentation(str(allowlisted_dir / "new.pptx"), confirm=False)

    # --- add_slide ---
    def test_add_slide_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            add_slide(str(allowlisted_dir / "x.pptx"), confirm=True)

    def test_add_slide_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            add_slide(str(allowlisted_dir / "x.pptx"), confirm=False)

    # --- copy_slide ---
    def test_copy_slide_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            copy_slide(str(allowlisted_dir / "x.pptx"), 0, str(allowlisted_dir / "x.pptx"), confirm=True)

    def test_copy_slide_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            copy_slide(str(allowlisted_dir / "x.pptx"), 0, str(allowlisted_dir / "x.pptx"), confirm=False)

    # --- delete_slide ---
    def test_delete_slide_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            delete_slide(str(allowlisted_dir / "x.pptx"), 0, confirm=True)

    def test_delete_slide_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            delete_slide(str(allowlisted_dir / "x.pptx"), 0, confirm=False)

    # --- reorder_slides ---
    def test_reorder_slides_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            reorder_slides(str(allowlisted_dir / "x.pptx"), [0, 1], confirm=True)

    def test_reorder_slides_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            reorder_slides(str(allowlisted_dir / "x.pptx"), [0, 1], confirm=False)

    # --- add_shape ---
    def test_add_shape_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            add_shape(str(allowlisted_dir / "x.pptx"), 0, "rectangle", 0, 0, 100, 100, confirm=True)

    def test_add_shape_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            add_shape(str(allowlisted_dir / "x.pptx"), 0, "rectangle", 0, 0, 100, 100, confirm=False)

    # --- add_textbox ---
    def test_add_textbox_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            add_textbox(str(allowlisted_dir / "x.pptx"), 0, 1.0, 1.0, 4.0, 1.0, "Hello", confirm=True)

    def test_add_textbox_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            add_textbox(str(allowlisted_dir / "x.pptx"), 0, 1.0, 1.0, 4.0, 1.0, "Hello", confirm=False)

    # --- add_table_to_slide ---
    def test_add_table_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            add_table_to_slide(str(allowlisted_dir / "x.pptx"), 0, 2, 2, 1.0, 1.0, 4.0, 2.0, confirm=True)

    def test_add_table_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            add_table_to_slide(str(allowlisted_dir / "x.pptx"), 0, 2, 2, 1.0, 1.0, 4.0, 2.0, confirm=False)

    # --- insert_image ---
    def test_insert_image_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            insert_image(str(allowlisted_dir / "x.pptx"), 0, str(allowlisted_dir / "img.png"), confirm=True)

    def test_insert_image_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            insert_image(str(allowlisted_dir / "x.pptx"), 0, str(allowlisted_dir / "img.png"), confirm=False)

    # --- set_text_format ---
    def test_set_text_format_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            set_text_format(str(allowlisted_dir / "x.pptx"), 0, 0, confirm=True)

    def test_set_text_format_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            set_text_format(str(allowlisted_dir / "x.pptx"), 0, 0, confirm=False)

    # --- set_paragraph_format ---
    def test_set_paragraph_format_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            set_paragraph_format(str(allowlisted_dir / "x.pptx"), 0, 0, confirm=True)

    def test_set_paragraph_format_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            set_paragraph_format(str(allowlisted_dir / "x.pptx"), 0, 0, confirm=False)

    # --- replace_slide_text ---
    def test_replace_slide_text_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            replace_slide_text(str(allowlisted_dir / "x.pptx"), "old", "new", confirm=True)

    def test_replace_slide_text_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            replace_slide_text(str(allowlisted_dir / "x.pptx"), "old", "new", confirm=False)

    # --- edit_text_placeholder ---
    def test_edit_text_placeholder_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            edit_text_placeholder(str(allowlisted_dir / "x.pptx"), 0, 0, "text", confirm=True)

    def test_edit_text_placeholder_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            edit_text_placeholder(str(allowlisted_dir / "x.pptx"), 0, 0, "text", confirm=False)

    # --- set_speaker_notes ---
    def test_set_speaker_notes_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            set_speaker_notes(str(allowlisted_dir / "x.pptx"), 0, "notes", confirm=True)

    def test_set_speaker_notes_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            set_speaker_notes(str(allowlisted_dir / "x.pptx"), 0, "notes", confirm=False)

    # --- save ---
    def test_save_presentation_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            save(str(allowlisted_dir / "x.pptx"), confirm=True)

    def test_save_presentation_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            save(str(allowlisted_dir / "x.pptx"), confirm=False)

    # --- manage_hyperlinks operation="add" ---
    def test_manage_hyperlinks_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            manage_hyperlinks(str(allowlisted_dir / "x.pptx"), 0, "add",
                              shape_id=1, target_url="https://example.com", confirm=True)

    def test_manage_hyperlinks_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            manage_hyperlinks(str(allowlisted_dir / "x.pptx"), 0, "add",
                              shape_id=1, target_url="https://example.com", confirm=False)

    # --- manage_hyperlinks operation="remove" (B-5 fix) ---
    def test_manage_hyperlinks_remove_enable_false(self, no_write_env, allowlisted_dir):
        """B-5: remove op must be gated even without PPT_ENABLE_WRITE."""
        with pytest.raises(NotAllowedError):
            manage_hyperlinks(str(allowlisted_dir / "x.pptx"), 0, "remove",
                              shape_id=1, confirm=True)

    def test_manage_hyperlinks_remove_confirm_false(self, write_env, allowlisted_dir):
        """B-5: remove op must require confirm=True."""
        with pytest.raises((NotAllowedError, ValidationError)):
            manage_hyperlinks(str(allowlisted_dir / "x.pptx"), 0, "remove",
                              shape_id=1, confirm=False)

    # --- add_hyperlink ---
    def test_add_hyperlink_enable_false(self, no_write_env, allowlisted_dir):
        with pytest.raises(NotAllowedError):
            add_hyperlink(str(allowlisted_dir / "x.pptx"), 0, 0, 0, "https://example.com", confirm=True)

    def test_add_hyperlink_confirm_false(self, write_env, allowlisted_dir):
        with pytest.raises((NotAllowedError, ValidationError)):
            add_hyperlink(str(allowlisted_dir / "x.pptx"), 0, 0, 0, "https://example.com", confirm=False)


@pytest.mark.unit
class TestPathSecurity:
    """G-PATH: Path traversal, outside-allowlist, and null-byte rejection."""

    def test_path_traversal_blocked(self, allowlisted_dir):
        traversal = str(allowlisted_dir) + "/../../etc/passwd.pptx"
        with pytest.raises((NotAllowedError, ValidationError)):
            from pptmcp.presentation_pptx import read_slide
            read_slide(traversal, 0)

    def test_absolute_outside_allowlist_blocked(self, allowlisted_dir):
        outside = "C:/Windows/System32/evil.pptx"
        with pytest.raises((NotAllowedError, ValidationError)):
            from pptmcp.presentation_pptx import read_slide
            read_slide(outside, 0)

    def test_null_byte_in_path_blocked(self, allowlisted_dir):
        null_path = str(allowlisted_dir / "file\x00.pptx")
        with pytest.raises((NotAllowedError, ValidationError)):
            from pptmcp.presentation_pptx import read_slide
            read_slide(null_path, 0)


@pytest.mark.unit
class TestStdoutDiscipline:
    """G-STDOUT: Calling tools must not write to stdout."""

    def test_no_stdout_on_read(self, capsys, allowlisted_dir):
        from pptmcp.presentation_pptx import read_slide
        # Will raise (file not found / not allowed) but must not print to stdout
        try:
            read_slide(str(allowlisted_dir / "x.pptx"), 0)
        except Exception:
            pass
        captured = capsys.readouterr()
        assert captured.out == ""
