"""
wordmcp._com — COM automation sub-package for wordmcp.

Explicit named re-exports of every public and private symbol from the
four sub-modules. Do NOT use star imports here.
"""
from wordmcp._com._base import (
    _COM_AVAILABLE,
    _is_com_error,
    _com_err_msg,
    _check_com_enabled,
    _word_app_context,
    OfficeCOMError,
    ValidationError,
    NotAllowedError,
    _check_confirm,
    _check_path,
    _check_write,
)
from wordmcp._com._read import (
    _CTRL_RE,
    _sanitize_revision_text,
    list_tracked_changes,
)
from wordmcp._com._write import (
    accept_all_track_changes,
    reject_all_track_changes,
    manage_tracked_changes,
)
from wordmcp._com._export import (
    _check_output_path,
    _check_output_path_html,
    export_as_pdf,
    export_document,
)

__all__ = [
    "_COM_AVAILABLE",
    "_is_com_error",
    "_com_err_msg",
    "_check_com_enabled",
    "_word_app_context",
    "OfficeCOMError",
    "ValidationError",
    "NotAllowedError",
    "_check_confirm",
    "_check_path",
    "_check_write",
    "_CTRL_RE",
    "_sanitize_revision_text",
    "list_tracked_changes",
    "accept_all_track_changes",
    "reject_all_track_changes",
    "manage_tracked_changes",
    "_check_output_path",
    "_check_output_path_html",
    "export_as_pdf",
    "export_document",
]
