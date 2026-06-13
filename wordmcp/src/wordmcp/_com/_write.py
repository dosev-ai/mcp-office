"""
Write COM operations: accept/reject tracked changes and manage_tracked_changes router.

No FastMCP/MCP imports.
"""
from __future__ import annotations

import contextlib

from wordmcp._com._base import (
    OfficeCOMError,
    ValidationError,
    _check_com_enabled,
    _check_path,
    _check_write,
    _check_confirm,
    _is_com_error,
    _com_err_msg,
    _word_app_context,
)
from wordmcp._com._read import list_tracked_changes  # explicit, NOT via _com package to avoid circular


def accept_all_track_changes(path: str, confirm: bool = False) -> dict:
    """Accept all tracked changes in a .docx file via Word COM and save.

    Gate order: _check_com_enabled → _check_write → _check_confirm → _check_path.

    Args:
        path:    Absolute path to the .docx file.
        confirm: Must be True to authorise the write operation.

    Returns:
        {"status": "ok", "accepted": True}
    """
    _check_com_enabled()
    _check_write()
    _check_confirm(confirm)
    checked_path = _check_path(path)
    doc = None
    with _word_app_context() as app:
        try:
            doc = app.Documents.Open(str(checked_path))
            doc.Revisions.AcceptAll()
            doc.Save()
            return {"status": "ok", "accepted": True}
        except OfficeCOMError:
            raise
        except Exception as exc:
            if _is_com_error(exc):
                raise OfficeCOMError(f"Accept revisions failed: {_com_err_msg(exc)}") from exc
            raise
        finally:
            if doc is not None:
                with contextlib.suppress(Exception):
                    doc.Close(False)


def reject_all_track_changes(path: str, confirm: bool = False) -> dict:
    """Reject all tracked changes in a .docx file via Word COM and save.

    Gate order: _check_com_enabled → _check_write → _check_confirm → _check_path.

    Args:
        path:    Absolute path to the .docx file.
        confirm: Must be True to authorise the write operation.

    Returns:
        {"status": "ok", "rejected": True}
    """
    _check_com_enabled()
    _check_write()
    _check_confirm(confirm)
    checked_path = _check_path(path)
    doc = None
    with _word_app_context() as app:
        try:
            doc = app.Documents.Open(str(checked_path))
            doc.Revisions.RejectAll()
            doc.Save()
            return {"status": "ok", "rejected": True}
        except OfficeCOMError:
            raise
        except Exception as exc:
            if _is_com_error(exc):
                raise OfficeCOMError(f"Reject revisions failed: {_com_err_msg(exc)}") from exc
            raise
        finally:
            if doc is not None:
                with contextlib.suppress(Exception):
                    doc.Close(False)


def manage_tracked_changes(
    path: str,
    operation: str,
    revision_index: int | None = None,
    confirm: bool = False,
) -> dict:
    """Route tracked-change operations to the appropriate COM function.

    Operations:
      'list'        → list_tracked_changes(path)
      'accept_all'  → accept_all_track_changes(path, confirm)
      'reject_all'  → reject_all_track_changes(path, confirm)
      'accept_one'  — accept individual revision by index
      'reject_one'  — reject individual revision by index
      unknown       → raises ValidationError
    """
    if operation == "list":
        return list_tracked_changes(path)
    elif operation == "accept_all":
        return accept_all_track_changes(path, confirm)
    elif operation == "reject_all":
        return reject_all_track_changes(path, confirm)
    elif operation == "accept_one":
        # H-006: canonical gate order is _check_com_enabled → _check_write → _check_confirm
        _check_com_enabled()
        _check_write()
        _check_confirm(confirm)
        checked_path = _check_path(path)
        doc = None
        with _word_app_context() as app:
            try:
                doc = app.Documents.Open(str(checked_path))
                count = doc.Revisions.Count
                if revision_index is None or revision_index < 0 or revision_index >= count:
                    raise OfficeCOMError(
                        f"revision_index out of range 0–{count - 1}"
                    )
                doc.Revisions(revision_index + 1).Accept()  # COM is 1-based
                doc.Save()
                return {
                    "status": "ok",
                    "operation": "accept_one",
                    "revision_index": revision_index,
                }
            except OfficeCOMError:
                raise
            except Exception as exc:
                if _is_com_error(exc):
                    raise OfficeCOMError(f"Accept revision failed: {_com_err_msg(exc)}") from exc
                raise
            finally:
                if doc is not None:
                    with contextlib.suppress(Exception):
                        doc.Close(False)
    elif operation == "reject_one":
        # H-006: canonical gate order is _check_com_enabled → _check_write → _check_confirm
        _check_com_enabled()
        _check_write()
        _check_confirm(confirm)
        checked_path = _check_path(path)
        doc = None
        with _word_app_context() as app:
            try:
                doc = app.Documents.Open(str(checked_path))
                count = doc.Revisions.Count
                if revision_index is None or revision_index < 0 or revision_index >= count:
                    raise OfficeCOMError(
                        f"revision_index out of range 0–{count - 1}"
                    )
                doc.Revisions(revision_index + 1).Reject()  # COM is 1-based
                doc.Save()
                return {
                    "status": "ok",
                    "operation": "reject_one",
                    "revision_index": revision_index,
                }
            except OfficeCOMError:
                raise
            except Exception as exc:
                if _is_com_error(exc):
                    raise OfficeCOMError(f"Reject revision failed: {_com_err_msg(exc)}") from exc
                raise
            finally:
                if doc is not None:
                    with contextlib.suppress(Exception):
                        doc.Close(False)
    else:
        raise ValidationError(
            f"Unknown operation: {operation!r}. "
            "Valid: list, accept_all, reject_all, accept_one, reject_one"
        )
