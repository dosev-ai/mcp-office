"""
Read-only COM operations: tracked change listing and text sanitization.

No FastMCP/MCP imports.
"""
from __future__ import annotations

import re

from wordmcp._com._base import (
    OfficeCOMError,
    _check_com_enabled,
    _check_path,
    _is_com_error,
    _com_err_msg,
    _word_app_context,
)

# Matches Word control characters unsafe for JSON transport.
# Excludes \t (\x09), \n (\x0a), \r (\x0d) — handled separately below.
_CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0e-\x1f\x7f]")


def _sanitize_revision_text(s: str) -> str:
    """Strip Word control characters unsafe for JSON transport; \\r becomes a space."""
    return _CTRL_RE.sub("", s).replace("\r", " ")


def list_tracked_changes(path: str, max_items: int = 200) -> dict:
    """List tracked changes (revisions) in a .docx file via Word COM.

    Read-only — no write gate or confirm required.
    Gate order: _check_com_enabled → _check_path.

    Sanitization (BLOCKER-1): both revision text and author are passed through
    _sanitize_revision_text() to strip Word control characters before JSON
    serialisation. \\r is replaced with a space so paragraph marks stay readable.

    Args:
        path:      Absolute path to the .docx file.
        max_items: Maximum number of revisions to return (default 200).

    Returns:
        {"revisions": [...], "count": int}
    """
    import contextlib

    _check_com_enabled()
    checked_path = _check_path(path)
    revisions: list[dict] = []
    doc = None
    with _word_app_context() as app:
        try:
            doc = app.Documents.Open(str(checked_path), ReadOnly=True)
            for i, rev in enumerate(doc.Revisions):
                if i >= max_items:
                    break
                # BLOCKER-1: sanitize both text and author before JSON transport
                text = _sanitize_revision_text(rev.Range.Text or "")[:120]
                author = _sanitize_revision_text(rev.Author or "")
                revisions.append(
                    {
                        "index": i,
                        "author": author,
                        "date": str(rev.Date),
                        "type": int(rev.Type),
                        "text": text,
                    }
                )
            return {"revisions": revisions, "count": len(revisions)}
        except OfficeCOMError:
            raise
        except Exception as exc:
            if _is_com_error(exc):
                raise OfficeCOMError(f"List revisions failed: {_com_err_msg(exc)}") from exc
            raise
        finally:
            if doc is not None:
                with contextlib.suppress(Exception):
                    doc.Close(False)
