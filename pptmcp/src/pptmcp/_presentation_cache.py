"""
Presentation cache and atomic-save helpers for pptmcp.

No FastMCP / MCP imports — pure python-pptx + stdlib only.
Imported by presentation_pptx.py (facade), _presentation_read.py,
and _presentation_write.py to avoid circular imports.
"""
from __future__ import annotations

import concurrent.futures
import contextlib
import logging
import os
import tempfile
from pathlib import Path
from typing import Any

from pptx import Presentation

from pptmcp._presentation_governance import OfficeCOMError

logger = logging.getLogger(__name__)

_LOAD_TIMEOUT_SECONDS = int(os.getenv("PPT_LOAD_TIMEOUT_SECONDS", "30"))

# ---------------------------------------------------------------------------
# Presentation in-memory cache (fixed-size, insertion-order eviction)
# Plain dict with a max-size cap; when the cap is reached the oldest
# inserted entry is evicted (dict insertion order is guaranteed in Python
# 3.7+). This is NOT an LRU cache — recently-read entries are not promoted.
# Exposed at module level so presentation_pptx.py can re-export and tests
# can patch pptmcp._presentation_cache._PRS_CACHE_MAX directly.
# ---------------------------------------------------------------------------

_prs_cache: dict[str, Presentation] = {}
_PRS_CACHE_MAX = int(os.getenv("PPT_CACHE_MAX", "20"))


def _load_prs(path: Path) -> Presentation:
    """Load a presentation from cache or disk.

    Reads _PRS_CACHE_MAX from pptmcp.presentation_pptx at call time so that
    tests patching ``pptmcp.presentation_pptx._PRS_CACHE_MAX`` take effect.
    The facade re-exports this module's _PRS_CACHE_MAX, and the deferred
    import below reads the patched value from the facade's namespace.
    """
    import pptmcp.presentation_pptx as _facade  # noqa: PLC0415

    key = str(path)
    if key not in _prs_cache:
        if len(_prs_cache) >= _facade._PRS_CACHE_MAX:
            _prs_cache.pop(next(iter(_prs_cache)))  # evict oldest (insertion-order)
        try:
            if not path.exists():
                raise OfficeCOMError(
                    f"Cannot open presentation — file not found: {path.name}. "
                    "If this path is on OneDrive or a network drive, check that the "
                    "file is synced locally and the MCP server has access to it."
                )
            # Run Presentation() in a thread so we can enforce a timeout.
            # OneDrive's sync engine can hold a brief write lock that causes
            # python-pptx's ZIP open to stall indefinitely otherwise.
            with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
                future = executor.submit(_facade.Presentation, str(path))
                try:
                    _prs_cache[key] = future.result(timeout=_LOAD_TIMEOUT_SECONDS)
                except concurrent.futures.TimeoutError:
                    raise OfficeCOMError(
                        f"Timed out opening presentation after {_LOAD_TIMEOUT_SECONDS}s: "
                        f"{path.name}. The file may be locked by OneDrive sync or another "
                        "process. Try again in a moment, or copy the file to C:\\Temp first."
                    )
        except OfficeCOMError:
            raise
        except PermissionError as e:
            logger.error("Permission denied opening presentation %s: %s", path, e)
            raise OfficeCOMError(
                f"Cannot open presentation — permission denied: {path.name}. "
                "The MCP server may be sandboxed (e.g. Claude Desktop) and unable "
                "to access this path. Try copying the file to a local directory "
                "such as C:\\Temp and opening it from there."
            ) from e
        except Exception as e:
            logger.error("Cannot open presentation %s: %s", path, e)
            raise OfficeCOMError(
                "Cannot open presentation — file may be corrupt or unreadable"
            ) from e
    return _prs_cache[key]


def _evict_prs(path: Path) -> None:
    """Remove a presentation from the in-memory cache."""
    _prs_cache.pop(str(path), None)


def _atomic_save(presentation: Any, dest: Path) -> None:
    """Save *presentation* to *dest* atomically via a sibling temp file.

    Uses ``tempfile.mkstemp`` in the same directory so ``os.replace`` is
    guaranteed to be atomic on POSIX and as-atomic-as-possible on NTFS
    (both paths on the same volume).
    """
    fd, tmp_path = tempfile.mkstemp(dir=str(dest.parent), suffix=".tmp")
    try:
        os.close(fd)
        presentation.save(tmp_path)
        os.replace(tmp_path, dest)
    except Exception:
        with contextlib.suppress(OSError):
            os.unlink(tmp_path)
        raise
