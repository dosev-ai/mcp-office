"""MCP tool registration and public-safe error translation helpers."""
from __future__ import annotations

import functools
import logging

from fastmcp.exceptions import ToolError

logger = logging.getLogger("mailmcp.tools")


def _safe(fn, *args, **kwargs):
    try:
        return fn(*args, **kwargs)
    except PermissionError as exc:
        logger.error("PermissionError in tool: %s", exc)
        raise ToolError(str(exc)) from exc
    except ValueError as exc:
        logger.error("ValueError in tool: %s", exc)
        raise ToolError(str(exc)) from exc
    except TimeoutError as exc:
        logger.error("TimeoutError in tool: %s", exc)
        raise ToolError(
            "Outlook operation timed out. Outlook may be syncing or unresponsive. "
            "Retry, or open Outlook manually and try again."
        ) from exc
    except RuntimeError as exc:
        logger.error("RuntimeError in tool: %s", exc)
        raise ToolError(f"Operation failed ({type(exc).__name__})") from None
    except Exception as exc:
        logger.exception("Unexpected error in tool")
        raise ToolError(f"Unexpected error: {type(exc).__name__}") from None


def make_register(mcp):
    def _register(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            return _safe(fn, *args, **kwargs)
        mcp.tool()(wrapper)
    return _register
