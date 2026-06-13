"""
Error taxonomy for pptmcp.

No FastMCP / MCP imports — pure stdlib only.
These classes are defined here and re-exported by presentation_pptx.py
so all modules in pptmcp can import errors without circular dependencies.
"""
from __future__ import annotations


class PPTMCPError(Exception):
    """Base class for all PowerPoint MCP errors."""


class ValidationError(PPTMCPError):
    """Input did not pass allowlist / type / format validation."""


class NotAllowedError(PPTMCPError):
    """Operation blocked by safety gate (PPT_ENABLE_WRITE or confirm=False)."""


class OfficeCOMError(PPTMCPError):
    """Wraps lower-level Office / file errors with a safe, user-visible message."""
