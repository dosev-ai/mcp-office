"""Shared FastMCP instance for excelmcp sub-modules.

server.py and each server_*.py sub-module import `mcp` from here so all
@mcp.tool() registrations go to the same FastMCP instance. This prevents
double-registration when server.py imports the sub-modules as side-effects.
"""
from __future__ import annotations

import os

# Suppress FastMCP startup banner before instance creation.
os.environ.setdefault("FASTMCP_SHOW_CLI_BANNER", "false")

from fastmcp import FastMCP

mcp = FastMCP("excelmcp")
