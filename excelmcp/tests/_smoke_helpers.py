"""Shared MCP dispatch helpers for test_smoke_* files.

Import with:  from ._smoke_helpers import _get_tools_dict, _run_tool

FastMCP 3.x compat: use mcp.list_tools() (async, returns Sequence[Tool])
and mcp.call_tool(name, arguments) to invoke tools.
"""
from __future__ import annotations

import asyncio
import json


class _ToolProxy:
    """Thin wrapper to preserve the tools[name].run(args) calling convention."""

    def __init__(self, mcp, name: str):
        self._mcp = mcp
        self._name = name

    async def run(self, arguments: dict):
        return await self._mcp.call_tool(self._name, arguments)


def _run_async(awaitable):
    """Run a FastMCP async API from synchronous tests."""
    return asyncio.run(awaitable)


def _get_tools_dict(mcp):
    """Return a dict of tool_name -> _ToolProxy.

    FastMCP 3.x: mcp.list_tools() is async and returns Sequence[Tool].
    Each proxy's async run(arguments) delegates to mcp.call_tool().
    """
    async def _load():
        tools = await mcp.list_tools()
        return {t.name: _ToolProxy(mcp, t.name) for t in tools}

    return _run_async(_load())


def _get_tool_names(mcp) -> set[str]:
    """Return the registered MCP tool names via the public FastMCP API."""
    async def _load():
        tools = await mcp.list_tools()
        return {tool.name for tool in tools}

    return _run_async(_load())


def _get_prompt_names(mcp) -> set[str]:
    """Return the registered MCP prompt names via the public FastMCP API."""
    async def _load():
        prompts = await mcp.list_prompts()
        return {prompt.name for prompt in prompts}

    return _run_async(_load())


def _get_capability_prompt_names(mcp) -> set[str]:
    """Return the full live prompt registry via the public FastMCP API."""
    return _get_prompt_names(mcp)


def _render_prompt(mcp, prompt_name: str, arguments: dict | None = None) -> str:
    """Render a prompt via the public FastMCP prompt API and normalize text output."""
    async def _load():
        prompt = await mcp.get_prompt(prompt_name)
        if prompt is None:
            raise AssertionError(f"Prompt not found: {prompt_name}")
        return await prompt.render(arguments or {})

    rendered = _run_async(_load())

    if hasattr(rendered, "messages") and rendered.messages:
        parts: list[str] = []
        for message in rendered.messages:
            content = getattr(message, "content", None)
            text = getattr(content, "text", None)
            if isinstance(text, str):
                parts.append(text)
        return "\n".join(parts)

    return str(rendered)


def _run_tool(mcp, tool_name: str, arguments: dict):
    """Call a tool via mcp.call_tool() and return the result dict.

    FastMCP 3.x: mcp.call_tool(name, arguments) returns ToolResult which
    carries structured_content (decoded dict) and content (TextContent list).
    """
    async def _call():
        return await mcp.call_tool(tool_name, arguments)

    raw = _run_async(_call())

    # FastMCP 3.x: ToolResult.structured_content is the decoded dict
    if hasattr(raw, "structured_content") and raw.structured_content is not None:
        return raw.structured_content

    # Fallback: parse JSON from TextContent items (FastMCP 2.x path)
    if hasattr(raw, "content") and raw.content:
        for item in raw.content:
            if hasattr(item, "text"):
                try:
                    return json.loads(item.text)
                except (ValueError, TypeError):
                    pass

    # FastMCP 1.x: dict returned directly
    if isinstance(raw, dict):
        return raw

    return raw


def _get_capabilities_snapshot(mcp) -> dict[str, set[str]]:
    """Return normalized tool and prompt inventories from the capabilities tool."""
    snapshot = _run_tool(mcp, "capabilities", {})
    prompts = snapshot.get("prompts", {}) if isinstance(snapshot, dict) else {}
    return {
        "tools": set(snapshot.get("tools", [])),
        "prompts": set(prompts.get("names", [])),
        "prompt_count": prompts.get("count", 0),
    }

