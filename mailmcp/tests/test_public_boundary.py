from __future__ import annotations

import ast
import re
from pathlib import Path

PACKAGE = Path(__file__).resolve().parents[1] / "src" / "mailmcp"
PUBLIC_ROOT = PACKAGE.parents[1]

EXPECTED_TOOLS = {
    "outlook_health", "outlook_list_accounts", "outlook_list_folders",
    "outlook_list_messages", "outlook_get_message", "outlook_search_messages",
    "outlook_save_attachments", "outlook_search_recipients",
    "outlook_get_mailbox_stats", "outlook_search_all_folders",
    "outlook_compose_mail", "outlook_edit_draft", "outlook_reply",
    "outlook_send_mail", "outlook_message_action", "outlook_bulk_message_action",
    "outlook_get_conversation_thread", "outlook_forward_mail",
    "outlook_reload_config", "outlook_forwarding_rule", "outlook_get_mail_context",
    "outlook_list_calendar_events", "outlook_get_calendar_event",
    "outlook_create_meeting_draft", "outlook_update_calendar_event",
    "outlook_check_freebusy", "outlook_task", "outlook_meeting_request",
    "outlook_category", "outlook_list_contacts", "outlook_get_contact",
    "outlook_search_contacts", "outlook_create_folder",
}

PUBLIC_MODULES = {
    "__init__.py", "__main__.py", "_admin.py", "_calendar.py",
    "_calendar_ops.py", "_categories.py", "_config_policy.py",
    "_contacts_com.py", "_contacts_ops.py", "_context.py", "_context_ops.py",
    "_core.py", "_folders.py", "_formatters.py", "_item_guards.py",
    "_mail_calendar.py", "_mail_compose.py", "_mail_edit.py", "_mail_ops.py",
    "_message_actions.py", "_message_fetch.py", "_message_list.py",
    "_message_ops.py", "_message_recipients.py", "_message_save.py",
    "_message_search.py", "_messages.py", "_msg_rules.py", "_prompts.py",
    "_task_category_ops.py", "_tasks.py", "_tools.py", "draft_routine.py",
    "outlook_com.py", "search_routine.py", "server.py",
}

_ACTION_ID_RE = re.compile(r"\baction-\d{6,}(?:-[0-9a-z-]+)?\b", re.IGNORECASE)
_NON_WEB_URI_RE = re.compile(r"\b(?!https?://)[a-z][a-z0-9+.-]*://", re.IGNORECASE)
_ABSOLUTE_USER_PATH_RE = re.compile(r"[A-Za-z]:\\Users\\[^\r\n\"']+", re.IGNORECASE)


def test_public_package_contains_only_reviewed_modules() -> None:
    existing = {path.name for path in PACKAGE.glob("*.py")}
    assert existing == PUBLIC_MODULES


def test_public_python_tree_has_no_internal_artifact_markers() -> None:
    offenders: list[str] = []
    for path in PUBLIC_ROOT.rglob("*.py"):
        text = path.read_text(encoding="utf-8")
        for label, pattern in (
            ("internal action identifier", _ACTION_ID_RE),
            ("non-web internal URI", _NON_WEB_URI_RE),
            ("absolute user profile path", _ABSOLUTE_USER_PATH_RE),
        ):
            if pattern.search(text):
                offenders.append(f"{path.relative_to(PUBLIC_ROOT)}: {label}")
    assert offenders == []


def test_server_registers_exact_outlook_public_surface() -> None:
    tree = ast.parse((PACKAGE / "server.py").read_text(encoding="utf-8"))
    registered: list[str] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Name) or node.func.id != "_register":
            continue
        assert len(node.args) == 1
        arg = node.args[0]
        assert isinstance(arg, ast.Attribute)
        registered.append(arg.attr)
    assert len(registered) == 33
    assert set(registered) == EXPECTED_TOOLS
    assert len(registered) == len(set(registered))


def test_server_uses_standalone_fastmcp() -> None:
    text = (PACKAGE / "server.py").read_text(encoding="utf-8")
    assert "from fastmcp import FastMCP" in text
    assert "mcp.server.fastmcp" not in text
