from __future__ import annotations

import ast
from pathlib import Path

PACKAGE = Path(__file__).resolve().parents[1] / "src" / "mailmcp"

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

EXCLUDED_MODULES = {
    "_executor.py", "_mailrepo_disposition.py", "_mailrepo_fts.py",
    "_mailrepo_hybrid.py", "_mailrepo_imports.py", "_mailrepo_ingest.py",
    "_mailrepo_models.py", "_mailrepo_multi.py", "_mailrepo_ops.py",
    "_mailrepo_search.py", "_prewarm.py", "_retrieval_router.py",
    "mailrepo_db.py", "mailrepo_tools.py",
}


def test_private_modules_are_not_in_public_package() -> None:
    existing = {path.name for path in PACKAGE.glob("*.py")}
    assert not (existing & EXCLUDED_MODULES)


def test_public_source_has_no_private_module_imports_or_internal_ids() -> None:
    forbidden = (
        "mailmcp._mailrepo", "mailmcp.mailrepo", "mailmcp._retrieval_router",
        "mailmcp._prewarm", "action-177", "action-178", "cortex://",
        "C:\\Users\\", "msoffice-mcps",
    )
    offenders: list[str] = []
    for path in PACKAGE.glob("*.py"):
        text = path.read_text(encoding="utf-8")
        for token in forbidden:
            if token in text:
                offenders.append(f"{path.name}: {token}")
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
