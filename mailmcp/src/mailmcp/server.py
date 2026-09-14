"""Public Outlook MailMCP FastMCP stdio server.

This public server intentionally exposes only local Outlook/Windows automation.
MailRepo, embedding, retrieval-router and internal governance integrations are
not part of the public package.
"""
from __future__ import annotations

import logging
import sys

logging.basicConfig(
    stream=sys.stderr,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s %(message)s",
)
logger = logging.getLogger("mailmcp.server")

from fastmcp import FastMCP  # noqa: E402
from mailmcp import _message_ops as _msg  # noqa: E402
from mailmcp import _calendar_ops as _cal  # noqa: E402
from mailmcp import _task_category_ops as _tc  # noqa: E402
from mailmcp import _contacts_ops as _contacts  # noqa: E402
from mailmcp import _context_ops as _ctx_ops  # noqa: E402
from mailmcp._tools import _safe, make_register  # noqa: E402
from mailmcp._prompts import register_prompts  # noqa: E402

__all__ = ["mcp", "_safe"]

mcp = FastMCP("mailmcp")
_register = make_register(mcp)

# Message / folder / mail tools (21)
_register(_msg.outlook_health)
_register(_msg.outlook_list_accounts)
_register(_msg.outlook_list_folders)
_register(_msg.outlook_list_messages)
_register(_msg.outlook_get_message)
_register(_msg.outlook_search_messages)
_register(_msg.outlook_save_attachments)
_register(_msg.outlook_search_recipients)
_register(_msg.outlook_get_mailbox_stats)
_register(_msg.outlook_search_all_folders)
_register(_msg.outlook_compose_mail)
_register(_msg.outlook_edit_draft)
_register(_msg.outlook_reply)
_register(_msg.outlook_send_mail)
_register(_msg.outlook_message_action)
_register(_msg.outlook_bulk_message_action)
_register(_msg.outlook_get_conversation_thread)
_register(_msg.outlook_forward_mail)
_register(_msg.outlook_reload_config)
_register(_msg.outlook_forwarding_rule)
_register(_ctx_ops.outlook_get_mail_context)

# Calendar / meeting / task / category tools (8)
_register(_cal.outlook_list_calendar_events)
_register(_cal.outlook_get_calendar_event)
_register(_cal.outlook_create_meeting_draft)
_register(_cal.outlook_update_calendar_event)
_register(_cal.outlook_check_freebusy)
_register(_tc.outlook_task)
_register(_msg.outlook_meeting_request)
_register(_tc.outlook_category)

# Contacts / folder tools (4)
_register(_contacts.outlook_list_contacts)
_register(_contacts.outlook_get_contact)
_register(_contacts.outlook_search_contacts)
_register(_contacts.outlook_create_folder)

register_prompts(mcp)


def main() -> None:
    logger.info("mailmcp starting (stdio transport; public Outlook-only surface)")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
