"""MCP prompt definitions for the public MailMCP server."""
from __future__ import annotations


def register_prompts(mcp) -> None:
    @mcp.prompt()
    def draft_reply(message_subject: str, message_from: str, message_body: str, instructions: str = "Draft a concise, professional reply.") -> str:
        return (
            f"You are a professional email assistant.\n"
            f"The user received an email from {message_from!r} with subject {message_subject!r}.\n"
            f"Original message:\n{message_body}\n\n"
            f"Task: {instructions}\n"
            f"Use the outlook_compose_mail tool to save the draft reply. Do NOT send it."
        )

    @mcp.prompt()
    def triage_inbox(messages_json: str, focus: str = "prioritise by urgency and action required") -> str:
        return (
            f"You are an inbox triage assistant.\nMessages (JSON array from outlook_list_messages):\n{messages_json}\n\n"
            f"Goal: {focus}.\nFor each message, output: urgency (high/medium/low), required action, and whether an immediate reply is needed. Use a structured bullet list."
        )

    @mcp.prompt()
    def schedule_meeting(attendees: str, purpose: str, proposed_slots: str, duration_minutes: int = 30) -> str:
        return (
            f"You are a scheduling assistant.\nMeeting purpose: {purpose}\nAttendees: {attendees}\n"
            f"Proposed time slots: {proposed_slots}\nDuration: {duration_minutes} minutes\n\n"
            "Steps:\n1. Call outlook_check_freebusy for all attendees across the proposed slots.\n"
            "2. Pick the earliest slot where all required attendees are free.\n"
            "3. Call outlook_create_meeting_draft to save the invite.\n"
            "Do NOT send the invite — leave it as a draft for the user to review."
        )

    @mcp.prompt()
    def compose_brief(task_description: str, recipient: str, tone: str = "formal") -> str:
        return (
            f"You are an email writing assistant.\nTask: {task_description}\nRecipient: {recipient}\nTone: {tone}\n\n"
            "Write a professional email body of no more than 200 words. Then call outlook_compose_mail to save it as a draft. Do NOT send it."
        )
