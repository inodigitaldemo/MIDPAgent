"""
Bot handler for TeamsCommunication.

Subclasses TeamsActivityHandler so the bot can:
- Receive messages in Teams and forward them to the Azure AI Foundry agent
- Reply with the agent's response as an Adaptive Card
- Handle Teams-specific events (member added, etc.)
"""

from __future__ import annotations

import asyncio
import logging
import traceback
from typing import Optional

from botbuilder.core import TurnContext
from botbuilder.core.teams import TeamsActivityHandler
from botbuilder.schema import Activity, ActivityTypes

from .adaptive_cards import (
    agent_response_attachment,
    error_card,
    hello_world_attachment,
)
from .agent_service import FoundryAgentService

logger = logging.getLogger(__name__)

# Max retries when sending a reply back to Teams (connector can drop idle
# connections while we wait for the Foundry agent).
_SEND_RETRIES = 3
_SEND_RETRY_DELAY = 1.0  # seconds


class MIDPBot(TeamsActivityHandler):
    """Teams bot that forwards messages to Azure AI Foundry and replies."""

    def __init__(self, agent_service: Optional[FoundryAgentService] = None) -> None:
        super().__init__()
        self._agent_service = agent_service

    # ------------------------------------------------------------------
    # Resilient send helper
    # ------------------------------------------------------------------

    @staticmethod
    async def _send_with_retry(
        turn_context: TurnContext,
        activity: Activity | str,
        retries: int = _SEND_RETRIES,
    ) -> None:
        """Try to send *activity*; retry on connection errors."""
        for attempt in range(1, retries + 1):
            try:
                await turn_context.send_activity(activity)
                return
            except Exception as exc:
                is_conn = "connection" in str(exc).lower()
                if attempt < retries and is_conn:
                    logger.warning(
                        "send_activity attempt %d/%d failed (%s) – retrying",
                        attempt, retries, exc,
                    )
                    await asyncio.sleep(_SEND_RETRY_DELAY)
                else:
                    raise

    # ------------------------------------------------------------------
    # Message handler
    # ------------------------------------------------------------------

    async def on_message_activity(self, turn_context: TurnContext) -> None:
        """Handle incoming messages.

        If an agent service is configured, forward the user's text to the
        Azure AI Foundry assistant and reply with its answer.  Otherwise
        fall back to the hello-world card.
        """
        user_text = (turn_context.activity.text or "").strip()

        if not user_text:
            await self._send_with_retry(turn_context, "Please send a text message.")
            return

        if not self._agent_service:
            # No agent configured – fall back to hello world
            attachment = hello_world_attachment()
            reply = Activity(type=ActivityTypes.message, attachments=[attachment])
            await self._send_with_retry(turn_context, reply)
            return

        # Show typing indicator while the agent processes
        try:
            await turn_context.send_activity(Activity(type=ActivityTypes.typing))
        except Exception:
            pass  # Typing indicator is nice-to-have; don't fail on it

        try:
            agent_reply = await self._agent_service.send_message(user_text)
            attachment = agent_response_attachment(user_text, agent_reply)
            reply = Activity(type=ActivityTypes.message, attachments=[attachment])
            await self._send_with_retry(turn_context, reply)
        except Exception as exc:
            logger.error("Agent error: %s\n%s", exc, traceback.format_exc())
            try:
                await self._send_with_retry(
                    turn_context,
                    Activity(
                        type=ActivityTypes.message,
                        attachments=[
                            error_card(
                                f"The AI agent could not process your request: {exc}"
                            )
                        ],
                    ),
                )
            except Exception as send_err:
                logger.error("Failed to send error card: %s", send_err)

    async def on_members_added_activity(self, members_added, turn_context: TurnContext):
        """Greet new members when added to a conversation."""
        for member in members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity(
                    "Hello! I'm the MIDPAgent bot. Send me a message "
                    "and I'll forward it to the AI agent for processing."
                )
