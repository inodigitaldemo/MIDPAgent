"""
Bot handler for TeamsCommunication.

Subclasses TeamsActivityHandler so the bot can:
- Receive messages in Teams and reply with plain text via the AI Foundry agent
- Handle adaptive card action submissions (yes/no for document production & approval)
- Handle Teams-specific events (member added, etc.)
"""

from __future__ import annotations

import asyncio
import logging
import traceback
from typing import Optional

from botbuilder.core import CardFactory, TurnContext
from botbuilder.core.teams import TeamsActivityHandler
from botbuilder.schema import Activity, ActivityTypes

from .agent_service import FoundryAgentService

logger = logging.getLogger(__name__)

# Max retries when sending a reply back to Teams (connector can drop idle
# connections while we wait for the Foundry agent).
_SEND_RETRIES = 3
_SEND_RETRY_DELAY = 1.0  # seconds


class MIDPBot(TeamsActivityHandler):
    """Teams bot that forwards messages to Azure AI Foundry and replies."""

    def __init__(
        self,
        agent_service: Optional[FoundryAgentService] = None,
        midp_service=None,
    ) -> None:
        super().__init__()
        self._agent_service = agent_service
        self._midp_service = midp_service  # set later by app.py

    # ------------------------------------------------------------------
    # Resilient send helper
    # ------------------------------------------------------------------

    @staticmethod
    async def _send_with_retry(
        turn_context: TurnContext,
        activity: Activity | str,
        retries: int = _SEND_RETRIES,
    ) -> None:
        """Try to send *activity*; retry on connection/timeout errors."""
        for attempt in range(1, retries + 1):
            try:
                await turn_context.send_activity(activity)
                return
            except Exception as exc:
                msg = str(exc).lower()
                is_transient = (
                    "connection" in msg
                    or "timeout" in msg
                    or isinstance(exc, (TimeoutError, asyncio.TimeoutError))
                )
                if attempt < retries and is_transient:
                    logger.warning(
                        "send_activity attempt %d/%d failed (%s) – retrying",
                        attempt, retries, exc,
                    )
                    await asyncio.sleep(_SEND_RETRY_DELAY)
                else:
                    raise

    # ------------------------------------------------------------------
    # Message handler — plain text replies
    # ------------------------------------------------------------------

    async def on_message_activity(self, turn_context: TurnContext) -> None:
        """Handle incoming messages.

        Forwards to the AI Foundry agent and replies with plain text.
        Adaptive cards are reserved for MIDP approval workflows only.
        """
        # Check if this is an adaptive card action submit
        if turn_context.activity.value:
            await self._handle_card_action(turn_context)
            return

        user_text = (turn_context.activity.text or "").strip()

        if not user_text:
            await self._send_with_retry(turn_context, "Vennligst send en tekstmelding.")
            return

        if not self._agent_service:
            await self._send_with_retry(
                turn_context,
                "Hei! Jeg er MIDPAgent-boten. AI-agenten er ikke konfigurert ennå.",
            )
            return

        # Show typing indicator while the agent processes
        try:
            await turn_context.send_activity(Activity(type=ActivityTypes.typing))
        except Exception:
            pass  # Typing indicator is nice-to-have; don't fail on it

        try:
            agent_reply = await self._agent_service.send_message(user_text)
            await self._send_with_retry(turn_context, agent_reply)
        except Exception as exc:
            logger.error("Agent error: %s\n%s", exc, traceback.format_exc())
            try:
                await self._send_with_retry(
                    turn_context,
                    f"Beklager, jeg kunne ikke behandle forespørselen din: {exc}",
                )
            except Exception as send_err:
                logger.error("Failed to send error message: %s", send_err)

    # ------------------------------------------------------------------
    # Replace adaptive card with a static confirmation
    # ------------------------------------------------------------------

    @staticmethod
    async def _disable_card(
        turn_context: TurnContext, status_text: str
    ) -> None:
        """Replace the original adaptive card with a non-interactive card
        showing *status_text* so the buttons can no longer be clicked."""
        updated = Activity(
            id=turn_context.activity.reply_to_id,
            type=ActivityTypes.message,
            attachments=[
                CardFactory.adaptive_card(
                    {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.4",
                        "body": [
                            {
                                "type": "TextBlock",
                                "text": status_text,
                                "wrap": True,
                                "weight": "Bolder",
                            }
                        ],
                    }
                )
            ],
        )
        try:
            await turn_context.update_activity(updated)
        except Exception as exc:
            # Some channels don't support activity updates – log and move on
            logger.warning("Could not update card: %s", exc)

    # ------------------------------------------------------------------
    # Adaptive card action handler
    # ------------------------------------------------------------------

    async def _handle_card_action(self, turn_context: TurnContext) -> None:
        """Process adaptive card submit actions (produce / approve)."""
        value = turn_context.activity.value or {}
        action = value.get("action")

        if action == "produce_document":
            await self._handle_produce(turn_context, value)
        elif action == "approve_document":
            await self._handle_approve(turn_context, value)
        elif action == "reject_document":
            await self._handle_reject(turn_context, value)
        else:
            await self._send_with_retry(
                turn_context, f"Ukjent handling: {action}"
            )

    async def _handle_produce(self, turn_context: TurnContext, value: dict) -> None:
        """User clicked 'Yes' on a produce-document card."""
        item_id = value.get("item_id")
        title = value.get("title", "Unknown")
        choice = value.get("choice", "").lower()

        if choice != "yes":
            await self._disable_card(
                turn_context,
                f"\u274c Produksjon hoppet over for **{title}**.",
            )
            await self._send_with_retry(
                turn_context,
                f"Forstått — hopper over dokumentproduksjon for **{title}**.",
            )
            return

        await self._disable_card(
            turn_context,
            f"\u2699\ufe0f Produserer dokument for **{title}**\u2026",
        )

        if not self._midp_service:
            await self._send_with_retry(
                turn_context, "MIDP-tjenesten er ikke konfigurert."
            )
            return

        await self._send_with_retry(
            turn_context,
            f"Starter dokumentproduksjon for **{title}**… dette kan ta litt tid.",
        )
        try:
            await turn_context.send_activity(Activity(type=ActivityTypes.typing))
        except Exception:
            pass

        # --- Step 1: Produce the document ---
        result = None
        try:
            result = await self._midp_service.produce_document(item_id)
        except Exception as exc:
            logger.error("Document production error: %s", exc, exc_info=True)
            await self._send_with_retry(
                turn_context,
                f"Dokumentproduksjon feilet: {exc}",
            )
            return

        if result.get("error"):
            await self._send_with_retry(
                turn_context,
                f"Dokumentproduksjon feilet: {result['error']}",
            )
            return

        # --- Step 2: Send approval card (separate try so a send failure
        #     doesn't falsely report that production failed) ---
        from .adaptive_cards import document_approval_attachment
        attachment = document_approval_attachment(
            title=title,
            item_id=item_id,
            doc_url=result.get("doc_url", ""),
            filename=result.get("filename", ""),
        )
        reply = Activity(
            type=ActivityTypes.message, attachments=[attachment]
        )
        try:
            await self._send_with_retry(turn_context, reply)
        except Exception as exc:
            logger.warning(
                "Could not send approval card via turn context: %s – "
                "trying proactive fallback", exc,
            )
            try:
                from .proactive import send_to_channel
                await send_to_channel(self._midp_service._config, attachment)
            except Exception as proactive_exc:
                logger.error("Proactive fallback failed: %s", proactive_exc)
            # Even if the card failed to send, production DID succeed.
            try:
                await self._send_with_retry(
                    turn_context,
                    f"✅ Dokumentet for **{title}** ble produsert og lastet opp, "
                    f"men godkjenningskortet kunne ikke vises. "
                    f"Sjekk dokumentbiblioteket.",
                )
            except Exception:
                pass  # connector fully dead — at least the doc is uploaded

    async def _handle_approve(self, turn_context: TurnContext, value: dict) -> None:
        """User approved a document."""
        item_id = value.get("item_id")
        title = value.get("title", "Unknown")
        choice = value.get("choice", "").lower()

        if choice != "yes":
            await self._send_with_retry(
                turn_context,
                f"Dokumentet for **{title}** ble **ikke godkjent**.",
            )
            return

        # Mark as approved in SharePoint if midp_service is available
        if self._midp_service:
            try:
                await self._midp_service.mark_approved(item_id)
            except Exception as exc:
                logger.error("Approval update failed: %s", exc)

        await self._send_with_retry(
            turn_context,
            f"Dokumentet for **{title}** er **godkjent**. ✅",
        )

    async def _handle_reject(self, turn_context: TurnContext, value: dict) -> None:
        """User rejected a document."""
        title = value.get("title", "Unknown")
        await self._disable_card(
            turn_context,
            f"❌ Dokument for **{title}** ble **avvist**.",
        )
        await self._send_with_retry(
            turn_context,
            f"Dokumentet for **{title}** er **avvist**. ❌",
        )

    async def on_members_added_activity(self, members_added, turn_context: TurnContext):
        """Greet new members when added to a conversation."""
        for member in members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity(
                    "Hei! Jeg er MIDPAgent-boten. Send meg en melding, "
                    "så videresender jeg den til AI-agenten for behandling."
                )
