"""
Azure AI Foundry Agent service for TeamsCommunication.

Provides an async wrapper around the Foundry Assistants REST API
(threads / messages / runs) so the Teams bot can forward incoming user
messages to the AI agent and return the generated response.

Auth uses ``ClientSecretCredential`` (works both locally and in
deployed services – no CLI dependency).
"""

from __future__ import annotations

import asyncio
import logging
from typing import Optional

import aiohttp
from azure.identity import ClientSecretCredential

logger = logging.getLogger(__name__)


class FoundryAgentService:
    """Manages a persistent OpenAI-compatible assistant in Azure AI Foundry."""

    API_VERSION = "v1"

    def __init__(
        self,
        endpoint: str,
        agent_name: str,
        tenant_id: str,
        client_id: str,
        client_secret: str,
    ) -> None:
        self._endpoint = endpoint.rstrip("/")
        self._agent_name = agent_name
        self._tenant_id = tenant_id
        self._client_id = client_id
        self._client_secret = client_secret

        self._credential = ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
        )

        # Cached assistant id (resolved lazily)
        self._assistant_id: Optional[str] = None

    # ------------------------------------------------------------------
    # Token helpers
    # ------------------------------------------------------------------

    async def _get_token(self) -> str:
        """Acquire an Azure AI Foundry token using client credentials."""
        token = await asyncio.to_thread(
            self._credential.get_token, "https://ai.azure.com/.default"
        )
        return token.token

    def _headers(self, token: str) -> dict:
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    # ------------------------------------------------------------------
    # Assistant management
    # ------------------------------------------------------------------

    async def _ensure_assistant(self, token: str) -> str:
        """Find or create the OpenAI-compatible assistant in Foundry."""
        if self._assistant_id:
            return self._assistant_id

        headers = self._headers(token)
        base = self._endpoint

        async with aiohttp.ClientSession() as session:
            # Check for an existing assistant with this name
            async with session.get(
                f"{base}/assistants?api-version={self.API_VERSION}",
                headers=headers,
            ) as resp:
                resp.raise_for_status()
                data = await resp.json()
                for asst in data.get("data", []):
                    if asst.get("name") == self._agent_name:
                        self._assistant_id = asst["id"]
                        logger.info("Found existing assistant %s", self._assistant_id)
                        return self._assistant_id

            # Fetch agent definition so we can create the assistant
            async with session.get(
                f"{base}/agents/{self._agent_name}?api-version={self.API_VERSION}",
                headers=headers,
            ) as resp:
                resp.raise_for_status()
                agent_def = (await resp.json())["versions"]["latest"]["definition"]

            # Create the assistant
            async with session.post(
                f"{base}/assistants?api-version={self.API_VERSION}",
                headers=headers,
                json={
                    "model": agent_def["model"],
                    "name": self._agent_name,
                    "instructions": agent_def.get("instructions", ""),
                },
            ) as resp:
                resp.raise_for_status()
                self._assistant_id = (await resp.json())["id"]
                logger.info("Created assistant %s", self._assistant_id)
                return self._assistant_id

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    async def send_message(self, user_text: str) -> str:
        """Send *user_text* to the Foundry agent and return its reply.

        Creates a new thread for each conversation turn (stateless).
        The full flow: token → ensure assistant → create thread →
        post user message → start run → poll → retrieve assistant reply.
        """
        token = await self._get_token()
        assistant_id = await self._ensure_assistant(token)
        headers = self._headers(token)
        base = self._endpoint
        v = self.API_VERSION

        async with aiohttp.ClientSession() as session:
            # 1. Create thread
            async with session.post(
                f"{base}/threads?api-version={v}",
                headers=headers,
                json={},
            ) as resp:
                resp.raise_for_status()
                thread_id = (await resp.json())["id"]

            # 2. Post user message
            async with session.post(
                f"{base}/threads/{thread_id}/messages?api-version={v}",
                headers=headers,
                json={"role": "user", "content": user_text},
            ) as resp:
                resp.raise_for_status()

            # 3. Start run
            async with session.post(
                f"{base}/threads/{thread_id}/runs?api-version={v}",
                headers=headers,
                json={"assistant_id": assistant_id},
            ) as resp:
                resp.raise_for_status()
                run_id = (await resp.json())["id"]

            # 4. Poll until terminal state
            status = "queued"
            while status not in ("completed", "failed", "cancelled", "expired"):
                await asyncio.sleep(1.5)
                async with session.get(
                    f"{base}/threads/{thread_id}/runs/{run_id}?api-version={v}",
                    headers=headers,
                ) as resp:
                    resp.raise_for_status()
                    status = (await resp.json())["status"]

            if status != "completed":
                logger.warning("Run %s ended with status '%s'", run_id, status)
                return f"[Agent run ended with status '{status}']"

            # 5. Retrieve assistant reply
            async with session.get(
                f"{base}/threads/{thread_id}/messages?api-version={v}",
                headers=headers,
            ) as resp:
                resp.raise_for_status()
                for msg in (await resp.json()).get("data", []):
                    if msg["role"] == "assistant":
                        for block in msg.get("content", []):
                            if block.get("type") == "text":
                                return block["text"]["value"]

        return "[No response from agent]"
