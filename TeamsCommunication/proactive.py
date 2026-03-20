"""
Proactive messaging helper for TeamsCommunication.

Uses the Bot Framework Connector SDK to send a message (e.g. an Adaptive
Card) to a Teams channel without waiting for an incoming activity.

Typical usage:
    from TeamsCommunication.proactive import send_to_channel
    from TeamsCommunication.adaptive_cards import hello_world_attachment
    from TeamsCommunication.config import load_config

    import asyncio
    config = load_config()
    asyncio.run(send_to_channel(config, hello_world_attachment()))
"""

from __future__ import annotations

import json
import sys
from typing import Optional

import aiohttp

from botbuilder.schema import Attachment

from .config import BotConfig, validate_bot_identity


async def _get_bot_token(app_id: str, app_password: str, tenant_id: str) -> str:
    """Acquire a Bot Framework access token via OAuth 2.0 client credentials.

    Tries multiple endpoint/scope combinations to find one the Connector accepts.
    """
    attempts = [
        # v1.0 endpoint (produces token with 'appid' claim)
        {
            "url": f"https://login.microsoftonline.com/{tenant_id}/oauth2/token",
            "data": {
                "grant_type": "client_credentials",
                "client_id": app_id,
                "client_secret": app_password,
                "resource": "https://api.botframework.com",
            },
            "label": "v1.0 / api.botframework.com",
        },
        # v2.0 endpoint with standard BF scope
        {
            "url": f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
            "data": {
                "grant_type": "client_credentials",
                "client_id": app_id,
                "client_secret": app_password,
                "scope": "https://api.botframework.com/.default",
            },
            "label": "v2.0 / api.botframework.com/.default",
        },
    ]
    last_error = ""
    for attempt in attempts:
        async with aiohttp.ClientSession() as session:
            async with session.post(attempt["url"], data=attempt["data"]) as resp:
                body = await resp.json()
                if resp.status == 200:
                    print(f"  Token acquired via {attempt['label']}")
                    return body["access_token"]
                last_error = body.get("error_description", str(body))
                print(
                    f"  {attempt['label']} failed: {last_error}",
                    file=sys.stderr,
                )
    raise PermissionError(f"Could not acquire bot token. Last error: {last_error}")


async def send_to_channel(
    config: BotConfig,
    attachment: Attachment,
    *,
    team_id: Optional[str] = None,
    channel_id: Optional[str] = None,
    service_url: Optional[str] = None,
) -> str:
    """
    Proactively post an Attachment (e.g. Adaptive Card) to a Teams channel.

    Parameters
    ----------
    config : BotConfig
        Bot configuration (must include app_id + app_password).
    attachment : Attachment
        The attachment to send (use ``hello_world_attachment()``).
    team_id : str, optional
        Override for the Teams team (group) ID.  Falls back to
        ``config.team_id``.
    channel_id : str, optional
        Override for the channel ID.  Falls back to ``config.channel_id``.
    service_url : str, optional
        Override for the Bot Framework service URL.

    Returns
    -------
    str
        The ID of the newly created activity (message).
    """
    validate_bot_identity(config)

    effective_channel_id = channel_id or config.channel_id
    effective_service_url = (service_url or config.service_url).rstrip("/")

    if not effective_channel_id:
        raise ValueError(
            "No channel ID provided.  Set TEAMS_CHANNEL_ID in env or "
            "pass channel_id explicitly."
        )

    # 1. Get Bot Framework token using the tenant-scoped endpoint
    token = await _get_bot_token(
        config.app_id, config.app_password, config.tenant_id or ""
    )
    print("  Bot token acquired.")

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    # 2. Create a new conversation in the channel (proactive)
    create_url = f"{effective_service_url}/v3/conversations"
    create_payload = {
        "isGroup": True,
        "channelData": {
            "channel": {"id": effective_channel_id},
        },
        "bot": {
            "id": f"28:{config.app_id}",
            "name": "MIDPAgent",
        },
        "tenantId": config.tenant_id,
        "activity": {
            "type": "message",
            "attachments": [
                {
                    "contentType": attachment.content_type,
                    "content": attachment.content,
                }
            ],
        },
    }

    async with aiohttp.ClientSession() as session:
        async with session.post(
            create_url, headers=headers, json=create_payload
        ) as resp:
            body_text = await resp.text()
            if resp.status >= 400:
                print(
                    f"  Connector error {resp.status}: {body_text}",
                    file=sys.stderr,
                )
                resp.raise_for_status()
            result = json.loads(body_text)
            return result.get("activityId", result.get("id", ""))

    return result_id
