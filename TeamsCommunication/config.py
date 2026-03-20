"""
Configuration for the TeamsCommunication Bot Framework module.

Reads from the `env` file (dotenv) with the following keys:

Bot identity (from Azure Bot Service resource):
- BOT_APP_ID          – Microsoft App ID of the bot
- BOT_APP_PASSWORD    – Client secret for the bot's app registration

Teams targeting (optional – used by proactive messaging):
- TEAMS_TEAM_ID       – Override: Graph group/team ID
- TEAMS_CHANNEL_ID    – Override: Teams channel ID
- TEAMS_SERVICE_URL   – Bot Framework service URL (defaults to https://smba.trafficmanager.net/emea/)

Existing SharePoint env vars are also loaded for team/channel lookup:
- SHAREPOINT_SITE_URL
- AZURE_TENANT_ID
- SHAREPOINT_CLIENT_ID
- SHAREPOINT_CLIENT_SECRET
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional

from dotenv import load_dotenv


@dataclass
class BotConfig:
    """Immutable configuration object for the bot."""

    # Bot identity
    app_id: str = ""
    app_password: str = ""

    # Teams targeting
    team_id: Optional[str] = None
    channel_id: Optional[str] = None
    service_url: str = "https://smba.trafficmanager.net/emea/"

    # Existing SharePoint env (for Graph-based team/channel resolution)
    tenant_id: Optional[str] = None
    sharepoint_site_url: Optional[str] = None
    sharepoint_client_id: Optional[str] = None
    sharepoint_client_secret: Optional[str] = None

    # Azure AI Foundry agent
    ai_project_endpoint: Optional[str] = None
    agent_name: Optional[str] = None

    # Web server
    port: int = 3978


def load_config() -> BotConfig:
    """Load configuration from environment / dotenv file."""
    load_dotenv("env")
    load_dotenv()

    return BotConfig(
        app_id=os.getenv("BOT_APP_ID", ""),
        app_password=os.getenv("BOT_APP_PASSWORD", ""),
        team_id=os.getenv("TEAMS_TEAM_ID"),
        channel_id=os.getenv("TEAMS_CHANNEL_ID"),
        service_url=os.getenv(
            "TEAMS_SERVICE_URL", "https://smba.trafficmanager.net/emea/"
        ),
        tenant_id=os.getenv("AZURE_TENANT_ID"),
        sharepoint_site_url=os.getenv("SHAREPOINT_SITE_URL"),
        sharepoint_client_id=os.getenv("SHAREPOINT_CLIENT_ID"),
        sharepoint_client_secret=os.getenv("SHAREPOINT_CLIENT_SECRET"),
        ai_project_endpoint=os.getenv("AZURE_AI_PROJECT_ENDPOINT"),
        agent_name=os.getenv("AGENT_NAME"),
        port=int(os.getenv("BOT_PORT", "3978")),
    )


def validate_bot_identity(config: BotConfig) -> None:
    """Raise if bot credentials are missing."""
    missing = []
    if not config.app_id:
        missing.append("BOT_APP_ID")
    if not config.app_password:
        missing.append("BOT_APP_PASSWORD")
    if missing:
        raise EnvironmentError(
            "Missing required bot environment variables: "
            + ", ".join(sorted(missing))
        )
