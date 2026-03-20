"""TeamsCommunication – Bot Framework SDK module for Teams messaging."""

from .adaptive_cards import (
    agent_response_attachment,
    agent_response_card,
    error_card,
    hello_world_attachment,
    hello_world_card,
)
from .agent_service import FoundryAgentService
from .bot import MIDPBot
from .config import BotConfig, load_config, validate_bot_identity
from .proactive import send_to_channel

__all__ = [
    "BotConfig",
    "FoundryAgentService",
    "MIDPBot",
    "agent_response_attachment",
    "agent_response_card",
    "error_card",
    "hello_world_attachment",
    "hello_world_card",
    "load_config",
    "send_to_channel",
    "validate_bot_identity",
]
