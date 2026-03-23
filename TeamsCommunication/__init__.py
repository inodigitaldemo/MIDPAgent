"""TeamsCommunication – Bot Framework SDK module for Teams messaging."""

from .adaptive_cards import (
    document_approval_attachment,
    document_approval_card,
    error_card,
    produce_document_attachment,
    produce_document_card,
)
from .agent_service import FoundryAgentService
from .bot import MIDPBot
from .config import BotConfig, load_config, validate_bot_identity
from .midp_service import MIDPService
from .proactive import send_to_channel

__all__ = [
    "BotConfig",
    "FoundryAgentService",
    "MIDPBot",
    "MIDPService",
    "document_approval_attachment",
    "document_approval_card",
    "error_card",
    "load_config",
    "produce_document_attachment",
    "produce_document_card",
    "send_to_channel",
    "validate_bot_identity",
]
