"""
Adaptive Card builders for TeamsCommunication.

Each builder returns a dict that can be serialized to JSON and attached
to a Bot Framework Activity or used with the Adaptive Card attachment helper.
"""

from __future__ import annotations

from botbuilder.core import CardFactory
from botbuilder.schema import Attachment


def hello_world_card() -> dict:
    """Return an Adaptive Card payload dict with a Hello World message."""
    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": [
            {
                "type": "TextBlock",
                "text": "Hello world",
                "weight": "Bolder",
                "size": "Large",
                "wrap": True,
            },
            {
                "type": "TextBlock",
                "text": "Posted by the MIDPAgent Teams bot.",
                "isSubtle": True,
                "wrap": True,
            },
        ],
    }


def hello_world_attachment() -> Attachment:
    """Return a Bot Framework Attachment wrapping the Hello World card."""
    return CardFactory.adaptive_card(hello_world_card())


def agent_response_card(user_message: str, agent_response: str) -> dict:
    """Return an Adaptive Card displaying the AI agent's response.

    Parameters
    ----------
    user_message : str
        The original question / prompt from the user.
    agent_response : str
        The text reply from the Azure AI Foundry agent.
    """
    # Truncate very long responses to stay within card size limits
    display_response = agent_response[:4000]
    if len(agent_response) > 4000:
        display_response += "\n\n_(response truncated)_"

    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": [
            {
                "type": "TextBlock",
                "text": "MIDPAgent",
                "weight": "Bolder",
                "size": "Medium",
                "color": "Accent",
                "wrap": True,
            },
            {
                "type": "TextBlock",
                "text": display_response,
                "wrap": True,
            },
        ],
    }


def agent_response_attachment(
    user_message: str, agent_response: str
) -> Attachment:
    """Return a Bot Framework Attachment wrapping an agent response card."""
    return CardFactory.adaptive_card(
        agent_response_card(user_message, agent_response)
    )


def error_card(error_message: str) -> Attachment:
    """Return an Adaptive Card indicating an error occurred."""
    return CardFactory.adaptive_card(
        {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.4",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "⚠ Something went wrong",
                    "weight": "Bolder",
                    "size": "Medium",
                    "color": "Attention",
                    "wrap": True,
                },
                {
                    "type": "TextBlock",
                    "text": error_message,
                    "wrap": True,
                    "isSubtle": True,
                },
            ],
        }
    )
