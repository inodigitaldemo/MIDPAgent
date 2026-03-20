"""
CLI entry point – proactively post a Hello World Adaptive Card to a
Teams channel using the Bot Framework SDK.

Prerequisites:
- BOT_APP_ID and BOT_APP_PASSWORD set in ``env``
- TEAMS_CHANNEL_ID set in ``env``
- The bot must be installed / added to the target Team

Run:
    python -m TeamsCommunication.console_app
"""

from __future__ import annotations

import asyncio
import sys

from .adaptive_cards import hello_world_attachment
from .config import load_config, validate_bot_identity
from .proactive import send_to_channel


async def _main() -> None:
    config = load_config()
    validate_bot_identity(config)

    if not config.channel_id:
        print(
            "ERROR: TEAMS_CHANNEL_ID is required for proactive messaging.\n"
            "Set it in the env file.",
            file=sys.stderr,
        )
        sys.exit(1)

    print(f"Bot App ID  : {config.app_id}")
    print(f"Channel ID  : {config.channel_id}")
    print(f"Service URL : {config.service_url}")

    attachment = hello_world_attachment()
    activity_id = await send_to_channel(config, attachment)

    print(f"\nAdaptive Card posted successfully.  Activity ID: {activity_id}")


def main() -> None:
    """Synchronous wrapper so the module works with ``python -m``."""
    asyncio.run(_main())


if __name__ == "__main__":
    main()
