"""
aiohttp web application exposing the Bot Framework messaging endpoint.

Run:
    python -m TeamsCommunication.app

This starts a web server on the configured port (default 3978) with a
single route: POST /api/messages – the Bot Framework Channel endpoint.
"""

from __future__ import annotations

import logging
import sys
import traceback

from aiohttp import web
from aiohttp.web import Request, Response
from botbuilder.core import (
    BotFrameworkAdapter,
    BotFrameworkAdapterSettings,
    TurnContext,
)
from botbuilder.core.integration import aiohttp_error_middleware
from botbuilder.schema import Activity

from .agent_service import FoundryAgentService
from .bot import MIDPBot
from .config import load_config, validate_bot_identity

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Bootstrap
# ---------------------------------------------------------------------------
CONFIG = load_config()
validate_bot_identity(CONFIG)

SETTINGS = BotFrameworkAdapterSettings(
    app_id=CONFIG.app_id,
    app_password=CONFIG.app_password,
    channel_auth_tenant=CONFIG.tenant_id,
)
ADAPTER = BotFrameworkAdapter(SETTINGS)

# Initialise the Foundry agent service (if configured)
AGENT_SERVICE = None
if CONFIG.ai_project_endpoint and CONFIG.agent_name:
    AGENT_SERVICE = FoundryAgentService(
        endpoint=CONFIG.ai_project_endpoint,
        agent_name=CONFIG.agent_name,
        tenant_id=CONFIG.tenant_id or "",
        client_id=CONFIG.app_id,
        client_secret=CONFIG.app_password,
    )
    logger.info(
        "Foundry agent service initialised (agent=%s)", CONFIG.agent_name
    )
else:
    logger.warning(
        "AZURE_AI_PROJECT_ENDPOINT or AGENT_NAME not set – "
        "bot will reply with hello-world card only."
    )

BOT = MIDPBot(agent_service=AGENT_SERVICE)


# Error handler -----------------------------------------------------------
async def on_error(context: TurnContext, error: Exception) -> None:
    """Global error handler for the adapter."""
    logger.error("[on_turn_error] unhandled error: %s", error)
    traceback.print_exc(file=sys.stderr)
    try:
        await context.send_activity("Sorry, the bot encountered an error.")
    except Exception as send_err:
        # Connection may already be dead – log but don't crash the server.
        logger.error("Could not send error message to user: %s", send_err)


ADAPTER.on_turn_error = on_error


# Routes ------------------------------------------------------------------
async def health(req: Request) -> Response:
    """Health check endpoint for Azure App Service warmup probes."""
    return Response(status=200, text="OK")


async def messages(req: Request) -> Response:
    """Handle incoming Bot Framework activity (POST /api/messages)."""
    if "application/json" not in (req.content_type or ""):
        return Response(status=415)

    body = await req.json()
    activity = Activity().deserialize(body)
    auth_header = req.headers.get("Authorization", "")

    response = await ADAPTER.process_activity(activity, auth_header, BOT.on_turn)
    if response:
        return Response(body=response.body, status=response.status)
    return Response(status=201)


# App factory -------------------------------------------------------------
def create_app() -> web.Application:
    """Build and return the aiohttp Application."""
    app = web.Application(middlewares=[aiohttp_error_middleware])
    app.router.add_get("/", health)
    app.router.add_post("/api/messages", messages)
    return app


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = create_app()
    print(f"Bot listening on http://localhost:{CONFIG.port}/api/messages")
    web.run_app(app, host="0.0.0.0", port=CONFIG.port)
