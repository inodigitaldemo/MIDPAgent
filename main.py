"""
MIDPAgent – Master Information Delivery Plan Agent

Connects to a SharePoint list, retrieves items created today, prints them
to the terminal, and forwards each item's Title to the Azure AI Foundry agent
named "inodidigtal-documentmanager".
"""

import os
import sys
from datetime import date

from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.client_credential import ClientCredential
from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential
from openai import OpenAI

load_dotenv()

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SHAREPOINT_SITE_URL = os.environ.get("SHAREPOINT_SITE_URL", "")
SHAREPOINT_LIST_NAME = os.environ.get("SHAREPOINT_LIST_NAME", "Master Information Delivery Plan")

SHAREPOINT_USERNAME = os.environ.get("SHAREPOINT_USERNAME")
SHAREPOINT_PASSWORD = os.environ.get("SHAREPOINT_PASSWORD")
SHAREPOINT_CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET")

AZURE_AI_PROJECT_ENDPOINT = os.environ.get("AZURE_AI_PROJECT_ENDPOINT", "")

AGENT_NAME = "inodidigtal-documentmanager"


# ---------------------------------------------------------------------------
# SharePoint helpers
# ---------------------------------------------------------------------------

def build_sharepoint_context() -> ClientContext:
    """Build an authenticated SharePoint ClientContext.

    Supports two authentication methods depending on which environment
    variables are set:
    - App-only (client credentials): SHAREPOINT_CLIENT_ID + SHAREPOINT_CLIENT_SECRET
    - Delegated (username/password): SHAREPOINT_USERNAME + SHAREPOINT_PASSWORD
    """
    if not SHAREPOINT_SITE_URL:
        raise EnvironmentError("SHAREPOINT_SITE_URL environment variable is required.")

    if SHAREPOINT_CLIENT_ID and SHAREPOINT_CLIENT_SECRET:
        credential = ClientCredential(SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET)
        return ClientContext(SHAREPOINT_SITE_URL).with_credentials(credential)

    if SHAREPOINT_USERNAME and SHAREPOINT_PASSWORD:
        credential = UserCredential(SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD)
        return ClientContext(SHAREPOINT_SITE_URL).with_credentials(credential)

    raise EnvironmentError(
        "SharePoint credentials are missing. "
        "Set either SHAREPOINT_CLIENT_ID + SHAREPOINT_CLIENT_SECRET "
        "or SHAREPOINT_USERNAME + SHAREPOINT_PASSWORD."
    )


def get_items_created_today(ctx: ClientContext, list_name: str) -> list:
    """Return all items in *list_name* whose Created date is today (UTC)."""
    today = date.today().isoformat()  # e.g. "2026-03-20"

    odata_filter = (
        f"Created ge datetime'{today}T00:00:00Z' "
        f"and Created le datetime'{today}T23:59:59Z'"
    )

    sp_list = ctx.web.lists.get_by_title(list_name)
    items = (
        sp_list.items
        .filter(odata_filter)
        .get()
        .execute_query()
    )
    return list(items)


# ---------------------------------------------------------------------------
# Azure AI helpers
# ---------------------------------------------------------------------------

def build_openai_client() -> OpenAI:
    """Build an authenticated OpenAI client via AIProjectClient.

    The AIProjectClient.get_openai_client() method returns a standard OpenAI
    client pre-configured with the project endpoint and Entra ID credentials.
    Conversational agents (Assistants) are accessed through this client via
    the openai_client.beta.assistants / beta.threads API.
    """
    if not AZURE_AI_PROJECT_ENDPOINT:
        raise EnvironmentError("AZURE_AI_PROJECT_ENDPOINT environment variable is required.")

    project_client = AIProjectClient(
        endpoint=AZURE_AI_PROJECT_ENDPOINT,
        credential=DefaultAzureCredential(),
    )
    return project_client.get_openai_client()


def find_assistant(openai_client: OpenAI, agent_name: str):
    """Return the first assistant whose name matches *agent_name*, or None."""
    for assistant in openai_client.beta.assistants.list():
        if assistant.name == agent_name:
            return assistant
    return None


def send_title_to_agent(openai_client: OpenAI, assistant, title: str) -> str:
    """Create a thread, post *title* as a user message, run the assistant, and
    return the assistant's last text reply."""
    thread = openai_client.beta.threads.create()
    openai_client.beta.threads.messages.create(
        thread_id=thread.id,
        role="user",
        content=title,
    )
    run = openai_client.beta.threads.runs.create_and_poll(
        thread_id=thread.id,
        assistant_id=assistant.id,
    )

    if run.status != "completed":
        return f"[Run ended with status '{run.status}']"

    messages = openai_client.beta.threads.messages.list(thread_id=thread.id)
    for message in messages:
        if message.role == "assistant":
            for content_block in message.content:
                if hasattr(content_block, "text"):
                    return content_block.text.value
    return "[No response from agent]"


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    # ── 1. Connect to SharePoint ────────────────────────────────────────────
    print("Connecting to SharePoint…")
    try:
        ctx = build_sharepoint_context()
        # Verify connectivity by loading the web title
        ctx.web.get().execute_query()
        print(f"Connected to: {ctx.web.url}\n")
    except Exception as exc:
        print(f"ERROR: Could not connect to SharePoint – {exc}", file=sys.stderr)
        sys.exit(1)

    # ── 2. Fetch items created today ────────────────────────────────────────
    print(f"Fetching items from list '{SHAREPOINT_LIST_NAME}' created today ({date.today()})…\n")
    try:
        items = get_items_created_today(ctx, SHAREPOINT_LIST_NAME)
    except Exception as exc:
        print(f"ERROR: Could not read SharePoint list – {exc}", file=sys.stderr)
        sys.exit(1)

    if not items:
        print("No new items found today. Exiting.")
        return

    # ── 3. Print all items to the terminal ─────────────────────────────────
    print(f"Found {len(items)} item(s) created today:\n")
    for item in items:
        props = item.properties
        print(
            f"  ID: {props.get('ID', props.get('Id', '?')):>6}  "
            f"Title: {props.get('Title', '(no title)')}"
        )
    print()

    # ── 4. Connect to Azure AI Projects ────────────────────────────────────
    print("Connecting to Azure AI Projects…")
    try:
        openai_client = build_openai_client()
    except Exception as exc:
        print(f"ERROR: Could not create AI client – {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Looking for agent '{AGENT_NAME}'…")
    try:
        assistant = find_assistant(openai_client, AGENT_NAME)
    except Exception as exc:
        print(f"ERROR: Could not list assistants – {exc}", file=sys.stderr)
        sys.exit(1)

    if assistant is None:
        print(
            f"ERROR: Agent '{AGENT_NAME}' was not found in the project. "
            "Please create it in Azure AI Foundry first.",
            file=sys.stderr,
        )
        sys.exit(1)

    print(f"Agent found (id={assistant.id}). Sending titles…\n")

    # ── 5. Send each title to the agent ────────────────────────────────────
    for item in items:
        title = item.properties.get("Title", "").strip()
        if not title:
            continue

        print(f"→ Sending: '{title}'")
        try:
            response = send_title_to_agent(openai_client, assistant, title)
            print(f"  Agent response: {response}\n")
        except Exception as exc:
            print(f"  ERROR processing '{title}': {exc}\n", file=sys.stderr)


if __name__ == "__main__":
    main()
