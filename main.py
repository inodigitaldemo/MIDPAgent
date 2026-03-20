"""
MIDPAgent – Master Information Delivery Plan Agent

Connects to a SharePoint list via Microsoft Graph API, retrieves items
created today, prints them to the terminal, and forwards each item's Title
to the Azure AI Foundry agent named by the AGENT_NAME environment variable.
"""

import os
import sys
from datetime import date, timezone, datetime
from urllib.parse import urlparse

import requests
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential, DefaultAzureCredential
from azure.ai.projects import AIProjectClient
from openai import OpenAI

load_dotenv()

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SHAREPOINT_SITE_URL = os.environ.get("SHAREPOINT_SITE_URL")
SHAREPOINT_LIST_NAME = os.environ.get("SHAREPOINT_LIST_NAME")

SHAREPOINT_CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET")
AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID")

AZURE_AI_PROJECT_ENDPOINT = os.environ.get("AZURE_AI_PROJECT_ENDPOINT")

AGENT_NAME = os.environ.get("AGENT_NAME")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def validate_environment_configuration() -> None:
    """Validate required environment variables for this script."""
    required = {
        "SHAREPOINT_SITE_URL": SHAREPOINT_SITE_URL,
        "SHAREPOINT_LIST_NAME": SHAREPOINT_LIST_NAME,
        "SHAREPOINT_CLIENT_ID": SHAREPOINT_CLIENT_ID,
        "SHAREPOINT_CLIENT_SECRET": SHAREPOINT_CLIENT_SECRET,
        "AZURE_TENANT_ID": AZURE_TENANT_ID,
        "AZURE_AI_PROJECT_ENDPOINT": AZURE_AI_PROJECT_ENDPOINT,
        "AGENT_NAME": AGENT_NAME,
    }

    missing = [name for name, value in required.items() if not value]
    if missing:
        raise EnvironmentError(
            "Missing required environment variables: "
            + ", ".join(sorted(missing))
        )


# ---------------------------------------------------------------------------
# Microsoft Graph helpers
# ---------------------------------------------------------------------------

def get_graph_token() -> str:
    """Acquire a Microsoft Graph access token using client credentials."""
    credential = ClientSecretCredential(
        tenant_id=AZURE_TENANT_ID,
        client_id=SHAREPOINT_CLIENT_ID,
        client_secret=SHAREPOINT_CLIENT_SECRET,
    )
    token = credential.get_token("https://graph.microsoft.com/.default")
    return token.token


def graph_headers(token: str) -> dict:
    """Return standard headers for Graph API requests."""
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }


def resolve_site_id(token: str) -> str:
    """Resolve the SharePoint site URL to a Graph site ID."""
    parsed = urlparse(SHAREPOINT_SITE_URL)
    hostname = parsed.hostname  # e.g. inodigitaldemooutlook.sharepoint.com
    site_path = parsed.path.rstrip("/")  # e.g. /sites/rogfaste03fcfbcc

    url = f"{GRAPH_BASE}/sites/{hostname}:{site_path}"
    resp = requests.get(url, headers=graph_headers(token), timeout=30)
    resp.raise_for_status()
    site = resp.json()
    print(f"  Site name: {site.get('displayName', 'N/A')}")
    return site["id"]


def resolve_list_id(token: str, site_id: str) -> str:
    """Resolve a list by display name to its Graph list ID."""
    url = f"{GRAPH_BASE}/sites/{site_id}/lists"
    resp = requests.get(url, headers=graph_headers(token), timeout=30)
    resp.raise_for_status()
    for sp_list in resp.json().get("value", []):
        if sp_list.get("displayName") == SHAREPOINT_LIST_NAME:
            return sp_list["id"]
    raise ValueError(
        f"List '{SHAREPOINT_LIST_NAME}' not found on site. "
        f"Available lists: {[l['displayName'] for l in resp.json().get('value', [])]}"
    )


def get_items_created_today(token: str, site_id: str, list_id: str) -> list:
    """Return all items whose Created date is today (UTC) via Graph API."""
    today = date.today().isoformat()  # e.g. "2026-03-20"

    # Graph doesn't support $filter on fields/Created for list items,
    # so we fetch all items and filter client-side.
    items = []
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?$expand=fields"

    while url:
        resp = requests.get(url, headers=graph_headers(token), timeout=30)
        resp.raise_for_status()
        data = resp.json()
        for item in data.get("value", []):
            created = item.get("createdDateTime", "")
            if created.startswith(today):
                items.append(item)
        url = data.get("@odata.nextLink")

    return items


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
    # ── 0. Validate environment configuration ───────────────────────────────
    try:
        validate_environment_configuration()
    except Exception as exc:
        print(f"ERROR: Invalid environment configuration – {exc}", file=sys.stderr)
        sys.exit(1)

    # ── 1. Authenticate with Microsoft Graph ────────────────────────────────
    print("Acquiring Microsoft Graph token…")
    try:
        token = get_graph_token()
        print("Token acquired.\n")
    except Exception as exc:
        print(f"ERROR: Could not acquire Graph token – {exc}", file=sys.stderr)
        sys.exit(1)

    # ── 2. Resolve SharePoint site and list ─────────────────────────────────
    print(f"Resolving SharePoint site: {SHAREPOINT_SITE_URL}")
    try:
        site_id = resolve_site_id(token)
        print(f"  Site ID: {site_id}\n")
    except Exception as exc:
        print(f"ERROR: Could not resolve SharePoint site – {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Resolving list: '{SHAREPOINT_LIST_NAME}'")
    try:
        list_id = resolve_list_id(token, site_id)
        print(f"  List ID: {list_id}\n")
    except Exception as exc:
        print(f"ERROR: Could not resolve SharePoint list – {exc}", file=sys.stderr)
        sys.exit(1)

    # ── 3. Fetch items created today ────────────────────────────────────────
    print(f"Fetching items created today ({date.today()})…\n")
    try:
        items = get_items_created_today(token, site_id, list_id)
    except Exception as exc:
        print(f"ERROR: Could not read SharePoint list items – {exc}", file=sys.stderr)
        sys.exit(1)

    if not items:
        print("No new items found today. Exiting.")
        return

    # ── 4. Print all items to the terminal ─────────────────────────────────
    print(f"Found {len(items)} item(s) created today:\n")
    for item in items:
        fields = item.get("fields", {})
        print(
            f"  ID: {fields.get('id', '?'):>6}  "
            f"Title: {fields.get('Title', '(no title)')}"
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
        title = item.get("fields", {}).get("Title", "").strip()
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
