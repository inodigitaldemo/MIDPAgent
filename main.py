"""
MIDPAgent – Master Information Delivery Plan Agent

Connects to a SharePoint list via Microsoft Graph API, retrieves items
created today, sends each item's full field data as JSON to the Azure AI
Foundry agent, parses the generated Markdown response, saves .md files
locally, and uploads them to the SharePoint site's default document library.
"""

import json
import os
import re
import sys
import time
from datetime import date
from pathlib import Path
from urllib.parse import urlparse

import requests
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential, AzureCliCredential

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
        f"Available lists: {[lst['displayName'] for lst in resp.json().get('value', [])]}"
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
# Azure AI Foundry Agents helpers (REST API)
# ---------------------------------------------------------------------------

def get_foundry_token() -> str:
    """Acquire a token for the Azure AI Foundry Agents API via Azure CLI."""
    credential = AzureCliCredential()
    return credential.get_token("https://ai.azure.com/.default").token


def foundry_headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def get_agent_definition(token: str) -> dict:
    """Fetch the Foundry agent and return its latest version definition."""
    url = f"{AZURE_AI_PROJECT_ENDPOINT}/agents/{AGENT_NAME}?api-version=v1"
    resp = requests.get(url, headers=foundry_headers(token), timeout=30)
    resp.raise_for_status()
    return resp.json()["versions"]["latest"]["definition"]


def ensure_assistant(token: str) -> str:
    """Ensure an OpenAI-compatible assistant exists for the Foundry agent.

    Foundry 'prompt' agents aren't automatically exposed as OpenAI assistants.
    This creates one (idempotent by name) and returns its assistant ID.
    """
    headers = foundry_headers(token)

    # Check if an assistant already exists with this name
    resp = requests.get(
        f"{AZURE_AI_PROJECT_ENDPOINT}/assistants?api-version=v1",
        headers=headers, timeout=30,
    )
    resp.raise_for_status()
    for asst in resp.json().get("data", []):
        if asst.get("name") == AGENT_NAME:
            return asst["id"]

    # Create from agent definition
    agent_def = get_agent_definition(token)
    resp = requests.post(
        f"{AZURE_AI_PROJECT_ENDPOINT}/assistants?api-version=v1",
        headers=headers,
        json={
            "model": agent_def["model"],
            "name": AGENT_NAME,
            "instructions": agent_def.get("instructions", ""),
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["id"]


def send_item_to_agent(token: str, assistant_id: str, fields: dict) -> str:
    """Create a thread, post the item fields, run the assistant, poll until
    complete, and return the assistant's text reply."""
    headers = foundry_headers(token)
    base = AZURE_AI_PROJECT_ENDPOINT
    payload = json.dumps(fields, indent=2, default=str)

    # Create thread
    resp = requests.post(f"{base}/threads?api-version=v1",
                         headers=headers, json={}, timeout=30)
    resp.raise_for_status()
    thread_id = resp.json()["id"]

    # Post user message
    requests.post(
        f"{base}/threads/{thread_id}/messages?api-version=v1",
        headers=headers,
        json={"role": "user", "content": payload},
        timeout=30,
    ).raise_for_status()

    # Start run
    resp = requests.post(
        f"{base}/threads/{thread_id}/runs?api-version=v1",
        headers=headers,
        json={"assistant_id": assistant_id},
        timeout=30,
    )
    resp.raise_for_status()
    run_id = resp.json()["id"]

    # Poll until terminal state
    while True:
        time.sleep(2)
        resp = requests.get(
            f"{base}/threads/{thread_id}/runs/{run_id}?api-version=v1",
            headers=headers, timeout=30,
        )
        resp.raise_for_status()
        status = resp.json()["status"]
        if status in ("completed", "failed", "cancelled", "expired"):
            break

    if status != "completed":
        return f"[Run ended with status '{status}']"

    # Retrieve messages
    resp = requests.get(
        f"{base}/threads/{thread_id}/messages?api-version=v1",
        headers=headers, timeout=30,
    )
    resp.raise_for_status()
    for msg in resp.json().get("data", []):
        if msg["role"] == "assistant":
            for block in msg.get("content", []):
                if block.get("type") == "text":
                    return block["text"]["value"]

    return "[No response from agent]"


# ---------------------------------------------------------------------------
# Markdown output helpers
# ---------------------------------------------------------------------------

def parse_agent_response(response: str) -> tuple[str | None, str | None]:
    """Extract Markdown content and filename from the agent's response.

    Returns (markdown_text, filename).  Either may be None if the agent
    returned an unexpected format.
    """
    # Extract fenced code block content (```markdown ... ``` or ``` ... ```)
    code_block = re.search(r"```(?:markdown|md)?\s*\n(.*?)```", response, re.DOTALL)
    markdown_text = code_block.group(1).strip() if code_block else None

    # Extract filename – agent is instructed to use [DocID]_[Title].md
    filename_match = re.search(r"`([^`]+\.md)`", response)
    if not filename_match:
        filename_match = re.search(r"(\S+\.md)", response)
    filename = filename_match.group(1).strip() if filename_match else None

    return markdown_text, filename


def save_markdown_locally(markdown_text: str, filename: str) -> Path:
    """Write *markdown_text* to output/<filename> and return the path."""
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)
    # Sanitise filename – keep only safe characters
    safe_name = re.sub(r'[<>:"/\\|?*]', "_", filename)
    path = output_dir / safe_name
    path.write_text(markdown_text, encoding="utf-8")
    return path


def upload_to_sharepoint(token: str, site_id: str, list_id: str,
                         item_id: str, filename: str, content: str) -> None:
    """Attach a file to a specific SharePoint list item."""
    safe_name = re.sub(r'[<>:"|?*]', "_", filename)
    url = (f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}"
           f"/items/{item_id}/attachments")
    resp = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={
            "name": safe_name,
            "contentBytes": __import__("base64").b64encode(
                content.encode("utf-8")
            ).decode("ascii"),
        },
        timeout=30,
    )
    resp.raise_for_status()
    print(f"  Attached to list item {item_id}: {safe_name}")


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

    # ── 4. Connect to Azure AI Foundry agent ─────────────────────────────
    print("Acquiring Azure AI Foundry token…")
    try:
        foundry_token = get_foundry_token()
    except Exception as exc:
        print(f"ERROR: Could not acquire Foundry token – {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Ensuring assistant for agent '{AGENT_NAME}'…")
    try:
        assistant_id = ensure_assistant(foundry_token)
    except Exception as exc:
        print(f"ERROR: Could not create/find assistant – {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Assistant ready (id={assistant_id}). Sending items…\n")

    # ── 5. Send each item to the agent and handle Markdown output ─────────
    for item in items:
        fields = item.get("fields", {})
        title = fields.get("Title", "").strip()
        if not title:
            continue

        print(f"→ Sending item: '{title}'")
        try:
            response = send_item_to_agent(foundry_token, assistant_id, fields)
        except Exception as exc:
            print(f"  ERROR processing '{title}': {exc}\n", file=sys.stderr)
            continue

        markdown_text, filename = parse_agent_response(response)

        if not markdown_text:
            print(f"  Agent did not return Markdown. Raw response:\n{response}\n")
            continue

        if not filename:
            filename = re.sub(r"\s+", "_", title) + ".md"
            print(f"  No filename detected; using '{filename}'")

        # Save locally
        local_path = save_markdown_locally(markdown_text, filename)
        print(f"  Saved locally: {local_path}")

        # Attach to the SharePoint list item
        item_id = str(fields.get("id", ""))
        try:
            upload_to_sharepoint(token, site_id, list_id, item_id,
                                 filename, markdown_text)
        except Exception as exc:
            print(f"  WARNING: SharePoint attachment failed – {exc}")

        print()


if __name__ == "__main__":
    main()
