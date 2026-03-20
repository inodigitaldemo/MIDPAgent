"""
MIDPAgent – Master Information Delivery Plan Agent

Connects to a SharePoint list via Microsoft Graph API, retrieves items
created today, sends each item's full field data as JSON to the Azure AI
Foundry agent, parses the generated Markdown response, saves .md files
locally, and uploads them to the SharePoint site's default document library.
"""

import io
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
from pypdf import PdfReader
from md_to_docx import convert_md_to_docx

load_dotenv("env")

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

SHAREPOINT_REFERENCE_LIST_NAME = os.environ.get(
    "SHAREPOINT_REFERENCE_LIST_NAME", "ArbeidsromYM"
)

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
    today = date.today().isoformat()

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
# ArbeidsromYM reference PDF helpers
# ---------------------------------------------------------------------------

def _extract_pdf_text(pdf_bytes: bytes) -> str:
    """Extract text from raw PDF bytes using pypdf."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            pages.append(text)
    return "\n\n".join(pages)


def fetch_reference_pdfs(token: str, site_id: str) -> list[dict[str, str]]:
    """Download PDFs from the ArbeidsromYM library and extract their text.

    Returns a list of dicts: [{"name": "file.pdf", "text": "..."}].
    """
    headers = graph_headers(token)

    # Resolve the document library by display name
    url = f"{GRAPH_BASE}/sites/{site_id}/lists"
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()

    library_id = None
    for sp_list in resp.json().get("value", []):
        if sp_list.get("displayName") == SHAREPOINT_REFERENCE_LIST_NAME:
            library_id = sp_list["id"]
            break

    if not library_id:
        print(
            f"  WARNING: Reference library '{SHAREPOINT_REFERENCE_LIST_NAME}' "
            "not found on the site."
        )
        return []

    # List files in the library's root folder
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/lists/{library_id}"
        "/drive/root/children?$select=id,name,@microsoft.graph.downloadUrl"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()

    pdfs: list[dict[str, str]] = []
    for item in resp.json().get("value", []):
        name = item.get("name", "")
        if not name.lower().endswith(".pdf"):
            continue

        # Prefer the pre-authenticated download URL; fall back to Graph
        download_url = item.get("@microsoft.graph.downloadUrl")
        if download_url:
            content_resp = requests.get(download_url, timeout=60)
        else:
            content_url = (
                f"{GRAPH_BASE}/sites/{site_id}/lists/{library_id}"
                f"/drive/items/{item['id']}/content"
            )
            content_resp = requests.get(
                content_url, headers=headers, timeout=60
            )
        content_resp.raise_for_status()

        text = _extract_pdf_text(content_resp.content)
        if text.strip():
            pdfs.append({"name": name, "text": text})
            print(f"  Extracted text from: {name} ({len(text)} chars)")

    return pdfs


def build_reference_context(pdfs: list[dict[str, str]]) -> str:
    """Combine extracted PDF texts into a single reference context string."""
    if not pdfs:
        return ""
    sections = []
    for pdf in pdfs:
        sections.append(f"### {pdf['name']}\n\n{pdf['text']}")
    return "\n\n---\n\n".join(sections)


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
    If the assistant already exists, its instructions are updated to match
    the latest Foundry agent definition so changes take effect immediately.
    """
    headers = foundry_headers(token)
    agent_def = get_agent_definition(token)
    latest_instructions = agent_def.get("instructions", "")

    # Check if an assistant already exists with this name
    resp = requests.get(
        f"{AZURE_AI_PROJECT_ENDPOINT}/assistants?api-version=v1",
        headers=headers, timeout=30,
    )
    resp.raise_for_status()
    for asst in resp.json().get("data", []):
        if asst.get("name") == AGENT_NAME:
            # Update the existing assistant so instruction changes take effect
            resp = requests.post(
                f"{AZURE_AI_PROJECT_ENDPOINT}/assistants/{asst['id']}?api-version=v1",
                headers=headers,
                json={
                    "model": agent_def["model"],
                    "name": AGENT_NAME,
                    "instructions": latest_instructions,
                },
                timeout=30,
            )
            resp.raise_for_status()
            print(f"  Updated assistant instructions (id={asst['id']})")
            return asst["id"]

    # Create from agent definition
    resp = requests.post(
        f"{AZURE_AI_PROJECT_ENDPOINT}/assistants?api-version=v1",
        headers=headers,
        json={
            "model": agent_def["model"],
            "name": AGENT_NAME,
            "instructions": latest_instructions,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["id"]


def send_item_to_agent(
    token: str,
    assistant_id: str,
    fields: dict,
    reference_context: str = "",
) -> str:
    """Create a thread, post the item fields (with optional reference
    context from ArbeidsromYM PDFs), run the assistant, poll until
    complete, and return the assistant's text reply."""
    headers = foundry_headers(token)
    base = AZURE_AI_PROJECT_ENDPOINT
    payload = json.dumps(fields, indent=2, default=str)

    # Build the user message: MIDP data + optional reference PDFs
    if reference_context:
        message_content = (
            "## MIDP Item Data\n\n"
            f"```json\n{payload}\n```\n\n"
            "## Reference Document Templates (from ArbeidsromYM)\n\n"
            "The following text was extracted from reference PDF documents. "
            "Use these to determine which MIDP fields are relevant and how "
            "to structure the output document. Only include fields that "
            "align with what these reference documents expect.\n\n"
            f"{reference_context}"
        )
    else:
        message_content = payload

    # Create thread
    resp = requests.post(f"{base}/threads?api-version=v1",
                         headers=headers, json={}, timeout=30)
    resp.raise_for_status()
    thread_id = resp.json()["id"]

    # Post user message
    requests.post(
        f"{base}/threads/{thread_id}/messages?api-version=v1",
        headers=headers,
        json={"role": "user", "content": message_content},
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


def upload_to_sharepoint(token: str, site_id: str, filename: str,
                         content: bytes) -> None:
    """Upload a file to the default document library on the SharePoint site."""
    safe_name = re.sub(r'[<>:"|?*]', "_", filename)
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{safe_name}:/content"
    resp = requests.put(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        },
        data=content,
        timeout=60,
    )
    resp.raise_for_status()
    print(f"  Uploaded to SharePoint: {resp.json().get('webUrl', safe_name)}")


# ---------------------------------------------------------------------------
# ArbeidsromYM document library helpers
# ---------------------------------------------------------------------------

def resolve_library_drive_id(token: str, site_id: str, library_name: str) -> str | None:
    """Resolve a document library's drive ID by display name.

    Returns the drive ID string, or None if the library is not found.
    """
    headers = graph_headers(token)

    # First resolve the list ID for the library
    url = f"{GRAPH_BASE}/sites/{site_id}/lists"
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()

    library_id = None
    for sp_list in resp.json().get("value", []):
        if sp_list.get("displayName") == library_name:
            library_id = sp_list["id"]
            break

    if not library_id:
        return None

    # Get the drive associated with this list
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{library_id}/drive"
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json()["id"]


def find_existing_document(
    token: str, site_id: str, drive_id: str, doc_id: str
) -> str | None:
    """Search the library for an existing document whose name starts with the
    MIDP item's DocID.  Returns the webUrl if found, or None.
    """
    headers = graph_headers(token)

    # List root children and look for a file starting with the DocID
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}"
        f"/root/children?$select=id,name,webUrl"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()

    for item in resp.json().get("value", []):
        name = item.get("name", "")
        if name.lower().startswith(doc_id.lower()):
            return item.get("webUrl")

    return None


def upload_to_library(
    token: str, site_id: str, drive_id: str, filename: str, content: bytes
) -> str:
    """Upload a file to a specific document library drive.

    Returns the webUrl of the uploaded file.
    """
    safe_name = re.sub(r'[<>:"|?*]', "_", filename)
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}"
        f"/root:/{safe_name}:/content"
    )
    resp = requests.put(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        },
        data=content,
        timeout=60,
    )
    resp.raise_for_status()
    web_url = resp.json().get("webUrl", "")
    print(f"  Uploaded to library: {web_url}")
    return web_url


def update_midp_item_link(
    token: str, site_id: str, list_id: str, item_id: str, doc_url: str
) -> None:
    """Update the Arbeidsdokument column on the MIDP list item with a link."""
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}"
        f"/items/{item_id}/fields"
    )
    resp = requests.patch(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={"Arbeidsdokument": doc_url},
        timeout=30,
    )
    resp.raise_for_status()
    print(f"  Updated MIDP Arbeidsdokument link: {doc_url}")


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
    # ── 2b. Resolve ArbeidsromYM document library drive ──────────────────
    print(f"Resolving document library: '{SHAREPOINT_REFERENCE_LIST_NAME}'")
    ref_drive_id = resolve_library_drive_id(
        token, site_id, SHAREPOINT_REFERENCE_LIST_NAME
    )
    if ref_drive_id:
        print(f"  Drive ID: {ref_drive_id}\n")
    else:
        print(
            f"  WARNING: Library '{SHAREPOINT_REFERENCE_LIST_NAME}' not found. "
            "Documents will be uploaded to the site root instead.\n"
        )
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

    # ── 3b. Print all items to the terminal ────────────────────────────────
    print(f"Found {len(items)} item(s) created today:\n")
    for item in items:
        fields = item.get("fields", {})
        print(
            f"  ID: {fields.get('id', '?'):>6}  "
            f"Title: {fields.get('Title', '(no title)')}"
        )
    print()

    # ── 4. Fetch reference PDFs from ArbeidsromYM ──────────────────────────
    print(f"Fetching reference PDFs from '{SHAREPOINT_REFERENCE_LIST_NAME}'…")
    try:
        reference_pdfs = fetch_reference_pdfs(token, site_id)
        reference_context = build_reference_context(reference_pdfs)
        if reference_pdfs:
            print(f"  Loaded {len(reference_pdfs)} reference PDF(s).\n")
        else:
            print("  No reference PDFs found. Agent will use MIDP data only.\n")
    except Exception as exc:
        print(f"  WARNING: Could not fetch reference PDFs – {exc}")
        reference_context = ""

    # ── 5. Connect to Azure AI Foundry agent ─────────────────────────────
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

    # ── 6. Send each item to the agent and handle Markdown output ─────────
    for item in items:
        fields = item.get("fields", {})
        title = fields.get("Title", "").strip()
        if not title:
            continue

        print(f"→ Sending item: '{title}'")
        try:
            response = send_item_to_agent(
                foundry_token, assistant_id, fields, reference_context
            )
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

        # Save .md locally
        local_path = save_markdown_locally(markdown_text, filename)
        print(f"  Saved locally: {local_path}")

        # Convert .md → .docx
        try:
            _, docx_path = convert_md_to_docx(local_path, output_dir=local_path.parent)
            print(f"  Converted to DOCX: {docx_path}")
        except Exception as exc:
            print(f"  WARNING: DOCX conversion failed – {exc}")
            continue

        # Upload .docx to the discipline's document library
        doc_id = fields.get("DocID", fields.get("Title", "")).strip()
        item_id = fields.get("id") or item.get("id")

        if ref_drive_id:
            # Check if a document already exists in ArbeidsromYM for this item
            existing_url = find_existing_document(
                token, site_id, ref_drive_id, doc_id
            )
            if existing_url:
                print(f"  Document already exists in library: {existing_url}")
                print(f"  Skipping upload.")
            else:
                try:
                    web_url = upload_to_library(
                        token, site_id, ref_drive_id,
                        docx_path.name, docx_path.read_bytes(),
                    )
                    # Update the MIDP item's Arbeidsdokument column
                    if item_id and web_url:
                        update_midp_item_link(
                            token, site_id, list_id, item_id, web_url
                        )
                except Exception as exc:
                    print(f"  WARNING: Library upload failed – {exc}")
        else:
            # Fallback: upload to site root if library not found
            try:
                upload_to_sharepoint(
                    token, site_id, docx_path.name, docx_path.read_bytes()
                )
            except Exception as exc:
                print(f"  WARNING: SharePoint upload failed – {exc}")

        print()


if __name__ == "__main__":
    main()
