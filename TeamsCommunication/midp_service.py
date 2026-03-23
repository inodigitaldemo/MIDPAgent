"""
MIDP Service – SharePoint polling, document production, and approval.

Responsibilities:
1.  Poll the MIDP SharePoint list every 60 s for new planned items.
2.  Post a *produce_document* adaptive card to the Teams channel for each
    new item found (yes / no prompt).
3.  On "yes" → fetch item metadata, send to the Foundry agent, parse the
    Markdown response, convert via md_to_docx, upload the .docx to the
    ArbeidsromYM library, and post a *document_approval* card.
4.  On "approve" → mark the item as approved in SharePoint (update the
    Arbeidsdokument column with the document URL).
"""

from __future__ import annotations

import asyncio
import io
import json
import logging
import re
import tempfile
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

import aiohttp
from azure.identity import ClientSecretCredential

from .config import BotConfig

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ── Polling tunables ──────────────────────────────────────────────────────
POLL_INTERVAL_SECONDS = 60


class MIDPService:
    """Async service that bridges SharePoint, Foundry, and md_to_docx."""

    def __init__(
        self,
        config: BotConfig,
        agent_service=None,          # FoundryAgentService (optional)
    ) -> None:
        self._config = config
        self._agent_service = agent_service

        # Graph credential (same app registration, Graph scope)
        self._graph_credential = ClientSecretCredential(
            tenant_id=config.tenant_id or "",
            client_id=config.sharepoint_client_id or config.app_id,
            client_secret=config.sharepoint_client_secret or config.app_password,
        )

        # Resolved IDs (populated lazily)
        self._site_id: Optional[str] = None
        self._list_id: Optional[str] = None
        self._drive_id: Optional[str] = None  # ArbeidsromYM drive

        # Track items we've already prompted about (in-memory)
        self._seen_item_ids: set[str] = set()

        # Background polling task handle
        self._poll_task: Optional[asyncio.Task] = None

    # ------------------------------------------------------------------
    # Graph token
    # ------------------------------------------------------------------

    async def _graph_token(self) -> str:
        token = await asyncio.to_thread(
            self._graph_credential.get_token,
            "https://graph.microsoft.com/.default",
        )
        return token.token

    def _graph_headers(self, token: str) -> dict:
        return {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        }

    # ------------------------------------------------------------------
    # Lazy resolution helpers
    # ------------------------------------------------------------------

    async def _ensure_site_id(self, token: str) -> str:
        if self._site_id:
            return self._site_id
        parsed = urlparse(self._config.sharepoint_site_url)
        hostname = parsed.hostname
        site_path = parsed.path.rstrip("/")
        url = f"{GRAPH_BASE}/sites/{hostname}:{site_path}"
        async with aiohttp.ClientSession() as session:
            async with session.get(
                url, headers=self._graph_headers(token)
            ) as resp:
                resp.raise_for_status()
                data = await resp.json()
                self._site_id = data["id"]
                logger.info("Resolved site ID: %s", self._site_id)
                return self._site_id

    async def _ensure_list_id(self, token: str) -> str:
        if self._list_id:
            return self._list_id
        site_id = await self._ensure_site_id(token)
        url = f"{GRAPH_BASE}/sites/{site_id}/lists"
        async with aiohttp.ClientSession() as session:
            async with session.get(
                url, headers=self._graph_headers(token)
            ) as resp:
                resp.raise_for_status()
                for sp_list in (await resp.json()).get("value", []):
                    if sp_list.get("displayName") == self._config.sharepoint_list_name:
                        self._list_id = sp_list["id"]
                        logger.info("Resolved list ID: %s", self._list_id)
                        return self._list_id
        raise ValueError(
            f"List '{self._config.sharepoint_list_name}' not found on site."
        )

    async def _ensure_drive_id(self, token: str) -> Optional[str]:
        if self._drive_id:
            return self._drive_id
        site_id = await self._ensure_site_id(token)
        lib_name = self._config.sharepoint_reference_library
        url = f"{GRAPH_BASE}/sites/{site_id}/lists"
        async with aiohttp.ClientSession() as session:
            async with session.get(
                url, headers=self._graph_headers(token)
            ) as resp:
                resp.raise_for_status()
                library_id = None
                for sp_list in (await resp.json()).get("value", []):
                    if sp_list.get("displayName") == lib_name:
                        library_id = sp_list["id"]
                        break
            if not library_id:
                logger.warning("Library '%s' not found.", lib_name)
                return None
            async with session.get(
                f"{GRAPH_BASE}/sites/{site_id}/lists/{library_id}/drive",
                headers=self._graph_headers(token),
            ) as resp:
                resp.raise_for_status()
                self._drive_id = (await resp.json())["id"]
                logger.info("Resolved drive ID: %s", self._drive_id)
                return self._drive_id

    # ------------------------------------------------------------------
    # SharePoint list helpers
    # ------------------------------------------------------------------

    async def _get_new_planned_items(self, token: str) -> list[dict]:
        """Fetch MIDP items with StatusIM='Under arbeid' that haven't been prompted yet."""
        site_id = await self._ensure_site_id(token)
        list_id = await self._ensure_list_id(token)

        items: list[dict] = []
        url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?$expand=fields"

        async with aiohttp.ClientSession() as session:
            while url:
                async with session.get(
                    url, headers=self._graph_headers(token)
                ) as resp:
                    resp.raise_for_status()
                    data = await resp.json()
                    for item in data.get("value", []):
                        item_id = str(item.get("id", ""))
                        fields = item.get("fields", {})
                        status_im = (fields.get("StatusIM") or "").strip()
                        if (
                            status_im == "Under arbeid"
                            and item_id not in self._seen_item_ids
                        ):
                            items.append(item)
                    url = data.get("@odata.nextLink")

        return items

    async def _get_item_by_id(self, token: str, item_id: str) -> dict:
        """Fetch a single MIDP list item by its ID."""
        site_id = await self._ensure_site_id(token)
        list_id = await self._ensure_list_id(token)
        url = (
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}"
            f"/items/{item_id}?$expand=fields"
        )
        async with aiohttp.ClientSession() as session:
            async with session.get(
                url, headers=self._graph_headers(token)
            ) as resp:
                resp.raise_for_status()
                return await resp.json()

    # ------------------------------------------------------------------
    # Reference PDFs (ArbeidsromYM)
    # ------------------------------------------------------------------

    async def _fetch_reference_context(self, token: str) -> str:
        """Download PDFs from ArbeidsromYM and build a reference text blob."""
        site_id = await self._ensure_site_id(token)
        drive_id = await self._ensure_drive_id(token)
        if not drive_id:
            return ""

        url = (
            f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}"
            "/root/children?$select=id,name,@microsoft.graph.downloadUrl"
        )

        sections: list[str] = []
        async with aiohttp.ClientSession() as session:
            async with session.get(
                url, headers=self._graph_headers(token)
            ) as resp:
                resp.raise_for_status()
                children = (await resp.json()).get("value", [])

            for child in children:
                name = child.get("name", "")
                if not name.lower().endswith(".pdf"):
                    continue
                download_url = child.get("@microsoft.graph.downloadUrl")
                if not download_url:
                    continue
                async with session.get(download_url) as dl_resp:
                    if dl_resp.status != 200:
                        continue
                    pdf_bytes = await dl_resp.read()
                try:
                    from pypdf import PdfReader
                    reader = PdfReader(io.BytesIO(pdf_bytes))
                    text = "\n\n".join(
                        p.extract_text() or "" for p in reader.pages
                    ).strip()
                    if text:
                        sections.append(f"### {name}\n\n{text}")
                        logger.info("Extracted reference PDF: %s (%d chars)", name, len(text))
                except Exception as exc:
                    logger.warning("Could not read PDF %s: %s", name, exc)

        return "\n\n---\n\n".join(sections) if sections else ""

    # ------------------------------------------------------------------
    # Document upload
    # ------------------------------------------------------------------

    async def _upload_to_library(
        self, token: str, filename: str, content: bytes
    ) -> str:
        """Upload a file to the ArbeidsromYM library. Returns the webUrl."""
        site_id = await self._ensure_site_id(token)
        drive_id = await self._ensure_drive_id(token)
        if not drive_id:
            raise RuntimeError("ArbeidsromYM library not resolved.")

        safe_name = re.sub(r'[<>:"|?*]', "_", filename)
        url = (
            f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}"
            f"/root:/{safe_name}:/content"
        )
        async with aiohttp.ClientSession() as session:
            async with session.put(
                url,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/octet-stream",
                },
                data=content,
            ) as resp:
                resp.raise_for_status()
                data = await resp.json()
                web_url = data.get("webUrl", "")
                logger.info("Uploaded to library: %s", web_url)
                return web_url

    async def _update_midp_item_fields(
        self, token: str, item_id: str, field_updates: dict
    ) -> None:
        """Patch one or more columns on a MIDP list item."""
        site_id = await self._ensure_site_id(token)
        list_id = await self._ensure_list_id(token)
        url = (
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}"
            f"/items/{item_id}/fields"
        )
        async with aiohttp.ClientSession() as session:
            async with session.patch(
                url,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                },
                json=field_updates,
            ) as resp:
                resp.raise_for_status()
                logger.info(
                    "Updated MIDP item %s fields: %s", item_id, field_updates
                )

    # ------------------------------------------------------------------
    # Document production pipeline
    # ------------------------------------------------------------------

    async def produce_document(self, item_id: str) -> dict:
        """Full pipeline: fetch item → agent → md_to_docx → upload → approval info.

        Returns a dict with keys: doc_url, filename, error (if any).
        """
        try:
            token = await self._graph_token()
            item = await self._get_item_by_id(token, item_id)
            fields = item.get("fields", {})
            title = fields.get("Title", "Untitled")

            # Build reference context from ArbeidsromYM PDFs
            reference_context = await self._fetch_reference_context(token)

            # Send to Foundry agent
            if not self._agent_service:
                return {"error": "Agent service not configured"}

            payload = json.dumps(fields, indent=2, default=str)
            if reference_context:
                message_content = (
                    "## MIDP Item Data\n\n"
                    f"```json\n{payload}\n```\n\n"
                    "## Reference Document Templates (from ArbeidsromYM)\n\n"
                    "The following text was extracted from reference PDF documents. "
                    "Use these to determine which MIDP fields are relevant and how "
                    "to structure the output document. Only include fields that "
                    "align with what these reference documents expect. Return only the markdown that should be in the document, without any other comments included.\n\n"
                    f"{reference_context}"
                )
            else:
                message_content = payload

            agent_reply = await self._agent_service.send_message(message_content)

            markdown_text = agent_reply
          
            filename = re.sub(r"\s+", "_", title) + ".md"

            # Save markdown to a temp directory, convert to .docx
            with tempfile.TemporaryDirectory() as tmp_dir:
                md_path = Path(tmp_dir) / filename
                md_path.write_text(markdown_text, encoding="utf-8")

                from md_to_docx import convert_md_to_docx

                _, docx_path = convert_md_to_docx(md_path, output_dir=Path(tmp_dir))
                docx_bytes = Path(docx_path).read_bytes()
                docx_filename = Path(docx_path).name

            # Upload to ArbeidsromYM
            web_url = await self._upload_to_library(token, docx_filename, docx_bytes)

            # Update Arbeidsdokument link + set StatusIM to "Til revisjon"
            if web_url:
                await self._update_midp_item_fields(token, item_id, {
                    "Arbeidsdokument": web_url,
                    "StatusIM": "Til revisjon",
                })

            return {"doc_url": web_url, "filename": docx_filename}

        except Exception as exc:
            logger.error("produce_document failed for item %s: %s", item_id, exc, exc_info=True)
            return {"error": str(exc)}

    # ------------------------------------------------------------------
    # Approval
    # ------------------------------------------------------------------

    async def mark_approved(self, item_id: str) -> None:
        """Mark a MIDP item as approved (no-op placeholder – extend as needed)."""
        logger.info("Item %s marked as approved.", item_id)

    # ------------------------------------------------------------------
    # Agent response parsing (mirrors main.py logic)
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_agent_response(response: str) -> tuple[Optional[str], Optional[str]]:
        code_block = re.search(
            r"```(?:markdown|md)?\s*\n(.*?)```", response, re.DOTALL
        )
        markdown_text = code_block.group(1).strip() if code_block else None

        filename_match = re.search(r"`([^`]+\.md)`", response)
        if not filename_match:
            filename_match = re.search(r"(\S+\.md)", response)
        filename = filename_match.group(1).strip() if filename_match else None

        return markdown_text, filename

    # ------------------------------------------------------------------
    # Background polling
    # ------------------------------------------------------------------

    async def start_polling(self) -> None:
        """Start the background polling task."""
        if self._poll_task and not self._poll_task.done():
            logger.warning("Polling task already running.")
            return
        self._poll_task = asyncio.create_task(self._poll_loop())
        logger.info("MIDP polling started (interval=%ds).", POLL_INTERVAL_SECONDS)

    async def stop_polling(self) -> None:
        """Cancel the background polling task."""
        if self._poll_task:
            self._poll_task.cancel()
            try:
                await self._poll_task
            except asyncio.CancelledError:
                pass
            self._poll_task = None
            logger.info("MIDP polling stopped.")

    async def _poll_loop(self) -> None:
        """Infinite loop: check for new items → post produce card."""
        from .adaptive_cards import produce_document_attachment
        from .proactive import send_to_channel

        # Give the bot a few seconds to finish startup
        await asyncio.sleep(5)

        while True:
            try:
                token = await self._graph_token()
                new_items = await self._get_new_planned_items(token)
                if new_items:
                    logger.info(
                        "Poll found %d new item(s) in MIDP list.", len(new_items)
                    )
                for item in new_items:
                    item_id = str(item.get("id", ""))
                    fields = item.get("fields", {})
                    title = fields.get("Title", "(no title)")
                    dokumentnummer = fields.get("Dokumentnummer", "")
                    dokumenttype = fields.get("Dokumenttype", "")
                    disiplin = fields.get("Disiplin", "")

                    try:
                        attachment = produce_document_attachment(
                            title=title,
                            item_id=item_id,
                            dokumentnummer=dokumentnummer,
                            dokumenttype=dokumenttype,
                            disiplin=disiplin,
                        )
                        await send_to_channel(self._config, attachment)
                        logger.info(
                            "Posted produce card for item %s (%s).",
                            item_id, title,
                        )
                    except Exception as exc:
                        logger.error(
                            "Failed to post produce card for item %s: %s",
                            item_id, exc,
                        )
                    # Mark as seen regardless of success to avoid spam
                    self._seen_item_ids.add(item_id)

            except Exception as exc:
                logger.error("MIDP poll error: %s", exc, exc_info=True)

            await asyncio.sleep(POLL_INTERVAL_SECONDS)
