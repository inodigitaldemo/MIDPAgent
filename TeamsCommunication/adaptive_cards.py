"""
Adaptive Card builders for TeamsCommunication.

- ``produce_document_card`` / ``produce_document_attachment`` – "Should I produce
  this document?" yes/no prompt posted when the MIDP poller finds a new item.
- ``document_approval_card`` / ``document_approval_attachment`` – "Approve this
  document?" card posted after a document has been generated and uploaded.
- ``error_card`` – error notification card.
"""

from __future__ import annotations

from botbuilder.core import CardFactory
from botbuilder.schema import Attachment


# ── Produce-document prompt ──────────────────────────────────────────────────

def produce_document_card(
    title: str,
    item_id: str,
    dokumentnummer: str = "",
    dokumenttype: str = "",
    disiplin: str = "",
) -> dict:
    """Adaptive Card (Norwegian) asking whether to produce a planned document."""
    facts = [
        {"title": "Tittel", "value": title or "–"},
        {"title": "Dokumentnummer", "value": dokumentnummer or "–"},
        {"title": "Dokumenttype", "value": dokumenttype or "–"},
        {"title": "Disiplin", "value": disiplin or "–"},
    ]
    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": [
            {
                "type": "TextBlock",
                "text": "📄 Nytt MIDP-element oppdaget",
                "weight": "Bolder",
                "size": "Medium",
                "color": "Accent",
                "wrap": True,
            },
            {
                "type": "FactSet",
                "facts": facts,
            },
            {
                "type": "TextBlock",
                "text": "Skal jeg produsere dette dokumentet?",
                "wrap": True,
            },
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Ja – Produser",
                "style": "positive",
                "data": {
                    "action": "produce_document",
                    "choice": "yes",
                    "item_id": item_id,
                    "title": title,
                },
            },
            {
                "type": "Action.Submit",
                "title": "Nei – Hopp over",
                "style": "destructive",
                "data": {
                    "action": "produce_document",
                    "choice": "no",
                    "item_id": item_id,
                    "title": title,
                },
            },
        ],
    }


def produce_document_attachment(
    title: str,
    item_id: str,
    dokumentnummer: str = "",
    dokumenttype: str = "",
    disiplin: str = "",
) -> Attachment:
    """Bot Framework Attachment wrapping the produce-document card."""
    return CardFactory.adaptive_card(
        produce_document_card(
            title, item_id, dokumentnummer, dokumenttype, disiplin
        )
    )


# ── Document approval ────────────────────────────────────────────────────────

def document_approval_card(
    title: str,
    item_id: str,
    doc_url: str,
    filename: str = "",
) -> dict:
    """Adaptive Card asking the user to approve or reject a generated document."""
    body = [
        {
            "type": "TextBlock",
            "text": "✅ Dokument klart for godkjenning",
            "weight": "Bolder",
            "size": "Medium",
            "color": "Good",
            "wrap": True,
        },
        {
            "type": "TextBlock",
            "text": f"**{title}**",
            "wrap": True,
        },
    ]
    if filename:
        body.append(
            {
                "type": "TextBlock",
                "text": f"Fil: {filename}",
                "isSubtle": True,
                "wrap": True,
            }
        )
    if doc_url:
        body.append(
            {
                "type": "TextBlock",
                "text": f"[Åpne dokument]({doc_url})",
                "wrap": True,
            }
        )
    body.append(
        {
            "type": "TextBlock",
            "text": "Godkjenner du dette dokumentet?",
            "wrap": True,
        }
    )

    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": body,
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Godkjenn",
                "style": "positive",
                "data": {
                    "action": "approve_document",
                    "choice": "yes",
                    "item_id": item_id,
                    "title": title,
                },
            },
            {
                "type": "Action.Submit",
                "title": "Avvis",
                "style": "destructive",
                "data": {
                    "action": "approve_document",
                    "choice": "no",
                    "item_id": item_id,
                    "title": title,
                },
            },
        ],
    }


def document_approval_attachment(
    title: str,
    item_id: str,
    doc_url: str,
    filename: str = "",
) -> Attachment:
    """Bot Framework Attachment wrapping the document-approval card."""
    return CardFactory.adaptive_card(
        document_approval_card(title, item_id, doc_url, filename)
    )


# ── Error notification ────────────────────────────────────────────────────────

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
                    "text": "⚠ Noe gikk galt",
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
