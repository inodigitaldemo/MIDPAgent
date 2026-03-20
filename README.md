# MIDPAgent

An AI Agent that listens to a master information delivery plan in a SharePoint list and works on tasks in it.

This repository also includes a Markdown → HTML → DOCX conversion toolchain with branded Word headers/footers.

## Configuration

This project reads settings from `config.json` (not environment variables).

1. Copy `config.example.json` to `config.json`.
2. Fill in:
   - `sharepoint.site_url`
   - `sharepoint.list_name`
   - `sharepoint.client_id`
   - `sharepoint.client_secret`
   - `azure.ai_project_endpoint`

SharePoint authentication uses **app credentials only** (`client_id` + `client_secret`).

## TeamsCommunication (Bot Framework)

A Teams integration module built on the **Bot Framework SDK** (`botbuilder-python`).
It authenticates as a registered Azure Bot and can:

- **Receive messages** – runs an aiohttp web server at `/api/messages`
- **Forward to AI** – incoming user messages are sent to the Azure AI Foundry agent; the agent's reply is returned as an Adaptive Card
- **Send proactive messages** – posts Adaptive Cards to a Teams channel from the CLI

### Azure Bot resource

```
/subscriptions/67ab6b87-3bb5-4202-8f73-f1d846582918/resourceGroups/rg-inodigitaldemo-documentmanager/providers/Microsoft.BotService/botServices/MIDPAgent
```

### Environment variables (add to `env`)

| Variable | Description |
|---|---|
| `BOT_APP_ID` | Microsoft App ID from the bot's app registration |
| `BOT_APP_PASSWORD` | Client secret for that app registration |
| `AZURE_TENANT_ID` | Azure AD / Entra tenant ID |
| `AZURE_AI_PROJECT_ENDPOINT` | Azure AI Foundry project endpoint URL |
| `AGENT_NAME` | Name of the Foundry agent / assistant (e.g. `Agent943`) |
| `TEAMS_CHANNEL_ID` | Channel ID to post proactive messages to |
| `TEAMS_SERVICE_URL` | Bot Framework service URL (default: `https://smba.trafficmanager.net/emea/`) |
| `BOT_PORT` | Web server port (default: `3978`) |

### Install dependencies

```powershell
pip install -r requirements.txt
```

### Run the bot locally

```powershell
python -m TeamsCommunication.app
```

This starts an aiohttp server on port 3978.  Point the Azure Bot resource's
**Messaging endpoint** to `https://<your-host>/api/messages`.

For local development, use [ngrok](https://ngrok.com) or the
[Dev Tunnels](https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/) VS Code extension
to expose the local port.

### Send a proactive Hello World card

```powershell
python -m TeamsCommunication.console_app
```

Requires `BOT_APP_ID`, `BOT_APP_PASSWORD`, and `TEAMS_CHANNEL_ID` in `env`.

### Module structure

| File | Purpose |
|---|---|
| `TeamsCommunication/app.py` | aiohttp web server (Bot Framework endpoint) |
| `TeamsCommunication/bot.py` | `TeamsActivityHandler` subclass – forwards messages to AI agent |
| `TeamsCommunication/agent_service.py` | Async Azure AI Foundry Assistants API client |
| `TeamsCommunication/config.py` | Reads env vars into a `BotConfig` dataclass |
| `TeamsCommunication/adaptive_cards.py` | Adaptive Card builders (hello world, agent response, error) |
| `TeamsCommunication/proactive.py` | Proactive channel messaging helper |
| `TeamsCommunication/console_app.py` | CLI entry point for proactive send |

---

## Deploying to Azure

All deployment files live in the **`AzureDeploy/`** folder — Bicep template,
parameter file, Dockerfile, and a single-command PowerShell deploy script.

```powershell
.\AzureDeploy\deploy.ps1 `
    -ResourceGroup  "rg-inodigitaldemo-documentmanager" `
    -BotAppPassword "<your-client-secret>"
```

See [AzureDeploy/README.md](AzureDeploy/README.md) for full details,
prerequisites, Docker instructions, and troubleshooting.

### Architecture overview

```
Teams user
    │
    ▼
Microsoft Teams ──► Azure Bot Service (channel routing)
                        │
                        ▼
                    App Service (aiohttp on port 8000)
                    POST /api/messages
                        │
                        ├─► MIDPBot.on_message_activity()
                        │       │
                        │       ▼
                        │   FoundryAgentService.send_message()
                        │       │
                        │       ▼
                        │   Azure AI Foundry (Assistants API)
                        │       │
                        │       ▼
                        │   Agent reply ──► Adaptive Card ──► Teams
                        │
                        └─► Proactive: send_to_channel()
                                │
                                ▼
                            Teams channel
```

### Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Bot replies with hello-world card only | `AZURE_AI_PROJECT_ENDPOINT` or `AGENT_NAME` not set | Add both env vars to the App Service configuration |
| 401 on token acquisition | App registration doesn't have `Cognitive Services User` role on the AI Foundry resource | Assign the role in the Azure portal → AI resource → IAM |
| Bot not responding in Teams | Messaging endpoint mismatch | Verify the Azure Bot resource's endpoint matches the Web App URL |
| `BotNotInConversationRoster` | Bot not installed in the Team | Sideload the Teams app from `TeamsCommunication/teams_app/` |

## Markdown to DOCX Tooling

### What it does

The conversion flow is:

1. Markdown (`.md`) → standalone HTML (`.html`) using Pandoc
2. HTML → Word (`.docx`) using Pandoc
3. Post-process DOCX to add **real Word header/footer content** (not just body HTML), including page numbers

### Key files

- `md_to_docx/console_app.py`: command-line entry point
- `md_to_docx/converter.py`: conversion and DOCX branding logic
- `DocTemplates/HTML/template.html`: HTML template used during Markdown → HTML
- `DocTemplates/CSS/style.css`: shared CSS and branding color variables

### Install dependencies

Use your virtual environment and install requirements:

```powershell
pip install -r requirements.txt
```

`pypandoc_binary` is used so Pandoc is bundled with the Python package.

### Usage

Run the converter as a module:

```powershell
python -m md_to_docx.console_app <input.md> [--output-dir DIR] [--css PATH]
```

Examples:

```powershell
# Convert a markdown file in place
python -m md_to_docx.console_app MDFiles/ExampleDocument.md

# Write outputs to a custom folder
python -m md_to_docx.console_app docs/spec.md --output-dir dist/

# Use a custom stylesheet
python -m md_to_docx.console_app report.md --css DocTemplates/CSS/style.css
```

### Output

For an input file named `MyDoc.md`, the tool generates:

- `MyDoc.html` (intermediate)
- `MyDoc.docx` (final output)

By default, output is written to the same folder as the input Markdown unless `--output-dir` is provided.

### Branding behavior

The DOCX branding step applies:

- A header band with title + subtitle
- A footer band with left/right text
- Dynamic page numbering (`Page X`) via native Word field codes

Brand colors are read from CSS custom properties in `DocTemplates/CSS/style.css`:

- `--document-header-bg`
- `--document-footer-bg`
- `--document-header-text`
- `--document-footer-text`

Current defaults are black backgrounds (`#000000`) with text in `#ffb81c`.
