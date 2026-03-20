# MIDPAgent – Azure Deployment Guide

Everything needed to deploy (or redeploy) the MIDPAgent Teams bot lives in
this folder. A single command handles infrastructure provisioning **and**
application code deployment.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| **Azure CLI** | `az --version` ≥ 2.50. Must be logged in (`az login`). |
| **PowerShell** | 5.1+ (Windows built-in) or PowerShell 7+ (cross-platform). |
| **Azure Subscription** | With a resource group already created. |
| **Entra App Registration** | `BOT_APP_ID` + client secret. Must have API permission **Teamwork.Migrate.All** and a **BotService/botServices** resource linked to it. |
| **Teams Channel** | The bot's Teams app manifest must be sideloaded into the target team. See `TeamsCommunication/teams_app/`. |

---

## Folder Contents

| File | Purpose |
|---|---|
| `deploy.ps1` | **Run this.** Single-invocation PowerShell deploy script. |
| `main.bicep` | Bicep IaC template (App Service Plan + Web App + Bot endpoint). |
| `main.parameters.json` | Pre-filled parameter values. Only `botAppPassword` is left empty (passed at deploy time). |
| `Dockerfile` | Optional – for containerised builds (ACR / Container Apps). |
| `.dockerignore` | Docker build-context exclusions. |
| `README.md` | This file. |

---

## Quick Start – Deploy in One Command

```powershell
cd AzureDeploy

.\deploy.ps1 `
    -ResourceGroup  "rg-inodigitaldemo-documentmanager" `
    -BotAppPassword "<your-client-secret>"
```

That's it. The script:

1. **Provisions infrastructure** – runs `az deployment group create` with the
   Bicep template. Creates (or updates) the App Service Plan, Web App, and Bot
   endpoint.
2. **Packages the application** – copies `TeamsCommunication/`, `requirements.txt`,
   and `env` from the repo root into a zip file (`deploy-package.zip`).
3. **Deploys the zip** – pushes the package to the Azure Web App via
   `az webapp deploy`.

> **Tip:** You can also invoke the script from the repo root:
> ```powershell
> .\AzureDeploy\deploy.ps1 -ResourceGroup "rg-inodigitaldemo-documentmanager" -BotAppPassword "<secret>"
> ```
> All paths resolve relative to the script's own location.

---

## What Gets Deployed

```
Azure
 ├─ App Service Plan   (Linux B1)
 ├─ Web App            (Python 3.12, always-on)
 │    ├─ TeamsCommunication/    ← bot code
 │    ├─ requirements.txt       ← pip dependencies (installed via SCM build)
 │    └─ env                    ← environment variables file
 └─ Bot Service        (endpoint updated to → https://<webapp>/api/messages)
```

### App Settings (set by Bicep)

| Setting | Source |
|---|---|
| `BOT_APP_ID` | Entra app registration |
| `BOT_APP_PASSWORD` | Client secret (secure param) |
| `AZURE_TENANT_ID` | Entra tenant |
| `AZURE_AI_PROJECT_ENDPOINT` | AI Foundry project URL |
| `AGENT_NAME` | Foundry agent / assistant name |
| `TEAMS_CHANNEL_ID` | Target Teams channel |
| `TEAMS_SERVICE_URL` | Bot Framework service URL |
| `BOT_PORT` / `WEBSITES_PORT` | `8000` |
| `SCM_DO_BUILD_DURING_DEPLOYMENT` | `true` – runs `pip install` on deploy |

---

## Redeploying After Code Changes

The exact same command works for redeployments – Bicep is idempotent:

```powershell
.\deploy.ps1 -ResourceGroup "rg-inodigitaldemo-documentmanager" -BotAppPassword "<secret>"
```

If you **only** changed Python code (no infra changes), you can skip the Bicep
step and just zip-deploy manually:

```powershell
# From repo root
Compress-Archive -Path TeamsCommunication, requirements.txt, env -DestinationPath pkg.zip -Force
az webapp deploy --resource-group rg-inodigitaldemo-documentmanager --name app-midpagent-<suffix> --src-path pkg.zip --type zip
Remove-Item pkg.zip
```

---

## Docker Build (Optional)

The Dockerfile is provided for containerised scenarios (e.g. Azure Container
Apps or local testing). Build from the **repo root** so that the application
sources are in the build context:

```powershell
docker build -f AzureDeploy/Dockerfile -t midpagent:latest .
docker run -p 8000:8000 --env-file env midpagent:latest
```

---

## Monitoring & Logs

```powershell
# Stream live logs
az webapp log tail --resource-group <rg> --name <webapp>

# Health check
curl https://<webapp>.azurewebsites.net/

# Check bot endpoint is registered
az bot show --resource-group <rg> --name MIDPAgent --query "properties.endpoint"
```

---

## Customising Parameters

Edit `main.parameters.json` to change default values. The only parameter you
**must** supply at the command line is `BotAppPassword` (it should never be
committed to source control).

To use a different SKU, region, or naming convention, edit `main.bicep`
directly.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `503 Service Unavailable` after deploy | Container still starting. Wait 1–2 min, then check `az webapp log tail`. |
| `401 Unauthorized` on `POST /api/messages` | Expected – Bot Framework authenticates via bearer token; raw POSTs are rejected. |
| `Bicep deployment failed` | Ensure the resource group exists and you have Contributor role. |
| Bot doesn't reply in Teams | Check `az webapp log tail` for errors. Verify the Foundry agent endpoint and credentials. |
