<#
.SYNOPSIS
    One-command deploy of the MIDPAgent Teams bot to Azure.

.DESCRIPTION
    Deploys infrastructure (Bicep) and application code (zip deploy) to Azure
    App Service. Run from anywhere – all paths resolve relative to this script.

    Steps performed:
      1. Provision / update infrastructure via Bicep template
      2. Package application code into a zip
      3. Deploy the zip to the Azure Web App

.PARAMETER ResourceGroup
    Name of the Azure resource group (must already exist).

.PARAMETER BotAppPassword
    Client secret for the bot's Entra app registration.
    Will also be passed into the Bicep template as a secure parameter.

.EXAMPLE
    .\deploy.ps1 -ResourceGroup rg-inodigitaldemo-documentmanager -BotAppPassword "<secret>"

.NOTES
    Prerequisites:
      - Azure CLI installed and logged in (az login)
      - PowerShell 5.1+ or PowerShell 7+
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ResourceGroup,

    [Parameter(Mandatory)]
    [string]$BotAppPassword
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Resolve paths relative to this script ───────────────────────────────────

$ScriptDir  = $PSScriptRoot                          # AzureDeploy/
$RepoRoot   = Split-Path $ScriptDir -Parent          # project root

$BicepFile  = Join-Path $ScriptDir 'main.bicep'
$ParamsFile = Join-Path $ScriptDir 'main.parameters.json'

# Directories / files that make up the deployable application
$AppSources = @(
    (Join-Path $RepoRoot 'TeamsCommunication'),
    (Join-Path $RepoRoot 'requirements.txt'),
    (Join-Path $RepoRoot 'env')
)

# ── Helpers ──────────────────────────────────────────────────────────────────

function Write-Step([string]$Message) {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  $Message" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}

# ── Step 1 – Infrastructure ─────────────────────────────────────────────────

Write-Step 'Step 1/3: Deploying Azure infrastructure (Bicep)'

$deployment = az deployment group create `
    --resource-group $ResourceGroup `
    --template-file $BicepFile `
    --parameters "@$ParamsFile" `
    --parameters botAppPassword=$BotAppPassword `
    --output json | ConvertFrom-Json

if ($LASTEXITCODE -ne 0) {
    Write-Error 'Bicep deployment failed – aborting.'
    exit 1
}

$webAppName      = $deployment.properties.outputs.webAppName.value
$webAppUrl       = $deployment.properties.outputs.webAppUrl.value
$messagingEndpoint = $deployment.properties.outputs.messagingEndpoint.value

Write-Host "  Web App      : $webAppName"
Write-Host "  URL          : $webAppUrl"
Write-Host "  Bot Endpoint : $messagingEndpoint"

# ── Step 2 – Package ────────────────────────────────────────────────────────

Write-Step 'Step 2/3: Packaging application code'

$ZipPath = Join-Path $ScriptDir 'deploy-package.zip'

if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }

# Stage into a temp folder so the zip has a flat structure at the root
$StageDir = Join-Path ([System.IO.Path]::GetTempPath()) "midpagent-stage-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $StageDir -Force | Out-Null

foreach ($src in $AppSources) {
    if (-not (Test-Path $src)) {
        Write-Error "Source not found: $src"
        exit 1
    }
    $dest = Join-Path $StageDir (Split-Path $src -Leaf)
    if (Test-Path $src -PathType Container) {
        Copy-Item $src $dest -Recurse -Force
    } else {
        Copy-Item $src $dest -Force
    }
}

Write-Host "  Staged to: $StageDir"

# Create zip
Push-Location $StageDir
Compress-Archive -Path (Get-ChildItem -Path $StageDir) -DestinationPath $ZipPath -Force
Pop-Location

# Clean up staging area
Remove-Item $StageDir -Recurse -Force

$zipSizeMB = [math]::Round((Get-Item $ZipPath).Length / 1MB, 2)
Write-Host "  Package: $ZipPath ($zipSizeMB MB)"

# ── Step 3 – Deploy ─────────────────────────────────────────────────────────

Write-Step 'Step 3/3: Deploying code to Azure Web App'

az webapp deploy `
    --resource-group $ResourceGroup `
    --name $webAppName `
    --src-path $ZipPath `
    --type zip `
    --async false

if ($LASTEXITCODE -ne 0) {
    Write-Error 'Code deployment failed.'
    exit 1
}

# ── Done ─────────────────────────────────────────────────────────────────────

Write-Host ''
Write-Step 'Deployment complete!'
Write-Host "  Health check : $webAppUrl"
Write-Host "  Bot endpoint : $messagingEndpoint"
Write-Host ''
Write-Host 'Tip: stream logs with:' -ForegroundColor Yellow
Write-Host "  az webapp log tail --resource-group $ResourceGroup --name $webAppName" -ForegroundColor Yellow
Write-Host ''
