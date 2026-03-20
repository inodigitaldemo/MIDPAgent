// ---------------------------------------------------------------------------
// MIDPAgent Bot – Azure infrastructure (Bicep)
//
// Deploys:
//   1. App Service Plan (Linux, B1)
//   2. Web App (Python 3.12) with all required app settings
//   3. Updates the existing Azure Bot resource's messaging endpoint
//
// Invoked automatically by deploy.ps1 – see AzureDeploy/README.md.
// ---------------------------------------------------------------------------

targetScope = 'resourceGroup'

// ── Parameters ──────────────────────────────────────────────────────────────

@description('Microsoft App ID of the bot (from Entra app registration)')
param botAppId string

@secure()
@description('Client secret for the bot app registration')
param botAppPassword string

@description('Azure AD / Entra tenant ID')
param tenantId string

@description('Azure AI Foundry project endpoint URL')
param aiProjectEndpoint string

@description('Name of the Foundry agent / assistant')
param agentName string

@description('Teams channel ID for proactive messages (optional)')
param teamsChannelId string = ''

@description('Bot Framework service URL (default EMEA)')
param teamsServiceUrl string = 'https://smba.trafficmanager.net/emea/'

@description('Azure region for all resources')
param location string = resourceGroup().location

@description('Unique suffix for resource names (default: deterministic hash of resource group)')
param nameSuffix string = uniqueString(resourceGroup().id)

@description('Name of the existing Azure Bot resource')
param botResourceName string = 'MIDPAgent'

// ── Variables ───────────────────────────────────────────────────────────────

var appServicePlanName = 'plan-midpagent-${nameSuffix}'
var webAppName = 'app-midpagent-${nameSuffix}'

// ── App Service Plan ────────────────────────────────────────────────────────

resource appServicePlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: appServicePlanName
  location: location
  kind: 'linux'
  sku: {
    name: 'B1'
    tier: 'Basic'
  }
  properties: {
    reserved: true // Linux
  }
}

// ── Web App ─────────────────────────────────────────────────────────────────

resource webApp 'Microsoft.Web/sites@2023-12-01' = {
  name: webAppName
  location: location
  kind: 'app,linux'
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    siteConfig: {
      linuxFxVersion: 'PYTHON|3.12'
      appCommandLine: 'python -m TeamsCommunication.app'
      alwaysOn: true
      appSettings: [
        { name: 'BOT_APP_ID', value: botAppId }
        { name: 'BOT_APP_PASSWORD', value: botAppPassword }
        { name: 'AZURE_TENANT_ID', value: tenantId }
        { name: 'AZURE_AI_PROJECT_ENDPOINT', value: aiProjectEndpoint }
        { name: 'AGENT_NAME', value: agentName }
        { name: 'TEAMS_CHANNEL_ID', value: teamsChannelId }
        { name: 'TEAMS_SERVICE_URL', value: teamsServiceUrl }
        { name: 'BOT_PORT', value: '8000' }
        { name: 'WEBSITES_PORT', value: '8000' }
        { name: 'SCM_DO_BUILD_DURING_DEPLOYMENT', value: 'true' }
      ]
    }
  }
}

// ── Update Bot messaging endpoint ───────────────────────────────────────────

resource bot 'Microsoft.BotService/botServices@2022-09-15' = {
  name: botResourceName
  location: 'global'
  kind: 'azurebot'
  sku: {
    name: 'F0'
  }
  properties: {
    displayName: botResourceName
    endpoint: 'https://${webApp.properties.defaultHostName}/api/messages'
    msaAppId: botAppId
    msaAppTenantId: tenantId
    msaAppType: 'SingleTenant'
  }
}

// ── Outputs ─────────────────────────────────────────────────────────────────

output webAppName string = webApp.name
output webAppUrl string = 'https://${webApp.properties.defaultHostName}'
output messagingEndpoint string = 'https://${webApp.properties.defaultHostName}/api/messages'
