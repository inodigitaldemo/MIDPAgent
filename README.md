# MIDPAgent
An AI Agent that listens to a master information delivery plan in a SharePoint list and works on tasks in it.

## Configuration

This project now reads settings from `config.json` (not environment variables).

1. Copy `config.example.json` to `config.json`.
2. Fill in:
	- `sharepoint.site_url`
	- `sharepoint.list_name`
	- `sharepoint.client_id`
	- `sharepoint.client_secret`
	- `azure.ai_project_endpoint`

SharePoint authentication uses **app credentials only** (`client_id` + `client_secret`).
