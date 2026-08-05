# PSD Workbook Webhook

A Flask-based webhook service for:
- Filling a Dropbox-hosted GPC Commercial Workbook template with customer data
- Exposing HubSpot CRM utility endpoints for deals, line items, contacts, teams, and service records
- Serving a static React-like portal UI for sales, operations, and service workflows
- Proxying webhook and Claude requests for external integrations

## Main features
- `POST /fill-workbook`: downloads a workbook template from Dropbox, populates it with customer fields, and returns a macro-enabled XLSM file
- `GET /health`: quick readiness check
- HubSpot CRM object endpoints for:
  - deals, line items, contacts, teams
  - service creation, updates, pipelines, and debug inspection
- Static frontends served at `/`, `/sales`, `/operations`, `/service`, and `/map`
- PIN-based role authentication with admin, sales, analyst, operations, and installer roles

## Environment variables
Required values are typically set in Railway or your deployment environment.

- `DROPBOX_REFRESH_TOKEN`: refresh token for Dropbox API access
- `DROPBOX_APP_KEY`: Dropbox app key
- `DROPBOX_APP_SECRET`: Dropbox app secret
- `TEMPLATE_PATH`: path to workbook template in Dropbox
- `HUBSPOT_API_KEY`: HubSpot private app access token
- `SERVICE_PIPELINE_ID`: HubSpot service pipeline ID to use for service creation
- `GOOGLE_MAPS_API_KEY`: optional for map page support
- `PIN_ADMIN`: admin PIN for portal access
- `PIN_SALES`: sales PIN
- `PIN_ANALYST`: analyst PIN for portal access
- `PIN_OPERATIONS`: operations PIN
- `PIN_INSTALLER_{team_id}`: installer PIN for each HubSpot team ID

## App routes

### Workbook endpoints
- `GET /health`
  - Returns `{ "status": "ok" }`
- `POST /fill-workbook`
  - Accepts JSON customer data
  - Populates the workbook template in Dropbox
  - Returns a downloadable `.xlsm` file

Example payload for `/fill-workbook`:
```json
{
  "gp_account_name": "Smith Farms",
  "gp_account_number": "1234567890",
  "owner_name": "John Smith",
  "phone": "555-123-4567",
  "email": "john@smithfarms.com",
  "farm_address": "123 Farm Road",
  "city": "Atlanta",
  "state": "GA",
  "zip": "30301"
}
```

### HubSpot utility endpoints
- `GET /hubspot/deals`
- `GET /hubspot/line-items/<deal_id>`
- `GET /hubspot/deals-with-addresses`
- `PATCH /hubspot/line-items/<line_item_id>`
- `PATCH /hubspot/deals/<deal_id>`
- `POST /hubspot/line-items`
- `POST /hubspot/deals`
- `POST /hubspot/contacts`
- `GET /hubspot/teams`
- `GET /hubspot/service-stages`
- `GET /hubspot/service-debug/<service_id>`
- `GET /hubspot/service-properties`
- `GET /hubspot/all-services`
- `PATCH /hubspot/services/<service_id>/repair`
- `GET /hubspot/service-inspect/<service_id>`
- `GET /hubspot/services/<deal_id>`
- `POST /hubspot/services`
- `PATCH /hubspot/services/<service_id>`

### Proxy endpoints
- `POST /proxy`
  - Forward JSON payloads to arbitrary HTTP endpoints
- `POST /claude`
  - Proxy requests to Anthropic Claude API

## Static frontend routes
- `/` — main portal entrypoint
- `/sales` — sales portal
- `/operations` — operations portal
- `/service` — service portal
- `/map` — map page
- `/maps-key` — returns `GOOGLE_MAPS_API_KEY`

## HubSpot service behavior
- Service creation and update operations normalize `status` values to HubSpot-approved options
- Accepted HubSpot `hs_status` values include: `on_track`, `delayed`, `failed`, `succeeded_completed`
- Service stage selection is mapped to the configured HubSpot service pipeline

## Deploy to Railway
1. Push this repo to GitHub.
2. Create a Railway project and connect the GitHub repository.
3. Configure the environment variables in Railway.
4. Deploy the project.
5. Railway will provide a public URL for the service.

## Tips
- Ensure Dropbox credentials are valid and the workbook path exists.
- Ensure the HubSpot API key has permissions for CRM objects and pipelines.
- Use the `/hubspot/service-properties` and `/hubspot/service-debug/<service_id>` endpoints to inspect available HubSpot service fields.
- If you need a raw service lookup, use `/hubspot/service-inspect/<service_id>`.

## License
This repository is provided as-is for internal integration and automation use.
