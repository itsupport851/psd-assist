# PSD Assist

A Flask-based service for:
- Filling a Dropbox-hosted GPC Commercial Workbook template with customer data
- Exposing HubSpot CRM utility endpoints for deals, line items, contacts, teams, and service records
- Serving a static React-like portal UI for sales, operations, and service workflows
- Hosting a customer-facing pre-installation intake form that creates HubSpot records
- Proxying webhook and Claude requests for external integrations

## Main features
- `POST /fill-workbook`: downloads a workbook template from Dropbox, populates it with customer fields, and returns a macro-enabled XLSM file
- `GET /health`: quick readiness check
- HubSpot CRM object endpoints for:
  - deals, line items, contacts, teams
  - service creation, updates, pipelines, and debug inspection
- Static frontends served at `/`, `/sales`, `/operations`, `/service`, `/map`, and `/customer-intake`
- PIN-based role authentication with admin, sales, analyst, operations, and installer roles
- A separate PIN-gated customer intake form that provisions a contact, deal, appointment, and line items in HubSpot

## Environment variables
Required values are typically set in Railway or your deployment environment.

- `DROPBOX_REFRESH_TOKEN`: refresh token for Dropbox API access
- `DROPBOX_APP_KEY`: Dropbox app key
- `DROPBOX_APP_SECRET`: Dropbox app secret
- `TEMPLATE_PATH`: path to workbook template in Dropbox
- `HUBSPOT_API_KEY`: HubSpot private app access token
- `SERVICE_PIPELINE_ID`: HubSpot service pipeline ID to use for service creation
- `ANTHROPIC_API_KEY`: API key used by the `/claude` proxy endpoint
- `GOOGLE_MAPS_API_KEY`: optional for map page support
- `PORTAL_ID`: HubSpot portal ID
- `OWNER_ID`: HubSpot owner ID assigned to deals created by the intake form
- `UNIT_PRICE`: default unit price applied to line items created from fan configuration rows
- `PIN_ADMIN`: admin PIN for portal access
- `PIN_SALES`: sales PIN
- `PIN_ANALYST`: analyst PIN for portal access
- `PIN_OPERATIONS`: operations PIN
- `PIN_INSTALLER_{team_id}`: installer PIN for each HubSpot team ID

### Customer intake form
- `CUSTOMER_INTAKE_PIN`: PIN customers enter to unlock `/customer-intake`. If unset, the form is not PIN-gated
- `DEALSTAGE_CUSTOMER_FORM_SENT`: deal stage applied to deals created from an intake submission (default `appointmentscheduled`)
- `HUBSPOT_FILES_URL`: HubSpot Files API endpoint for customer document uploads (default `https://api.hubapi.com/files/v3/files`). The Files API is versioned `v3`; a dated segment such as `/files/2026-03/files` appears in some doc samples as a placeholder and returns 404
- `INTAKE_FILES_FOLDER`: root file-manager folder for uploaded documents (default `/customer-intake`); each submission gets a subfolder named after the account

Intake deals are always created on the `default` deal pipeline, regardless of power company.

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
- `GET /hubspot/team-services/<team_id>`
- `GET /hubspot/service-details/<service_id>`
- `GET /hubspot/service-stages`
- `GET /hubspot/service-debug/<service_id>`
- `GET /hubspot/service-properties`
- `GET /hubspot/all-services`
- `PATCH /hubspot/services/<service_id>/repair`
- `GET /hubspot/service-inspect/<service_id>`
- `GET /hubspot/services/<deal_id>`
- `POST /hubspot/services`
- `PATCH /hubspot/services/<service_id>`

### Customer intake endpoints
- `POST /auth`
  - Validates a portal PIN and returns the matching role
- `POST /customer-intake-auth`
  - Validates `CUSTOMER_INTAKE_PIN` before the form is unlocked
- `POST /submit-customer-form`
  - Accepts the completed intake payload plus `access_pin`
  - Requires `gp_account_name`, `gp_account_number`, `first_name`, `last_name`, `email`, `farm_address`, and `power_company`
  - Returns `contact_id` and `deal_id`
- `POST /customer-intake-upload`
  - `multipart/form-data` with `deal_id`, `access_pin`, optional `account_name`, and one or more `files` parts
  - Uploads each file to HubSpot Files as `PRIVATE`, then attaches them all to the deal through a single note
  - Returns `uploaded` and `failed` filename lists; a partial failure is reported rather than raised

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
- `/customer-intake` — customer-facing pre-installation intake form
- `/maps-key` — returns `GOOGLE_MAPS_API_KEY`
- `/Images/<path>` and `/Docs/<path>` — static assets referenced by the portals and intake form

## Customer intake form

`static/customer-intake.html` is a self-contained, PIN-gated survey sent to customers before installation. It collects farm and contact details, poultry operation and fan inventory, an equipment readiness checklist, and an installation scheduling window across four steps.

On submit, `POST /submit-customer-form` performs the following against HubSpot:
1. Creates the contact, or patches it if one already exists with the same email
2. Creates a deal on the `default` pipeline at `DEALSTAGE_CUSTOMER_FORM_SENT`, and associates it with the contact
3. Best-effort: creates an appointment spanning the installation window and associates it with the deal
4. Best-effort: converts each fan configuration row into a line item priced at `UNIT_PRICE`, then updates the deal amount
5. Best-effort: writes the full checklist as a note on the deal and contact

Steps 3–5 are best-effort — a failure there does not fail the submission.

### Power company and state
The State field is derived from the selected power company and is not editable:

| Power company | State |
| --- | --- |
| Georgia Power | GA |
| Entergy Louisiana | LA |
| Entergy Arkansas | A |

Entergy Arkansas maps to `A`, not the USPS code `AR`, as specified by the business owner.

The form renders State read-only, and `POWER_COMPANY_STATE_MAP` in `app.py` re-derives it server-side so a hand-edited payload cannot override it. The portal's New Contact form (`index.html`) carries its own copy in `POWER_COMPANIES` and derives State the same way, though it does not send `power_company` to HubSpot. Adding a power company means updating all three lists.

### Document uploads
Step 4 accepts up to 10 optional files of 20 MB each. They are sent only after `/submit-customer-form` returns a `deal_id`, in a second `multipart/form-data` request to `/customer-intake-upload`. Each file is posted to the HubSpot Files API with `{"access": "PRIVATE"}` under `INTAKE_FILES_FOLDER/<account name>`, and the resulting file ids are attached to the deal as one note:

```json
{
  "associations": [
    { "to": { "id": "<deal_id>" },
      "types": [{ "associationCategory": "HUBSPOT_DEFINED", "associationTypeId": 214 }] }
  ],
  "properties": {
    "hs_note_body": "Customer-uploaded documents…",
    "hs_timestamp": "2026-09-01T15:48:22Z",
    "hs_attachment_ids": "<fileId>;<fileId>"
  }
}
```

Association type `214` is HubSpot's `note → deal` type. (`202` is `note → contact` and will not attach a note to a deal.)

Because uploads happen after the deal exists, an upload failure never discards a successful submission — the success screen shows a warning naming the files that did not attach and why. Each failure is also logged server-side with HubSpot's own message, so the two common causes are distinguishable: a **404** means `HUBSPOT_FILES_URL` is wrong, and a **403** means the private app token is missing the `files` scope.

## Theme

All three static pages use a single light palette. Colors are literal hex values in each page's `<style>` block and in the inline React style objects in `index.html` — there are no CSS variables, so a palette change means a sweep across all three files.

| Role | Value |
| --- | --- |
| Page background | `#f4f6f9` |
| Card / input / chrome surface | `#ffffff` |
| Inset or read-only surface | `#eef2f6` |
| Border | `#dce3ea` (dashed `#c3cedb`) |
| Primary text | `#16212c` |
| Label / secondary text | `#3d6076` |
| Muted text | `#64798c` |
| Brand green (fills, borders) | `#25a35a` / `#1a6b3c` |
| Green text on a light tint | `#1a6b3c` |
| Green tint background | `#e9f7ef` |
| Error | `#c62828` |
| Warning | `#a86a00` |

Two rules matter when editing:

- `color: #fff` is only ever correct on a **saturated fill** — the green gradient buttons, a selected Yes/No pill, a stage pill. On a light tint it is invisible.
- `#25a35a` is a fill and border color. As *text* on `#e9f7ef` it lands at 2.94:1, so use `#1a6b3c` there instead.

Map marker `stroke` values in `map.html` are deliberately left dark; they outline saturated markers against the light Google basemap. The floating map panels carry a `box-shadow` because a border alone no longer separates a white panel from a light map.

### Frontend notes
The page renders itself by assigning to `#app.innerHTML`. **Do not call `render()` from an input handler.** Doing so destroys and recreates the focused element on every keystroke, which makes `<input type="date">` unusable (the picker closes and segment focus resets) and causes visible flicker. Instead:

- `update()` records state only; the browser already reflects what the user typed
- `clearErrorText()` removes the error banner in place rather than re-rendering
- Handlers that reformat as you type (`updatePhone`, `updateDigits`) write the value back to the live input via `setInputValue()` and restore the caret
- `showError()` / `clearErrorText()` patch the error banner node directly
- `toggleReadiness()` flips a class on the button it was handed
- `render()` is reserved for structural changes: step navigation, adding/removing fan rows, submit, and PIN verification

Selected upload files live in the module-level `selectedFiles` array, deliberately outside `state`. A file input's selection cannot be restored programmatically, so a full `render()` would silently drop the user's choices; keeping the `File` objects outside the rendered DOM makes re-rendering safe. It also keeps `JSON.stringify(state)` on submit from choking on them. `addFiles`/`removeFile` repaint only the `#file-list` container via `renderFileList()`.

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
