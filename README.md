# PSD Workbook Filler Webhook

Simple Flask API that fills GPC Commercial Workbook with customer data.

## Environment Variables (set in Railway dashboard)
- `DROPBOX_TOKEN` — your Dropbox API access token
- `TEMPLATE_PATH` — path to template in Dropbox (default: /PSD Customers/GPC_Commercial_Workbook_template.xlsm)

## Endpoints
- `GET /health` — health check
- `POST /fill-workbook` — fill workbook with customer data

## Request body (POST /fill-workbook)
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

## Deploy to Railway
1. Push this folder to a GitHub repo
2. Connect repo to Railway
3. Set environment variables
4. Deploy — Railway gives you a public URL

## Make.com HTTP module settings
- URL: https://YOUR-RAILWAY-URL/fill-workbook
- Method: POST
- Body type: Raw
- Content type: application/json
- Body: { customer data from module 3 }
- Parse response: No (returns binary file)
