from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
import requests
import os
import io
import re
import json
import tempfile

app = Flask(__name__, static_folder='static')
CORS(app)

DROPBOX_REFRESH_TOKEN = os.environ.get('DROPBOX_REFRESH_TOKEN', '')
DROPBOX_APP_KEY = os.environ.get('DROPBOX_APP_KEY', '')
DROPBOX_APP_SECRET = os.environ.get('DROPBOX_APP_SECRET', '')
TEMPLATE_PATH = os.environ.get('TEMPLATE_PATH', '/PSD Customers/GPC_Commercial_Workbook_template.xlsm')
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY', '')
HUBSPOT_API_KEY = os.environ.get('HUBSPOT_API_KEY', '')
HUBSPOT_BASE = 'https://api.hubapi.com'
HUBSPOT_HEADERS = lambda: {'Authorization': f'Bearer {HUBSPOT_API_KEY}', 'Content-Type': 'application/json'}
PORTAL_ID = '246901747'

def get_dropbox_access_token():
    response = requests.post(
        'https://api.dropbox.com/oauth2/token',
        data={
            'grant_type': 'refresh_token',
            'refresh_token': DROPBOX_REFRESH_TOKEN,
            'client_id': DROPBOX_APP_KEY,
            'client_secret': DROPBOX_APP_SECRET,
        }
    )
    if response.status_code != 200:
        raise Exception(f'Failed to refresh Dropbox token: {response.text}')
    return response.json()['access_token']

def download_template_from_dropbox():
    access_token = get_dropbox_access_token()
    url = 'https://content.dropboxapi.com/2/files/download'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Dropbox-API-Arg': f'{{"path": "{TEMPLATE_PATH}"}}'
    }
    response = requests.post(url, headers=headers)
    if response.status_code != 200:
        raise Exception(f'Failed to download template: {response.text}')
    return response.content

def fill_workbook(template_bytes, data):
    with tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False) as tmp:
        tmp.write(template_bytes)
        tmp_path = tmp.name

    wb = openpyxl.load_workbook(tmp_path, keep_vba=True)
    ws = wb['Inputs']

    ws['F8'] = data.get('gp_account_name', '')
    ws['F9'] = data.get('owner_name', '')
    address_parts = [data.get('farm_address', ''), data.get('city', ''), data.get('state', ''), data.get('zip', '')]
    ws['F10'] = ', '.join(p for p in address_parts if p)
    ws['F11'] = data.get('phone', '')
    ws['F12'] = data.get('email', '')
    ws['M8'] = f"{data.get('gp_account_name', '')} — GP {data.get('gp_account_number', '')}"
    ws['F13'] = 'Agriculture'

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    os.unlink(tmp_path)
    return output

# ── Serve frontend ──────────────────────────────────────────
@app.route('/', methods=['GET'])
def index():
    return app.send_static_file('index.html')

@app.route('/map', methods=['GET'])
def map_page():
    return app.send_static_file('map.html')

@app.route('/maps-key', methods=['GET'])
def maps_key():
    return jsonify({'key': os.environ.get('GOOGLE_MAPS_API_KEY', '')})

# ── Health check ────────────────────────────────────────────
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

# ── Fill workbook ────────────────────────────────────────────
@app.route('/fill-workbook', methods=['POST'])
def fill_workbook_endpoint():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No JSON data provided'}), 400
        required = ['gp_account_name', 'gp_account_number']
        missing = [f for f in required if not data.get(f)]
        if missing:
            return jsonify({'error': f'Missing required fields: {missing}'}), 400
        template_bytes = download_template_from_dropbox()
        filled_workbook = fill_workbook(template_bytes, data)
        safe_name = data.get('gp_account_name', 'Customer').replace('/', '-').replace('\\', '-')
        filename = f"GPC_Commercial_Workbook_{safe_name}.xlsm"
        return send_file(filled_workbook, mimetype='application/vnd.ms-excel.sheet.macroEnabled.12', as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── HubSpot: Get open deals ──────────────────────────────────
@app.route('/hubspot/deals', methods=['GET'])
def hs_get_deals():
    try:
        props = 'dealname,amount,dealstage,closedate,hs_object_id,total_fans,total_barns,operation_type'
        stage_filter = request.args.get('stage', None)
        url = f'{HUBSPOT_BASE}/crm/v3/objects/deals?limit=50&properties={props}'
        res = requests.get(url, headers=HUBSPOT_HEADERS())
        if res.status_code != 200:
            return jsonify({'error': res.text}), res.status_code
        data = res.json()
        deals = []
        for d in data.get('results', []):
            props_data = d.get('properties', {})
            if stage_filter and props_data.get('dealstage') != stage_filter:
                continue
            deals.append({
                'id': d['id'],
                'properties': props_data
            })
        return jsonify({'deals': deals})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── HubSpot: Get line items for a deal ──────────────────────
@app.route('/hubspot/line-items/<deal_id>', methods=['GET'])
def hs_get_line_items(deal_id):
    try:
        # Get associated line item IDs
        assoc_url = f'{HUBSPOT_BASE}/crm/v3/objects/deals/{deal_id}/associations/line_items'
        assoc_res = requests.get(assoc_url, headers=HUBSPOT_HEADERS())
        if assoc_res.status_code != 200:
            return jsonify({'error': assoc_res.text}), assoc_res.status_code

        assoc_data = assoc_res.json()
        line_item_ids = [r['id'] for r in assoc_data.get('results', [])]

        if not line_item_ids:
            return jsonify({'line_items': []})

        # Fetch each line item
        props = 'name,price,quantity,description,fan_brand,fan_size,motor_size,number_of_fans,hs_object_id'
        line_items = []
        for lid in line_item_ids:
            li_res = requests.get(f'{HUBSPOT_BASE}/crm/v3/objects/line_items/{lid}?properties={props}', headers=HUBSPOT_HEADERS())
            if li_res.status_code == 200:
                li = li_res.json()
                line_items.append({
                    'id': li['id'],
                    'properties': li.get('properties', {})
                })

        return jsonify({'line_items': line_items})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── HubSpot: Get deals with addresses ───────────────────────
@app.route('/hubspot/deals-with-addresses', methods=['GET'])
def hs_deals_with_addresses():
    try:
        # Get appointmentscheduled deals
        props = 'dealname,amount,dealstage,total_fans,total_barns,hs_object_id'
        url = f'{HUBSPOT_BASE}/crm/v3/objects/deals?limit=50&properties={props}'
        res = requests.get(url, headers=HUBSPOT_HEADERS())
        if res.status_code != 200:
            return jsonify({'error': res.text}), res.status_code

        deals = [d for d in res.json().get('results', []) if d.get('properties', {}).get('dealstage') == 'appointmentscheduled']

        result = []
        contact_props = 'address,city,state,zip,company,firstname,ower_name,email,phone,georgia_power_account'

        for deal in deals:
            deal_id = deal['id']
            # Get associated contact
            assoc_res = requests.get(f'{HUBSPOT_BASE}/crm/v3/objects/deals/{deal_id}/associations/contacts', headers=HUBSPOT_HEADERS())
            if assoc_res.status_code != 200:
                continue
            contacts = assoc_res.json().get('results', [])
            if not contacts:
                continue

            contact_id = contacts[0]['id']
            contact_res = requests.get(f'{HUBSPOT_BASE}/crm/v3/objects/contacts/{contact_id}?properties={contact_props}', headers=HUBSPOT_HEADERS())
            if contact_res.status_code != 200:
                continue

            contact = contact_res.json().get('properties', {})
            dp = deal.get('properties', {})

            result.append({
                'deal_id': deal_id,
                'deal_name': dp.get('dealname', ''),
                'amount': dp.get('amount', ''),
                'total_fans': dp.get('total_fans', ''),
                'total_barns': dp.get('total_barns', ''),
                'company': contact.get('company', ''),
                'owner': contact.get('ower_name', '') or contact.get('firstname', ''),
                'address': contact.get('address', ''),
                'city': contact.get('city', ''),
                'state': contact.get('state', ''),
                'zip': contact.get('zip', ''),
                'email': contact.get('email', ''),
                'phone': contact.get('phone', '')
            })

        return jsonify({'deals_with_addresses': result})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── Proxy to Make webhooks (write operations only) ───────────
@app.route('/proxy', methods=['POST'])
def proxy():
    try:
        data = request.get_json()
        target_url = data.get('url')
        payload = data.get('payload')

        if not target_url:
            return jsonify({'error': 'Missing target url'}), 400

        response = requests.post(target_url, json=payload, headers={'Content-Type': 'application/json'}, timeout=30)

        content_type = response.headers.get('Content-Type', '')
        if 'text/html' in content_type or response.text.strip().startswith('<'):
            return jsonify({'error': 'Make returned an error', 'status': response.status_code}), 500

        try:
            return jsonify(response.json()), response.status_code
        except Exception:
            try:
                cleaned = response.text.strip()
                cleaned = re.sub(r',\s*}', '}', cleaned)
                cleaned = re.sub(r',\s*]', ']', cleaned)
                return jsonify(json.loads(cleaned)), response.status_code
            except Exception:
                return jsonify({'raw': response.text}), response.status_code

    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── Proxy to Anthropic API ───────────────────────────────────
@app.route('/claude', methods=['POST'])
def claude_proxy():
    try:
        data = request.get_json()
        response = requests.post(
            'https://api.anthropic.com/v1/messages',
            json=data,
            headers={'Content-Type': 'application/json', 'x-api-key': ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01'},
            timeout=60
        )
        return jsonify(response.json()), response.status_code
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
