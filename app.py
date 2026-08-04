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

# ── PIN Authentication ──────────────────────────────────────
@app.route('/auth', methods=['POST'])
def auth():
    try:
        data = request.get_json()
        pin = str(data.get('pin', ''))
        if not pin:
            return jsonify({'error': 'PIN required'}), 400

        # Check Sales PIN
        if pin == os.environ.get('PIN_SALES', ''):
            return jsonify({'success': True, 'role': 'sales', 'name': 'Sales'})

        # Check Analyst PIN
        if pin == os.environ.get('PIN_ANALYST', ''):
            return jsonify({'success': True, 'role': 'analyst', 'name': 'Analyst'})

        # Check Installer PINs (PIN_INSTALLER_{team_id})
        for key, value in os.environ.items():
            if key.startswith('PIN_INSTALLER_') and pin == value:
                team_id = key.replace('PIN_INSTALLER_', '')
                # Get team name
                try:
                    res = requests.get(f'{HUBSPOT_BASE}/settings/v3/users/teams', headers=HUBSPOT_HEADERS())
                    teams = res.json().get('results', [])
                    team_name = next((t['name'] for t in teams if str(t['id']) == team_id), 'Installer')
                except:
                    team_name = 'Installer'
                return jsonify({'success': True, 'role': 'installer', 'name': team_name, 'team_id': team_id})

        return jsonify({'error': 'Invalid PIN'}), 401
    except Exception as e:
        return jsonify({'error': str(e)}), 500

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

# ── HubSpot: Update line item ───────────────────────────────
@app.route('/hubspot/line-items/<line_item_id>', methods=['PATCH'])
def hs_update_line_item(line_item_id):
    try:
        data = request.get_json()
        payload = {'properties': {}}
        if 'unit_price' in data:
            payload['properties']['price'] = str(data['unit_price'])
        if 'quantity' in data:
            payload['properties']['quantity'] = str(data['quantity'])
        if 'name' in data:
            payload['properties']['name'] = data['name']
        res = requests.patch(f'{HUBSPOT_BASE}/crm/v3/objects/line_items/{line_item_id}', json=payload, headers=HUBSPOT_HEADERS())
        if res.status_code not in [200, 204]:
            return jsonify({'error': res.text}), res.status_code
        return jsonify({'success': True, 'line_item_id': line_item_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── HubSpot: Update deal ─────────────────────────────────────
@app.route('/hubspot/deals/<deal_id>', methods=['PATCH'])
def hs_update_deal(deal_id):
    try:
        data = request.get_json()
        payload = {'properties': {}}
        if 'amount' in data:
            payload['properties']['amount'] = str(data['amount'])
        if 'dealstage' in data:
            payload['properties']['dealstage'] = data['dealstage']
        if 'dealname' in data:
            payload['properties']['dealname'] = data['dealname']
        res = requests.patch(f'{HUBSPOT_BASE}/crm/v3/objects/deals/{deal_id}', json=payload, headers=HUBSPOT_HEADERS())
        if res.status_code not in [200, 204]:
            return jsonify({'error': res.text}), res.status_code
        return jsonify({'success': True, 'deal_id': deal_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── HubSpot: Add line item to deal ──────────────────────────
@app.route('/hubspot/line-items', methods=['POST'])
def hs_create_line_item():
    try:
        data = request.get_json()
        deal_id = data.get('deal_id')
        if not deal_id:
            return jsonify({'error': 'deal_id is required'}), 400
        payload = {
            'properties': {
                'name': data.get('name', 'Fan Motor'),
                'price': str(data.get('unit_price', 0)),
                'quantity': str(data.get('quantity', 1)),
                'description': data.get('description', ''),
                'fan_brand': data.get('fan_brand', ''),
                'fan_size': str(data.get('fan_size', '')),
                'motor_size': str(data.get('motor_size', '')),
            }
        }
        res = requests.post(f'{HUBSPOT_BASE}/crm/v3/objects/line_items', json=payload, headers=HUBSPOT_HEADERS())
        if res.status_code != 201:
            return jsonify({'error': res.text}), res.status_code
        line_item_id = res.json()['id']
        assoc_payload = {'inputs': [{'from': {'id': deal_id}, 'to': {'id': line_item_id}, 'type': 'deal_to_line_item'}]}
        requests.post(f'{HUBSPOT_BASE}/crm/v3/associations/deals/line_items/batch/create', json=assoc_payload, headers=HUBSPOT_HEADERS())
        return jsonify({'success': True, 'line_item_id': line_item_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── HubSpot: Create new deal ─────────────────────────────────
@app.route('/hubspot/deals', methods=['POST'])
def hs_create_deal():
    try:
        data = request.get_json()
        payload = {
            'properties': {
                'dealname': data.get('dealname', ''),
                'dealstage': data.get('dealstage', 'appointmentscheduled'),
                'pipeline': 'default',
                'amount': str(data.get('amount', 0)),
                'psd_gp_account_name': data.get('gp_account_name', ''),
                'psd_gp_account_number': data.get('gp_account_number', ''),
                'psd_owner_name': data.get('owner_name', ''),
                'psd_farm_address': data.get('farm_address', ''),
                'psd_city': data.get('city', ''),
                'psd_state': data.get('state', ''),
                'psd_zip': data.get('zip', ''),
                'total_fans': str(data.get('total_fans', '')),
                'total_barns': str(data.get('total_barns', '')),
            }
        }
        res = requests.post(f'{HUBSPOT_BASE}/crm/v3/objects/deals', json=payload, headers=HUBSPOT_HEADERS())
        if res.status_code != 201:
            return jsonify({'error': res.text}), res.status_code
        deal = res.json()
        return jsonify({'success': True, 'deal_id': deal['id'], 'dealname': deal['properties'].get('dealname')})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── HubSpot: Create new contact ──────────────────────────────
@app.route('/hubspot/contacts', methods=['POST'])
def hs_create_contact():
    try:
        data = request.get_json()
        payload = {
            'properties': {
                'email': data.get('email', ''),
                'firstname': data.get('firstname', ''),
                'lastname': data.get('lastname', ''),
                'phone': data.get('phone', ''),
                'company': data.get('company', ''),
                'address': data.get('address', ''),
                'city': data.get('city', ''),
                'state': data.get('state', ''),
                'zip': data.get('zip', ''),
                'ower_name': data.get('owner_name', ''),
                'georgia_power_account': str(data.get('gp_account_number', '')),
            }
        }
        res = requests.post(f'{HUBSPOT_BASE}/crm/v3/objects/contacts', json=payload, headers=HUBSPOT_HEADERS())
        if res.status_code != 201:
            return jsonify({'error': res.text}), res.status_code
        contact = res.json()
        deal_id = data.get('deal_id')
        if deal_id:
            assoc_payload = {'inputs': [{'from': {'id': deal_id}, 'to': {'id': contact['id']}, 'type': 'deal_to_contact'}]}
            requests.post(f'{HUBSPOT_BASE}/crm/v3/associations/deals/contacts/batch/create', json=assoc_payload, headers=HUBSPOT_HEADERS())
        return jsonify({'success': True, 'contact_id': contact['id'], 'email': contact['properties'].get('email')})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── HubSpot: Get teams ───────────────────────────────────────
@app.route('/hubspot/teams', methods=['GET'])
def hs_get_teams():
    try:
        res = requests.get(f'{HUBSPOT_BASE}/settings/v3/users/teams', headers=HUBSPOT_HEADERS())
        if res.status_code != 200:
            return jsonify({'error': res.text}), res.status_code
        teams = []
        for t in res.json().get('results', []):
            teams.append({
                'id': t.get('id'),
                'name': t.get('name'),
                'userIds': t.get('userIds', [])
            })
        return jsonify({'teams': teams})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── HubSpot: Get services for a deal ────────────────────────
@app.route('/hubspot/services/<deal_id>', methods=['GET'])
def hs_get_services(deal_id):
    try:
        # Get associated service IDs
        assoc_res = requests.get(f'{HUBSPOT_BASE}/crm/v3/objects/deals/{deal_id}/associations/services', headers=HUBSPOT_HEADERS())
        if assoc_res.status_code != 200:
            return jsonify({'services': []})
        service_ids = [r['id'] for r in assoc_res.json().get('results', [])]
        if not service_ids:
            return jsonify({'services': []})
        props = 'hs_service_name,description,hs_pipeline,hs_pipeline_stage,hs_service_status,start_date,target_end_date,hs_total_cost,hs_amount_paid,hs_remaining_amount'
        services = []
        for sid in service_ids:
            res = requests.get(f'{HUBSPOT_BASE}/crm/v3/objects/services/{sid}?properties={props}', headers=HUBSPOT_HEADERS())
            if res.status_code == 200:
                s = res.json()
                services.append({'id': s['id'], 'properties': s.get('properties', {})})
        return jsonify({'services': services})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── HubSpot: Create service for a deal ──────────────────────
@app.route('/hubspot/services', methods=['POST'])
def hs_create_service():
    try:
        data = request.get_json()
        deal_id = data.get('deal_id')
        if not deal_id:
            return jsonify({'error': 'deal_id is required'}), 400

        payload = {
            'properties': {
                'subject': data.get('name', ''),
                'description': data.get('description', ''),
                'hs_pipeline': 'default',
                'hs_pipeline_stage': data.get('stage', 'new'),
                'hs_ticket_priority': data.get('status', 'ON_TRACK'),
                'createdate': data.get('start_date', ''),
                'hs_due_date': data.get('target_end_date', ''),
                'hs_ticket_category': 'Fan Motor Installation',
            }
        }
        if data.get('team_id'):
            payload['properties']['hubspot_team_id'] = str(data.get('team_id', ''))
        # Remove empty values
        payload['properties'] = {k: v for k, v in payload['properties'].items() if v}

        res = requests.post(f'{HUBSPOT_BASE}/crm/v3/objects/services', json=payload, headers=HUBSPOT_HEADERS())
        if res.status_code != 201:
            return jsonify({'error': res.text}), res.status_code

        service_id = res.json()['id']

        # Associate service to deal
        assoc_payload = {'inputs': [{'from': {'id': deal_id}, 'to': {'id': service_id}, 'type': 'deal_to_service'}]}
        requests.post(f'{HUBSPOT_BASE}/crm/v3/associations/deals/services/batch/create', json=assoc_payload, headers=HUBSPOT_HEADERS())

        return jsonify({'success': True, 'service_id': service_id, 'name': data.get('name', '')})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── Proxy to Make webhooks (intake pipeline only) ───────────
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
