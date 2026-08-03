from flask import Flask, request, jsonify, send_file
import openpyxl
import requests
import os
import io
import tempfile

app = Flask(__name__)

DROPBOX_REFRESH_TOKEN = os.environ.get('DROPBOX_REFRESH_TOKEN', '')
DROPBOX_APP_KEY = os.environ.get('DROPBOX_APP_KEY', '')
DROPBOX_APP_SECRET = os.environ.get('DROPBOX_APP_SECRET', '')
TEMPLATE_PATH = os.environ.get('TEMPLATE_PATH', '/PSD Customers/GPC_Commercial_Workbook_template.xlsm')

def get_dropbox_access_token():
    """Exchange refresh token for a fresh access token"""
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
    """Download template xlsm from Dropbox"""
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
    """Fill customer fields in workbook and return as bytes"""
    with tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False) as tmp:
        tmp.write(template_bytes)
        tmp_path = tmp.name

    wb = openpyxl.load_workbook(tmp_path, keep_vba=True)
    ws = wb['Inputs']

    ws['F8'] = data.get('gp_account_name', '')
    ws['F9'] = data.get('owner_name', '')

    address_parts = [
        data.get('farm_address', ''),
        data.get('city', ''),
        data.get('state', ''),
        data.get('zip', '')
    ]
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

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

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

        return send_file(
            filled_workbook,
            mimetype='application/vnd.ms-excel.sheet.macroEnabled.12',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
