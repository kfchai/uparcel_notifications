import os
import json
import requests
import pytz
from datetime import datetime, timedelta
import pandas as pd
from dotenv import load_dotenv

# ----------------------------
# Configuration
# ----------------------------
# ----------------------------
# Load Environment Variables
# ----------------------------

load_dotenv()  # Loads from .env

TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
USER_EMAIL = os.getenv('USER_EMAIL')
TESTMODE = os.getenv('TESTMODE')
ADMIN_PHONE = os.getenv('ADMIN_PHONE')
WHATSAPP_TOKEN = os.getenv('WHATSAPP_TOKEN')
PHONE_NUMBER_ID = os.getenv('PHONE_NUMBER_ID')
WHATSAPP_API_URL = f"https://graph.facebook.com/v19.0/{PHONE_NUMBER_ID}/messages"

# Email filtering settings
SEARCH_SUBJECT_KEYWORD = "Uparcel Integration Daily Job CSV"
FOLDER_PATH = ["Automation", "Uparcel Notifications"]
DOWNLOAD_DIR = "./attachments"
TOKEN_FILE = "graph_token.json"

timezone = pytz.timezone(os.getenv("TIMEZONE"))
TODAY = datetime.now(timezone).date()
# ----------------------------
# Token caching and refresh
# ----------------------------

def get_graph_token():
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'r') as f:
            token_data = json.load(f)
            expires_at = datetime.fromisoformat(token_data['expires_at'])
            if datetime.utcnow() < expires_at:
                return token_data['access_token']

    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'grant_type': 'client_credentials',
        'scope': 'https://graph.microsoft.com/.default'
    }
    resp = requests.post(url, data=data)
    resp.raise_for_status()
    new_token = resp.json()

    access_token = new_token['access_token']
    expires_in = new_token['expires_in']
    expires_at = datetime.utcnow() + timedelta(seconds=expires_in - 60)

    with open(TOKEN_FILE, 'w') as f:
        json.dump({
            'access_token': access_token,
            'expires_at': expires_at.isoformat()
        }, f)

    return access_token

access_token = get_graph_token()

headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

# ----------------------------
# Navigate to target subfolder
# ----------------------------

def find_folder_id(user_email, path_list):
    folder_id = None
    current_path = "https://graph.microsoft.com/v1.0/users/{}/mailFolders".format(user_email)
    
    for folder_name in path_list:
        response = requests.get(current_path, headers=headers)
        response.raise_for_status()
        folders = response.json().get('value', [])
        matched = next((f for f in folders if f['displayName'] == folder_name), None)
        if not matched:
            raise Exception(f"Folder '{folder_name}' not found")
        folder_id = matched['id']
        current_path = f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/{folder_id}/childFolders"
    
    return folder_id

folder_id = find_folder_id(USER_EMAIL, FOLDER_PATH)

# ----------------------------
# Find today's email with matching subject
# ----------------------------

#messages_url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/mailFolders/{folder_id}/messages?$orderby=receivedDateTime desc&$top=10"
messages_url = (
    f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/mailFolders/{folder_id}/messages"
    f"?$filter=isRead eq false&$orderby=receivedDateTime desc&$top=10"
)
response = requests.get(messages_url, headers=headers)
response.raise_for_status()
messages = response.json().get('value', [])

selected_message = None
for msg in messages:
    subject = msg.get('subject', '')
    utc_time = datetime.strptime(msg['receivedDateTime'], "%Y-%m-%dT%H:%M:%SZ")
    utc_time = utc_time.replace(tzinfo=pytz.utc)
    received_time = utc_time.astimezone(timezone).date()
    if (SEARCH_SUBJECT_KEYWORD in subject and received_time == TODAY) or subject == "TEST UPARCEL CSV":
        selected_message = msg
        break

if not selected_message:
    raise Exception("No email with matching subject received today.")

message_id = selected_message['id']

# ----------------------------
# Download Excel attachment
# ----------------------------

attachments_url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/messages/{message_id}/attachments"
response = requests.get(attachments_url, headers=headers)
response.raise_for_status()
attachments = response.json().get('value', [])

attachment_path = None
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

for att in attachments:
    if att['name'].endswith('.csv'):
        content_url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/messages/{message_id}/attachments/{att['id']}/$value"
        att_response = requests.get(content_url, headers=headers)
        att_response.raise_for_status()
        file_path = os.path.join(DOWNLOAD_DIR, att['name'])
        with open(file_path, 'wb') as f:
            f.write(att_response.content)
        attachment_path = file_path
        break

if not attachment_path:
    raise Exception("No Excel attachment found in the email.")

mark_read_url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/messages/{message_id}"
mark_read_payload = {
    "isRead": True
}
mark_resp = requests.patch(mark_read_url, headers=headers, json=mark_read_payload)
mark_resp.raise_for_status()

# ----------------------------
# Send WhatsApp message using Meta API
# ----------------------------

def send_whatsapp_message(to_number: str, name: str, delivery_time: str, order_num: str):
    headers = {
        'Authorization': f'Bearer {WHATSAPP_TOKEN}',
        'Content-Type': 'application/json'
    }
    if TESTMODE == '1':
        to_number = ADMIN_PHONE
        
    payload = {
        "messaging_product": "whatsapp",
        "to": to_number,
        "type": "template",
        "template": {
            "name": "uparcel_delivery_reminder",  # This must match your approved template
            "language": { "code": "en" },
            "components": [
                {
                    "type": "body",
                    "parameters": [
                        { "type": "text", "text": name },
                        { "type": "text", "text": order_num },
                        { "type": "text", "text": delivery_time }
                    ]
                }
            ]
        }
    }

    response = requests.post(WHATSAPP_API_URL, headers=headers, json=payload)
    print(f"Sent to {to_number}: {response.status_code} - {response.text}")
    return response.json()


# ----------------------------
# Send WhatsApp Reminder Report To Admin
# ----------------------------

def send_whatsapp_report(to_number: str, report: list):
    headers = {
        'Authorization': f'Bearer {WHATSAPP_TOKEN}',
        'Content-Type': 'application/json'
    }
        
    payload = {
        "messaging_product": "whatsapp",
        "to": to_number,
        "type": "template",
        "template": {
            "name": "messages_report",  # This must match your approved template
            "language": { "code": "en" },
            "components": [
                {
                    "type": "body",
                    "parameters": [
                        { "type": "text", "text": "Uparcel Reminder" },
                        { "type": "text", "text": report['total'] },
                        { "type": "text", "text": report['success'] },
                        { "type": "text", "text": report['failed'] }
                    ]
                }
            ]
        }
    }

    response = requests.post(WHATSAPP_API_URL, headers=headers, json=payload)
    print(f"Sent to {to_number}: {response.status_code} - {response.text}")
    return response.json()

# ----------------------------
# Process Excel and send messages
# ----------------------------

df = pd.read_csv(attachment_path)

wa_report = {
    "success": 0,
    "failed": 0,
    "total": 0
}

for _, row in df.iterrows():
    wa_report['total'] += 1
    name = str(row.get('delivery_contact_person', 'Customer'))
    phone = str(row.get('delivery_contact_number', '')).replace(' ', '')
    delivery_time = str(row.get('delivery_time', 'today'))
    order_num = str(row.get('reference_number','NA'))
    print("Sending WhatsApp Message {0} {1} {2} {3}".format(name, phone, delivery_time, order_num))
    if phone:
        wa_resp = send_whatsapp_message(phone, name, delivery_time, order_num)
        if wa_resp and len(wa_resp['messages']) > 0:
            if wa_resp['messages'][0]['message_status'] == "accepted":
               wa_report['success'] += 1
            else:
               wa_report['failed'] += 1
        else:
           wa_report['failed'] += 1
    else:
        wa_report['failed'] += 1


#sending report
send_whatsapp_report(ADMIN_PHONE, wa_report)
