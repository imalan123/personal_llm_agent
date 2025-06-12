import os
from datetime import datetime

from bs4 import BeautifulSoup
from dotenv import load_dotenv
import os.path
import base64
import re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from googleScripts.Transaction import Transaction

load_dotenv()

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
URL = [
    "https://mail.google.com/",
    "https://www.googleapis.com/auth/spreadsheets",
    'https://www.googleapis.com/auth/drive.readonly'
]

CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, 'google_credentials.json')
TOKEN_FILE = os.path.join(SCRIPT_DIR, 'google_token.json')
TARGET_LABEL = os.getenv("TARGET_LABEL")
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME")

today = datetime.now().strftime("%Y/%m/%d")
query = f"after:{today}"


def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, URL)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise EnvironmentError(f"Error: '{CREDENTIALS_FILE}' not found.")

            print("Opening browser for Google authentication...")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, URL)
            creds = flow.run_local_server(port=0)  # port=0 finds an available port
            print("Authentication successful.")

        # Save the updated credentials for future runs
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    gmail = build('gmail', 'v1', credentials=creds)
    sheets= build('sheets', 'v4', credentials=creds)
    drive = build('drive', 'v3', credentials=creds)
    return gmail, sheets, drive

def get_label_id(service):
    label_id = None
    labels_response = service.users().labels().list(userId='me').execute()
    labels = labels_response.get('labels', [])
    for label in labels:
        if label['name'].upper() == TARGET_LABEL.upper():
            label_id = label['id']
            break

    if label_id is None:
        print("Available labels:")
        for label in labels:
            print(f"- {label['name']}")
        raise ValueError(f"Label not found")

    return label_id

def get_emails_under_label(label_id, service):
    messages_response = service.users().messages().list(
        userId='me',
        labelIds=[label_id],
        q=query
    ).execute()
    message_ids = messages_response.get('messages', [])

    if not message_ids:
        raise ValueError(f"No messages found under label '{TARGET_LABEL}' matching today's date.")

    all_transactions = []
    for message_id in message_ids:
        message = service.users().messages().get(userId='me', id=message_id['id'], format="full").execute()
        transaction_email = decode_payload_of_email(message)
        print(transaction_email)
        all_transactions.append(transaction_email)
        print(message_id['id'])
        service.users().messages().delete(userId='me', id=message_id['id']).execute()


    return all_transactions

def decode_payload_of_email(message):
    queue = [message['payload']]
    html_body = ""
    while queue:
        current_part = queue.pop(0)

        if 'body' in current_part and 'data' in current_part['body']:
            try:
                data = current_part['body']['data']
                # Decode base64 URL safe string, ignoring errors for robustness
                decoded_data = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                mime_type = current_part.get('mimeType', '')

                if mime_type == 'text/html':
                    html_body += decoded_data  # Accumulate HTML parts

            except Exception as e:
                print(f"Error decoding part: {e}")
                continue

    # If only HTML is found, convert it to plain text using BeautifulSoup
    soup = BeautifulSoup(html_body, 'html.parser')
    soup = soup.get_text(separator='\n').strip()

    date_pattern = r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}'
    date = re.search(date_pattern, soup)
    extracted_date = date.group(0)

    merchant_pattern = r"(Merchant\s*([^\r\n]+))"
    merchant = re.search(merchant_pattern, soup)
    extracted_merchant = merchant.group(2)

    amount_pattern = r"(Amount\s*([^\r\n]+))"
    amount = re.search(amount_pattern, soup)
    extracted_amount = amount.group(2)

    account_pattern = r"(Account\s*([^\r\n]+))"
    account = re.search(account_pattern, soup)
    extracted_account = account.group(2)

    return Transaction(
        **{
        "Credit Card": extracted_account,
        "Merchant": extracted_merchant,
        "Paid Amount": extracted_amount,
        "Date": extracted_date,
        }
    )

def get_sheet_id_with_specific_name(drive_service):
    sheet_query = f"name='{SPREADSHEET_NAME}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"

    # fields='files(id,name)' requests only the ID and name of the matching files
    results = drive_service.files().list(
        q=sheet_query,
        spaces='drive',
        fields='files(id, name)'
    ).execute()

    items = results.get('files', [])

    if not items:
        print(f"No spreadsheet found with the name: '{SPREADSHEET_NAME}'")
        return None
    else:
        return items[0]['id']

def send_transactions_to_sheets(all_transactions, spreadsheet_id, sheet_service):
    rows_for_sheets = []

    for transaction in all_transactions:
        rows_for_sheets.append([
            transaction.card,    # Corresponds to "Credit Card"
            transaction.merchant, # Corresponds to "Merchant"
            transaction.amount,  # Corresponds to "Paid Amount"
            transaction.date     # Corresponds to "Date"
        ])
    body = {
        'values': rows_for_sheets
    }

    sheet_tab_name = create_new_sheet_tab_if_new_month(sheet_service, spreadsheet_id)

    sheets_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,  # Identifies which spreadsheet to write to
        range=f"{sheet_tab_name}!A1",
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',  # appends to existing
        body=body
    ).execute()

    label_column_letter = 'F'
    sum_column_letter = 'G'
    label_target_cell_range = f"'{sheet_tab_name}'!{label_column_letter}3"
    sum_target_cell_range = f"'{sheet_tab_name}'!{sum_column_letter}3"

    label_text = "Total Spend:"
    sum_formula = f"=SUM('{sheet_tab_name}'!C:C)"

    body = {
        'valueInputOption': 'USER_ENTERED',
        'data': [
            {
                'range': label_target_cell_range,
                'values': [[label_text]]
            },
            {
                'range': sum_target_cell_range,
                'values': [[sum_formula]]
            }
        ]
    }

    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()


    print(f"Successfully sent {len(rows_for_sheets)} transactions to {sheet_tab_name}")

def create_new_sheet_tab_if_new_month(sheet_service, spreadsheet_id):
    today_date_obj = datetime.now().date()
    current_month_name = today_date_obj.strftime("%B %Y")

    sheet_exists = False
    spreadsheet_metadata = sheet_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields='sheets.properties'
    ).execute()

    for sheet_prop in spreadsheet_metadata.get('sheets', []):
        if sheet_prop['properties']['title'] == current_month_name:
            print(f"Sheet tab '{current_month_name}' already exists.")
            return current_month_name

    if not sheet_exists:
        print(f"Sheet tab '{current_month_name}' does not exist. Creating new tab...")
        batch_update_request_body = {
            'requests': [
                {
                    'addSheet': {
                        'properties': {
                            'title': current_month_name
                        }
                    }
                }
            ]
        }
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=batch_update_request_body
        ).execute()

    return current_month_name

if __name__ == "__main__":
    gmail_service, sheets_service, drive_service= get_credentials()

    print(f"\n--- Searching for messages under label: '{TARGET_LABEL}' ---")
    gmail_label_id = get_label_id(gmail_service)
    gmail_transactions = get_emails_under_label(gmail_label_id, gmail_service)
    budget_sheet_id = get_sheet_id_with_specific_name(drive_service)

    send_transactions_to_sheets(gmail_transactions, budget_sheet_id, sheets_service)

