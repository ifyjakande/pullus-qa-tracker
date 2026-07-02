#!/usr/bin/env python3
"""
Check if source data worksheets have been modified since last update.
Uses content hash of source worksheets to detect actual data changes.
Sends Google Chat notifications when changes are detected.
"""

import base64
import binascii
import json
import os
import sys
import time
import hashlib
import gspread
from gspread.utils import absolute_range_name, fill_gaps
import requests
from datetime import datetime
import pytz
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

# Constants
HASH_FILE = 'last_source_hash.json'
MAX_RETRIES = 4
RETRYABLE_STATUS_CODES = (429, 500, 502, 503, 504)

def api_call_with_retry(func, *args, **kwargs):
    """Call a Sheets API function, retrying 429/5xx errors with exponential backoff."""
    for attempt in range(MAX_RETRIES):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = e.response.status_code
            if status in RETRYABLE_STATUS_CODES and attempt < MAX_RETRIES - 1:
                wait = 5 * (2 ** attempt)
                print(f"⚠️ API error {status}, retrying in {wait}s (attempt {attempt + 1}/{MAX_RETRIES})")
                time.sleep(wait)
            else:
                raise

def load_env_file():
    """Load environment variables from .env file if it exists."""
    env_file = os.path.join(os.path.dirname(__file__), '.env')
    if os.path.exists(env_file):
        load_dotenv(env_file)
        print("🔐 Loaded credentials from .env file")

def get_credentials():
    """Get Google API credentials from environment."""
    # Try CI environment first, then local .env
    if os.getenv('CI') != 'true':
        load_env_file()

    credentials_path = os.getenv('GOOGLE_CREDENTIALS_PATH')
    if not credentials_path:
        raise ValueError("GOOGLE_CREDENTIALS_PATH environment variable not set")

    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/drive.metadata.readonly'
        ]

        credentials_value = credentials_path.strip()
        if (credentials_value.startswith('"') and credentials_value.endswith('"')) or \
           (credentials_value.startswith("'") and credentials_value.endswith("'")):
            credentials_value = credentials_value[1:-1].strip()

        def _parse_credentials(raw_value):
            try:
                return json.loads(raw_value)
            except json.JSONDecodeError:
                return None

        credentials_info = _parse_credentials(credentials_value)
        credential_source = None

        if credentials_info is None:
            cleaned = ''.join(credentials_value.split())
            if cleaned:
                padding = len(cleaned) % 4
                if padding:
                    cleaned += '=' * (4 - padding)
                for decoder in (base64.b64decode, base64.urlsafe_b64decode):
                    try:
                        decoded_bytes = decoder(cleaned)
                        decoded_str = decoded_bytes.decode('utf-8').strip()
                        credentials_info = _parse_credentials(decoded_str)
                        if credentials_info is not None:
                            credential_source = 'base64'
                            break
                    except (binascii.Error, UnicodeDecodeError):
                        continue

        else:
            credential_source = 'embedded'

        if credentials_info is not None:
            source_label = 'embedded JSON' if credential_source == 'embedded' else 'base64 JSON'
            print(f"🔑 Using {source_label} service account credentials")
            return Credentials.from_service_account_info(credentials_info, scopes=scopes)

        if os.path.isfile(credentials_value):
            print("🔑 Using service account credentials from file path")
            return Credentials.from_service_account_file(credentials_value, scopes=scopes)

        raise ValueError(
            "GOOGLE_CREDENTIALS_PATH must contain either service account JSON "
            "(raw or base64-encoded) or a valid filesystem path to the JSON file"
        )

    except FileNotFoundError:
        print(f"❌ Credentials file not found: {credentials_path}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error loading credentials: {e}")
        sys.exit(1)

def get_source_data_hash(spreadsheet_id, credentials, source_worksheets):
    """Get combined content hash of all source worksheets."""
    try:
        gc = gspread.authorize(credentials)
        print(f"✅ Authorized with Google Sheets API")

        spreadsheet = gc.open_by_key(spreadsheet_id)
        print(f"✅ Opened spreadsheet successfully")

        existing_titles = {ws.title for ws in api_call_with_retry(spreadsheet.worksheets)}
        found_worksheets = []
        for worksheet_name in source_worksheets:
            if worksheet_name in existing_titles:
                found_worksheets.append(worksheet_name)
            else:
                print(f"⚠️ Worksheet not found, skipping")

        combined_content = []

        if found_worksheets:
            # Single batchGet for all tabs instead of 2 requests per tab
            ranges = [absolute_range_name(name) for name in found_worksheets]
            response = api_call_with_retry(spreadsheet.values_batch_get, ranges)

            for worksheet_name, value_range in zip(found_worksheets, response.get('valueRanges', [])):
                # Same padding as get_all_values() so the stored hash stays valid
                all_values = fill_gaps(value_range.get('values', [[]]))

                # Add worksheet content to combined content
                combined_content.append({
                    'worksheet': worksheet_name,
                    'data': all_values
                })

                rows_with_data = len([row for row in all_values if any(cell.strip() for cell in row)])
                print(f"📊 {worksheet_name}: {rows_with_data} rows with data")

        # Create hash of combined content
        content_str = str(combined_content)
        content_hash = hashlib.md5(content_str.encode('utf-8')).hexdigest()

        print(f"🔗 Combined content hash generated")
        return content_hash

    except gspread.exceptions.APIError:
        print(f"❌ Google Sheets API Error occurred")
        sys.exit(1)
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"❌ Spreadsheet not found or access denied")
        print(f"   Please ensure the service account has been granted access")
        sys.exit(1)
    except Exception:
        print(f"❌ Error accessing spreadsheet data")
        sys.exit(1)

def get_drive_modified_time(spreadsheet_id, credentials):
    """Get the spreadsheet's Drive modifiedTime (cheap pre-check, separate quota)."""
    session = AuthorizedSession(credentials)
    response = session.get(
        f"https://www.googleapis.com/drive/v3/files/{spreadsheet_id}",
        params={'fields': 'modifiedTime', 'supportsAllDrives': 'true'},
        timeout=30
    )
    response.raise_for_status()
    return response.json()['modifiedTime']

def load_state():
    """Load the last processed state (content hash + Drive modifiedTime) from file."""
    try:
        if os.path.exists(HASH_FILE):
            with open(HASH_FILE, 'r') as f:
                return json.load(f)
        return {}
    except Exception as e:
        print(f"⚠️ Warning: Could not load last hash: {e}")
        return {}

def save_state(content_hash, modified_time=None):
    """Save the current content hash and Drive modifiedTime to file."""
    try:
        data = {
            'content_hash': content_hash
        }
        if modified_time:
            data['modified_time'] = modified_time
        with open(HASH_FILE, 'w') as f:
            json.dump(data, f, indent=2)
        print(f"💾 Saved new content hash")
    except Exception as e:
        print(f"❌ Error saving hash: {e}")
        sys.exit(1)

def send_google_chat_notification(worksheets_changed, spreadsheet_id):
    """Send a formatted card notification to Google Chat webhook."""
    webhook_url = os.getenv('GOOGLE_CHAT_WEBHOOK_URL')
    if not webhook_url:
        print("⚠️ Warning: GOOGLE_CHAT_WEBHOOK_URL not set, skipping notification")
        return

    # Get current time in WAT (West Africa Time) with 12-hour format
    wat_tz = pytz.timezone('Africa/Lagos')
    current_time = datetime.now(wat_tz)
    formatted_time = current_time.strftime('%B %d, %Y at %I:%M %p WAT')

    # Create spreadsheet URL
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"

    # Format worksheet names
    sheets_list = ', '.join(worksheets_changed)

    # Create card message
    card_message = {
        "cards": [{
            "header": {
                "title": "🔔 QA Tracker Update Detected",
                "subtitle": "Changes found in source worksheets"
            },
            "sections": [{
                "widgets": [
                    {
                        "keyValue": {
                            "topLabel": "Updated Sheet(s)",
                            "content": sheets_list,
                            "icon": "DESCRIPTION"
                        }
                    },
                    {
                        "keyValue": {
                            "topLabel": "Detected At",
                            "content": formatted_time,
                            "icon": "CLOCK"
                        }
                    },
                    {
                        "buttons": [{
                            "textButton": {
                                "text": "OPEN SPREADSHEET",
                                "onClick": {
                                    "openLink": {
                                        "url": sheet_url
                                    }
                                }
                            }
                        }]
                    }
                ]
            }]
        }]
    }

    try:
        response = requests.post(
            webhook_url,
            json=card_message,
            headers={'Content-Type': 'application/json; charset=UTF-8'},
            timeout=10
        )

        if response.status_code == 200:
            print(f"✅ Google Chat notification sent successfully")
        else:
            print(f"⚠️ Warning: Google Chat notification failed with status {response.status_code}: {response.text}")

    except Exception as e:
        print(f"⚠️ Warning: Could not send Google Chat notification: {e}")

def main():
    """Main function to check for changes in source data."""
    try:
        # Get environment variables
        spreadsheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not spreadsheet_id:
            print("❌ GOOGLE_SHEET_ID environment variable not set")
            sys.exit(1)

        # Source worksheets to monitor
        worksheets_env = os.getenv('SOURCE_WORKSHEETS', '').strip()
        if not worksheets_env:
            print("❌ SOURCE_WORKSHEETS environment variable not set")
            sys.exit(1)

        source_worksheets = [w.strip() for w in worksheets_env.split(',') if w.strip()]

        print(f"🔍 Checking for changes in source worksheets...")

        # Get credentials
        credentials = get_credentials()

        # Load last processed state
        state = load_state()
        last_hash = state.get('content_hash')
        last_modified = state.get('modified_time')

        # Drive modifiedTime early exit: the service account never writes to this
        # file, so an unchanged modifiedTime means unchanged content. Fail open on
        # any error and fall through to the full hash check.
        current_modified = None
        try:
            current_modified = get_drive_modified_time(spreadsheet_id, credentials)
            if last_hash and last_modified and current_modified == last_modified:
                print("⏭️ No changes in source data detected. Skipping update.")
                print("NEEDS_UPDATE=false")
                return False
        except Exception as e:
            print(f"⚠️ Warning: Drive modifiedTime check failed, falling back to hash check: {e}")

        # Get current source data hash
        current_hash = get_source_data_hash(spreadsheet_id, credentials, source_worksheets)

        if last_hash:
            print(f"📅 Previous check completed")
        else:
            print(f"📅 First time check")

        # Compare hashes
        if current_hash != last_hash:
            print("✅ Source data changes detected! Update needed.")
            save_state(current_hash, current_modified)

            # Send Google Chat notification
            send_google_chat_notification(source_worksheets, spreadsheet_id)

            print("NEEDS_UPDATE=true")
            return True
        else:
            # Content unchanged; remember the new modifiedTime so the cheap
            # gate can short-circuit future runs
            if current_modified and current_modified != last_modified:
                save_state(current_hash, current_modified)
            print("⏭️ No changes in source data detected. Skipping update.")
            print("NEEDS_UPDATE=false")
            return False

    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
