from __future__ import print_function
import os
import sys
import json
import base64
import re
from email import message_from_bytes
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import pandas as pd  # ✅ NEW — for Excel operations

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# === File paths ===
main_path = "C:\\Users\\pc\\Desktop\\Documents\\Programming\\Automation Tableau De Board"
excel_name = "TDB_CP_BPILOT_V4 (12).xlsm"
sheet_name = "TdeB_OFFI"


# ==========================================================
# Authenticate Gmail API
# ==========================================================
def authenticate_gmail():
    creds = None
    token_path = f"{main_path}\\main\\dist\\token.pickle"
    credentials_path = f"{main_path}\\main\\dist\\credentials.json"

    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)

    return build('gmail', 'v1', credentials=creds)

# ==========================================================
# Read Excel and get Designation value
# ==========================================================
def get_designation_from_excel(project_name, organe_name, designation):
    """
    Filters Excel file by 'Project' and 'Organe' columns, 
    then retrieves the last 'Designation' value from the result.
    """
    excel_path = f"{main_path}\\excel\\{excel_name}"

    # Load Excel sheet
    df = pd.read_excel(excel_path, header=None, sheet_name=sheet_name, engine='openpyxl')

    # Apply filtering 
    # index of Projet culomn is 1
    # index of Organe culomn is 2
    # index of Designation culomn is 5
    filtered_df = df[
        (df[1].astype(str).str.strip() == project_name.strip()) &
        (df[2].astype(str).str.strip() == organe_name.strip()) &
        (df[5].astype(str).str.strip() == designation.strip())
    ]

    if filtered_df.empty:
        print(f"No rows found for Project={project_name} Organe={organe_name} and designation={designation}")
        return None

    # Get the last row index and the corresponding Designation value
    # index of Designation culomn is 1
    last_row = filtered_df.tail(1)
    last_index = filtered_df.index[-1]
    # designation_value = last_row[5].values[0]

    print(f"last_index: {last_index}")

    # print(last_row)

    # print(f"✅ Designation found: {designation_value}")
    return last_index

# ==========================================================
# Fetch email content
# ==========================================================
def get_email(service, query):
    print(f"query: {query}")
    results = service.users().messages().list(userId='me', q=query, maxResults=1).execute()
    messages = results.get('messages', [])
    
    if not messages:
        print("No matching email found.")
        return None

    msg = service.users().messages().get(userId='me', id=messages[0]['id'], format='raw').execute()
    msg_raw = base64.urlsafe_b64decode(msg['raw'])
    email_message = message_from_bytes(msg_raw)

    if email_message.is_multipart():
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode('utf-8', errors='ignore')
    else:
        return email_message.get_payload(decode=True).decode('utf-8', errors='ignore')

# ==========================================================
# Extract values from email body
# ==========================================================
def extract_values(body, last_index):
    hw = re.search(r"HW ref[:\-]\s*(.*)", body)
    sw = re.search(r"SW ref[:\-]\s*(.*)", body)
    cal = re.search(r"CAL ref[:\-]\s*(.*)", body)

    hwp = re.search(r"HW patterns[:\-]\s*(.*)", body)
    swp = re.search(r"SW patterns[:\-]\s*(.*)", body)
    calp = re.search(r"CAL patterns[:\-]\s*(.*)", body)
    
    print(f"Extract value last_index: {last_index}")

    return {
        "hw_ref": hw.group(1).strip() if hw else "",
        "sw_ref": sw.group(1).strip() if sw else "",
        "cal_ref": cal.group(1).strip() if cal else "",
        "hw_patterns": hwp.group(1).strip() if hwp else "",
        "sw_patterns": swp.group(1).strip() if swp else "",
        "cal_patterns": calp.group(1).strip() if calp else "",
        "last_index": int(last_index) + 1 if last_index else 0
    }

# ==========================================================
# Main entry point
# ==========================================================
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python get_gmail_data.py <ProjectName> <Organe>")
        sys.exit(1)

    project_name = sys.argv[1]
    organe_name = sys.argv[2]
    designation = sys.argv[3]

    # Step 1: Authenticate Gmail
    service = authenticate_gmail()

    # Step 2: Get designation and Excel data
    last_index = get_designation_from_excel(project_name, organe_name, designation)

    # Step 3: Search Gmail
    query = f'subject:("{project_name}_{organe_name}_{designation}")'
    email_body = get_email(service, query)

    if not email_body:
        sys.exit("❌ No email found for given query.")

    # Step 4: Extract values
    extracted = extract_values(email_body, last_index)

    # Step 5: Save values to data.json
    result_path = f"{main_path}\\main\\result\\data.json"
    
    with open(result_path, "w", encoding="utf-8") as f:
        json.dump(extracted, f, indent=4)

    print(f"💾 Extracted values saved: {extracted}")
