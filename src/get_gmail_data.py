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

# ==========================================================
# Authenticate Gmail API
# ==========================================================
def authenticate_gmail():
    creds = None
    token_path = r"C:\Users\pc\Desktop\Documents\Programming\Automation Excel\main\dist\token.pickle"
    credentials_path = r"C:\Users\pc\Desktop\Documents\Programming\Automation Excel\main\dist\credentials.json"

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
def get_designation_from_excel(project_name, organe_name):
    """
    Filters Excel file by 'Project' and 'Organe' columns, 
    then retrieves the last 'Designation' value from the result.
    """
    excel_path = r"C:\Users\pc\Desktop\Documents\Programming\Automation Excel\excel\TDB_CP_BPILOT_V4 (12).xlsm"
    sheet_name = "TdeB_OFFI"

    # Load Excel sheet
    df = pd.read_excel(excel_path, header=None, sheet_name=sheet_name, engine='openpyxl')

    # Ensure column names are stripped and consistent
    # df.columns = df.columns.str.strip()

    column_names_index = df.columns
    # print(column_names_index)

    # print(df[1])

    # Apply filtering 
    # index of Projet culomn is 1
    # index of Organe culomn is 2
    filtered_df = df[
        (df[1].astype(str).str.strip() == project_name.strip()) &
        (df[2].astype(str).str.strip() == organe_name.strip())
    ]

    if filtered_df.empty:
        print(f"No rows found for Project={project_name} and Organe={organe_name}")
        return None

    # Get the last row index and the corresponding Designation value
    # index of Designation culomn is 1
    last_row = filtered_df.tail(1)
    last_index = filtered_df.index[-1]
    designation_value = last_row[5].values[0]

    print(f"last_index: {last_index}")

    # print(last_row)

    print(f"✅ Designation found: {designation_value}")
    return designation_value, last_index

def update_excel(filtered_df, last_index):
    # --- Step 2: Load values from result.json ---
    with open(r"C:\Users\pc\Desktop\Documents\Programming\Automation Excel\main\dist\result.json", "r", encoding="utf-8") as f:
        result_data = json.load(f)

    cal_ref = result_data.get("cal_ref")
    sw_ref = result_data.get("sw_ref")
    hw_ref = result_data.get("hw_ref")

    # --- Step 3: Fill the Excel cells ---
    # Reference NFC 'FIAT'  → index - 1
    filtered_df.loc[last_index - 1, "Reference NFC 'FIAT'"] = cal_ref

    # SW' → index - 1
    filtered_df.loc[last_index - 1, "SW'"] = sw_ref

    # HW' → index - 2
    filtered_df.loc[last_index - 2, "HW'"] = hw_ref

    # --- Step 4: Save changes back to the Excel file ---
    excel_path = r"C:\Users\pc\Desktop\Documents\Programming\Automation Excel\main\dist\your_excel_file.xlsx"

    # Load the original full Excel (not only the filtered part)
    df = pd.read_excel(excel_path)

    # Update the corresponding rows in the full DataFrame
    for idx, row in filtered_df.iterrows():
        df.loc[idx] = row

    # Save the updated Excel file
    df.to_excel(excel_path, index=False)
    print("✅ Excel updated successfully.")
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

    # Step 1: Authenticate Gmail
    service = authenticate_gmail()

    # Step 2: Get designation and Excel data
    designation_value, last_index = get_designation_from_excel(project_name, organe_name)
    print(f"✅ Found designation: {designation_value}")

    # Step 3: Search Gmail
    query = f'subject:("{project_name}_{organe_name}_{designation_value}")'
    email_body = get_email(service, query)

    if not email_body:
        sys.exit("❌ No email found for given query.")

    # Step 4: Extract values
    extracted = extract_values(email_body, last_index)

    # Step 5: Save values to result.json
    result_path = r"C:\Users\pc\Desktop\Documents\Programming\Automation Excel\main\result\result.json"
    with open(result_path, "w", encoding="utf-8") as f:
        json.dump(extracted, f, indent=4)

    print(f"💾 Extracted values saved: {extracted}")

    # Step 6: Fill Excel fields based on extracted values
    # cal_ref = extracted.get("cal_ref")
    # sw_ref = extracted.get("sw_ref")
    # hw_ref = extracted.get("hw_ref")

    # Column index references (0-based indexing)
    # col_ref_nfc = 15  # "Reference NFC 'FIAT'"
    # col_hw_fiat = 16  # "HW 'FIAT'"
    # col_sw_fiat = 17  # "SW 'FIAT'"

    # try:
    #     # Update the Excel values using .iloc[row_index, col_index]
    #     df.iloc[last_index - 1, col_ref_nfc] = cal_ref
    #     df.iloc[last_index - 1, col_sw_fiat] = sw_ref
    #     df.iloc[last_index - 2, col_hw_fiat] = hw_ref

    #     print("✅ Excel cells updated successfully:")
    #     print(f"  Reference NFC 'FIAT' (index {col_ref_nfc}) = {cal_ref}")
    #     print(f"  SW 'FIAT' (index {col_sw_fiat}) = {sw_ref}")
    #     print(f"  HW 'FIAT' (index {col_hw_fiat}) = {hw_ref}")

    # except Exception as e:
    #     print(f"⚠️ Error updating Excel cells: {e}")

    # # Step 7: Save Excel
    # df.to_excel(excel_path, index=False, engine='openpyxl')
    print("✅ Excel file saved successfully.")
