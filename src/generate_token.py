from google_auth_oauthlib.flow import InstalledAppFlow
import pickle

# The Gmail API scope
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# Path to your credentials.json
creds_path = r"C:\Users\pc\Desktop\Documents\Programming\Automation Excel\main\dist\credentials.json"

# Output path for your token
token_path = r"C:\Users\pc\Desktop\Documents\Programming\Automation Excel\main\dist\token.pickle"

flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
creds = flow.run_local_server(port=0)

# Save the token
with open(token_path, 'wb') as token:
    pickle.dump(creds, token)

print(f"✅ Token saved successfully at: {token_path}")
