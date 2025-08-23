import msal
import requests
import json
from datetime import datetime

# ================= CONFIG =================
CLIENT_ID = "02aca95c-4084-407f-bff3-3704a66570d2"  # App registration client ID
SCOPES = ["Mail.Read"]                     # Delegated permission
TARGET_DATE = "2025-08-21"                 # YYYY-MM-DD
KEYWORDS = [
    "FDL GL TR CSOT PROD Refresh",
    "FDL GL COMPASS SOT Refresh",
    "FDL GEMERALD PAC Refresh",
    "By Cluster GTM"
]
# ==========================================

AUTHORITY = "https://login.microsoftonline.com/common"

# --- Step 1: Device code flow ---
def acquire_token_device_code():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to create device flow: {flow}")
    
    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)  # blocks until user completes login
    if "access_token" not in result:
        raise RuntimeError(f"Failed to acquire token: {result}")
    return result["access_token"]

access_token = acquire_token_device_code()
headers = {"Authorization": f"Bearer {access_token}"}

# --- Step 2: Date filter for Graph API ---
date_start = f"{TARGET_DATE}T00:00:00Z"
date_end = f"{TARGET_DATE}T23:59:59Z"
date_filter = f"receivedDateTime ge {date_start} and receivedDateTime le {date_end}"

# --- Step 3: Fetch emails with pagination ---
emails = []
endpoint = f"https://graph.microsoft.com/v1.0/me/messages?$top=50&$filter={date_filter}"

while endpoint:
    resp = requests.get(endpoint, headers=headers)
    if resp.status_code != 200:
        raise RuntimeError(f"Graph API error: {resp.status_code} - {resp.text}")
    
    data = resp.json()
    emails.extend(data.get("value", []))
    endpoint = data.get("@odata.nextLink")

# --- Step 4: Filter by subject keywords ---
filtered_emails = [
    m for m in emails
    if any(k.lower() in m.get("subject", "").lower() for k in KEYWORDS)
]

# --- Step 5: Save filtered emails to JSON ---
output_file = f"personal_emails_{TARGET_DATE}.json"
with open(output_file, "w", encoding="utf-8") as f:
    json.dump(filtered_emails, f, indent=2)

print(f"Saved {len(filtered_emails)} emails containing keywords for {TARGET_DATE} to {output_file}")
