import sys
import time
import json
import requests
import msal

# ========= CONFIG =========
""
CLIENT_ID = "02aca95c-4084-407f-bff3-3704a66570d2"  # from your app registration
AUTHORITY = "https://login.microsoftonline.com/consumers"  # personal Outlook.com accounts
# AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Mail.Read"]

# Optional: set a subject phrase to filter (case-insensitive contains)
SUBJECT_CONTAINS = "[Corpora-List]"   # set to "" or None to fetch without search
MAX_MESSAGES = 50             # total to print/save
SAVE_JSON_PATH = "messages.json"  # set to None to skip saving
TARGET_DATE = "2025-08-18"
# ==========================


def acquire_token_device_code(client_id: str, authority: str, scopes: list[str]) -> str:
    app = msal.PublicClientApplication(client_id=client_id, authority=authority)
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to create device code flow: {flow}")
    print(f"\n== Sign in =="
          f"\nOpen: {flow['verification_uri']}"
          f"\nCode:  {flow['user_code']}\n")
    result = app.acquire_token_by_device_flow(flow)  # blocks until completed
    if "access_token" not in result:
        raise RuntimeError(f"Token acquisition failed: {result}")
    return result["access_token"]


def list_messages(access_token: str,
                  subject_contains: str | None,
                  max_messages: int = 50) -> list[dict]:
    """
    Returns a list of message dicts with basic fields.
    Uses $search for subject contains (KQL) when subject_contains is provided.
    Paginates via @odata.nextLink.
    """
    s = requests.Session()
    s.headers.update({
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json"
    })

    base_url = "https://graph.microsoft.com/v1.0/me/messages"
    params = {
        "$select": "id,subject,from,receivedDateTime,conversationId,webLink",
        "$top": "25",
        "$orderby": "receivedDateTime DESC"
    }

    if subject_contains:
        # Enable search across message properties with KQL query on subject
        # KQL example: subject:"report"
        s.headers["ConsistencyLevel"] = "eventual"
        params.clear()
        params.update({
            "$search": f"\"subject:{subject_contains}\"",
            "$select": "id,subject,from,receivedDateTime,conversationId,webLink",
            "$top": "25"
        })

    results = []
    url = base_url
    while url and len(results) < max_messages:
        r = s.get(url, params=params if url == base_url else None)
        if r.status_code in (429, 503, 504):
            # Basic retry on throttling/transient errors
            retry_after = int(r.headers.get("Retry-After", "5"))
            time.sleep(retry_after)
            continue
        r.raise_for_status()
        payload = r.json()
        batch = payload.get("value", [])
        results.extend(batch)
        url = payload.get("@odata.nextLink") if len(results) < max_messages else None

    return results[:max_messages]


def main():
    if CLIENT_ID == "YOUR-APP-CLIENT-ID":
        print("ERROR: Put your Client ID into CLIENT_ID at the top of the script.")
        sys.exit(1)

    token = acquire_token_device_code(CLIENT_ID, AUTHORITY, SCOPES)
    msgs = list_messages(token, SUBJECT_CONTAINS, MAX_MESSAGES)

    if not msgs:
        print("No messages found.")
        return

    print(f"\nFetched {len(msgs)} message(s):\n")
    for i, m in enumerate(msgs, 1):
        sender = (m.get("from") or {}).get("emailAddress") or {}
        print(f"{i:02d}. {m.get('receivedDateTime')} | {sender.get('address','(no sender)')} | {m.get('subject','(no subject)')}")
        # Optional: print a direct link to the message in Outlook on the web
        # print(f"     {m.get('webLink')}")

    if SAVE_JSON_PATH:
        with open(SAVE_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(msgs, f, ensure_ascii=False, indent=2)
        print(f"\nSaved to {SAVE_JSON_PATH}")


if __name__ == "__main__":
    main()
