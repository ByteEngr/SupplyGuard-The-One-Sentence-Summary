import os
import uuid
import datetime
import requests
import msal

# ==== CONFIG ====
TENANT_ID = os.getenv("TENANT_ID")          # your tenant
CLIENT_ID = os.getenv("CLIENT_ID")          # app registration
CLIENT_SECRET = os.getenv("CLIENT_SECRET")  # client secret

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"

SUPPLIER_NAME = "Contoso Logistics"
SUPPLIER_DOMAIN = "contoso-logistics.example"
SUPPLIER_USERS = ["alice@contoso-logistics.example", "bob@contoso-logistics.example"]
EXPIRY_DAYS = 90

# ==== AUTH ====
def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        raise Exception(f"Could not obtain access token: {result}")
    return result["access_token"]

def graph_post(url, token, json_body):
    resp = requests.post(
        f"{GRAPH_ENDPOINT}{url}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=json_body
    )
    if not resp.ok:
        raise Exception(f"Graph POST {url} failed: {resp.status_code} {resp.text}")
    return resp.json()

def graph_get(url, token, params=None):
    resp = requests.get(
        f"{GRAPH_ENDPOINT}{url}",
        headers={"Authorization": f"Bearer {token}"},
        params=params
    )
    if not resp.ok:
        raise Exception(f"Graph GET {url} failed: {resp.status_code} {resp.text}")
    return resp.json()

# ==== MAIN WORKFLOW ====
def main():
    token = get_access_token()

    # 1) Create Unified Group
    mail_nickname = "".join(c.lower() for c in SUPPLIER_NAME if c.isalnum())
    if len(mail_nickname) > 40:
        mail_nickname = mail_nickname[:40]

    expiry_date = (datetime.datetime.utcnow() + datetime.timedelta(days=EXPIRY_DAYS)).date().isoformat()

    group_body = {
        "displayName": f"SUPPLIER - {SUPPLIER_NAME}",
        "mailNickname": mail_nickname,
        "mailEnabled": True,
        "securityEnabled": False,
        "groupTypes": ["Unified"],
        "description": f"Supplier: {SUPPLIER_NAME} ({SUPPLIER_DOMAIN}) | Expires on {expiry_date}"
    }

    group = graph_post("/groups", token, group_body)
    group_id = group["id"]
    print(f"Created group: {group['displayName']} ({group_id})")

    # 2) Create Team on top of group (standard template)
    team_body = {
        "memberSettings": {"allowCreateUpdateChannels": True},
        "messagingSettings": {
            "allowUserEditMessages": True,
            "allowUserDeleteMessages": True
        },
        "funSettings": {"allowGiphy": True},
        "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
    }

    # PUT /teams/{group-id}
    resp = requests.put(
        f"{GRAPH_ENDPOINT}/teams/{group_id}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=team_body
    )
    if resp.status_code not in (200, 202):
        raise Exception(f"Failed to create team: {resp.status_code} {resp.text}")

    print("Team creation initiated (may take some time).")

    # 3) Get associated site (optional)
    # It may not be immediately available – in real code you'd poll/retry.
    try:
        sites = graph_get(f"/groups/{group_id}/sites", token)
        site_url = sites["value"][0]["webUrl"] if sites.get("value") else None
    except Exception:
        site_url = None

    print(f"Site URL: {site_url}")

    # 4) Invite supplier users as guests and add to group
    for email in SUPPLIER_USERS:
        invite_body = {
            "invitedUserEmailAddress": email,
            "inviteRedirectUrl": site_url or "https://teams.microsoft.com",
            "sendInvitationMessage": True,
            "invitedUserMessageInfo": {
                "customizedMessageBody": f"You've been granted secure supplier access for {SUPPLIER_NAME}."
            }
        }
        invitation = graph_post("/invitations", token, invite_body)
        guest_user = invitation.get("invitedUser", {})
        guest_id = guest_user.get("id")

        if not guest_id:
            print(f"Warning: no guestId returned for {email}, skipping group membership.")
            continue

        # Add to group members
        ref_body = {
            "@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{guest_id}"
        }
        graph_post(f"/groups/{group_id}/members/$ref", token, ref_body)
        print(f"Invited and added {email} to group.")

    print("Supplier onboarding complete.")

if __name__ == "__main__":
    main()
