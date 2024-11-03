import os
import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
AUTHORITY = os.getenv("AUTHORITY")
USER_EMAIL = os.getenv("USER_EMAIL")
SCOPES = ['User.Read']
GRAPH_API_URL = "https://graph.microsoft.com/v1.0"


def get_access_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )

    token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" in token_response:
        return token_response["access_token"]
    else:
        raise Exception("Could not obtain access token")


def get_outlook_emails(access_token, user_email):
    endpoint = f"{GRAPH_API_URL}/users/{user_email}/messages"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        emails = response.json().get("value", [])
        return emails
    else:
        raise Exception(f"Error fetching emails: {response.status_code} - {response.text}")


if __name__ == "__main__":    
    try:
        access_token = get_access_token()

        print("\nFetching emails from Outlook...\n")
        emails = get_outlook_emails(access_token, USER_EMAIL)
        for email in emails:
            print(f"Subject: {email['subject']}, From: {email['from']['emailAddress']['address']}\n")

      
    except Exception as e:
        print(f"An error occurred: {e}")
