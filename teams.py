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
SCOPES = ['https://graph.microsoft.com/.default']
GRAPH_API_URL = "https://graph.microsoft.com/v1.0"


def get_access_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )

    token_response = app.acquire_token_for_client(scopes=SCOPES)

    if "access_token" in token_response:
        return token_response["access_token"]
    else:
        raise Exception("Could not obtain access token")


def get_all_outlook_emails(access_token, user_email):
    endpoint = f"{GRAPH_API_URL}/users/{user_email}/messages"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }
    emails = []
    while endpoint:
        response = requests.get(endpoint, headers=headers)
        if response.status_code == 200:
            data = response.json()
            emails.extend(data.get("value", []))
            endpoint = data.get("@odata.nextLink")  # Get the link for the next page if available
        else:
            raise Exception(f"Error fetching emails: {response.status_code} - {response.text}")
    
    return emails


def format_email_details(email):
    """
    Format the details of an email in a readable manner.
    """
    subject = email.get("subject", "No Subject")
    sender = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown Sender")
    received_date = email.get("receivedDateTime", "Unknown Date")
    body_preview = email.get("bodyPreview", "No Preview Available")
    to_recipients = ", ".join([recipient['emailAddress']['address'] for recipient in email.get("toRecipients", [])])
    cc_recipients = ", ".join([recipient['emailAddress']['address'] for recipient in email.get("ccRecipients", [])])

    return (
        f"Subject: {subject}\n"
        f"From: {sender}\n"
        f"To: {to_recipients}\n"
        f"CC: {cc_recipients}\n"
        f"Received Date: {received_date}\n"
        f"Body Preview: {body_preview}\n"
        f"{'-'*40}\n"
    )


if __name__ == "__main__":    
    try:
        access_token = get_access_token()

        print("\nFetching emails from Outlook...\n")
        emails = get_all_outlook_emails(access_token, USER_EMAIL)
        for email in emails:
            print(format_email_details(email))
        print(f"\nTotal emails fetched: {len(emails)}")

    except Exception as e:
        print(f"An error occurred: {e}")
