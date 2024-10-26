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


def send_outlook_email(access_token, subject, body, recipients):
    endpoint = f"{GRAPH_API_URL}/users/{USER_EMAIL}/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    email_data = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [{"emailAddress": {"address": recipient}} for recipient in recipients]
        },
        "saveToSentItems": "true"
    }

    response = requests.post(endpoint, json=email_data, headers=headers)
    
    if response.status_code == 202:
        print("Email sent successfully!")
    else:
        raise Exception(f"Error sending email: {response.status_code} - {response.text}")


def get_teams_messages(access_token):
    endpoint = f"{GRAPH_API_URL}/users/{USER_EMAIL}/chats"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        teams_messages = response.json().get("value", [])
        return teams_messages
    else:
        raise Exception(f"Error fetching Teams messages: {response.status_code} - {response.text}")


if __name__ == "__main__":    
    try:
        access_token = get_access_token()

        print("\nFetching emails from Outlook...\n")
        emails = get_outlook_emails(access_token, USER_EMAIL)
        for email in emails:
            print(f"Subject: {email['subject']}, From: {email['from']['emailAddress']['address']}\n")

        print("\nSending a test email...\n")
        send_outlook_email(
            access_token,
            subject="Test Email",
            body="This is a test email sent from a Python script.",
            recipients=[USER_EMAIL]
        )

        print("\nFetching messages from Microsoft Teams...\n")
        teams_messages = get_teams_messages(access_token)
        for message in teams_messages:
            print(f"Chat Id: {message['id']}, Created DateTime: {message['createdDateTime']}\n")

    except Exception as e:
        print(f"An error occurred: {e}")
