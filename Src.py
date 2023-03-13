import requests

import json

# Replace these variables with your own values

tenant_id = "<your tenant ID>"

client_id = "<your client ID>"

client_secret = "<your client secret>"

user_email = "<the email address of the user whose inbox you want to access>"

# Get an access token using client credentials flow

token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

data = {

    "grant_type": "client_credentials",

    "client_id": client_id,

    "client_secret": client_secret,

    "scope": "https://graph.microsoft.com/.default"

}

response = requests.post(token_url, data=data)

access_token = json.loads(response.text)["access_token"]

# Use the access token to retrieve the user's mailbox ID

headers = {

    "Authorization": f"Bearer {access_token}",

    "Accept": "application/json"

}

url = f"https://graph.microsoft.com/v1.0/users/{user_email}/mailboxsettings"

response = requests.get(url, headers=headers)

mailbox_id = json.loads(response.text)["id"]

# Use the mailbox ID to retrieve messages from the inbox

url = f"https://graph.microsoft.com/v1.0/users/{user_email}/mailfolders/inbox/messages"

params = {

    "$select": "subject,body,sender,from,receivedDateTime",

    "$orderby": "receivedDateTime DESC"

}

response = requests.get(url, headers=headers, params=params)

messages = json.loads(response.text)["value"]

# Print the subject and body of each message

for message in messages:

    print("Subject:", message["subject"])

    print("Body:", message["body"]["content"])

    print("------------------------")

