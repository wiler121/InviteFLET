import flet
from msal import ConfidentialClientApplication, PublicClientApplication
import requests
import json
import asyncio
import threading

# The App Registration's application (client) ID
client_id = '12981ac8-3f9b-48b0-9839-918ecebbca70'

# Your Azure AD tenant ID
tenant_id = 'ea7b03ca-00db-4c19-9b8c-6442ab16a46f'

#SCOPES = ['User.Read.All']
SCOPES = ["https://graph.microsoft.com/.default"]

# The App Registration's application (client) ID
client_id = '12981ac8-3f9b-48b0-9839-918ecebbca70'

# Your Azure AD tenant ID
tenant_id = 'ea7b03ca-00db-4c19-9b8c-6442ab16a46f'

ENDPOINT_URI = 'https://graph.microsoft.com/v1.0/'

app = PublicClientApplication(
    client_id=client_id,
    authority=f"https://login.microsoftonline.com/{tenant_id}")

req_headers = {}


def invite_user(mail, message):
    json_body = {
    "invitedUserEmailAddress": mail,
    "invitedUserMessageInfo": {
        "customizedMessageBody": message
    },
    "sendInvitationMessage": True,
    "inviteRedirectUrl": f"https://account.activedirectory.windowsazure.com/?tenantid={tenant_id}&login_hint={mail}"
    }
    response = requests.post(url=ENDPOINT_URI + 'invitations', headers=req_headers, json=json_body)
    print(json.dumps(response.json(), indent=5, ensure_ascii=False))
    response_json = response.json()['invitedUserEmailAddress'] + " | " + response.json()['status']
    return response_json


def gen_token():
    acquire_tokens_result = app.acquire_token_interactive(scopes=SCOPES)
    req_headers = {
        "Authorization": "Bearer " + acquire_tokens_result['access_token'],
        "Content-Type": "application/json"
    }
    return req_headers


def main(page: flet.page):
    def token(p):
        global req_headers
        req_headers = gen_token()
        token_field.value = 'Generated'
        token_row.update()

    def invite(p):
        mails = mails_field.value
        mails_list = mails.split()
        message = message_field.value
        for mail in mails_list:
            result = (invite_user(mail, message))
            result_field.controls.append(flet.Text(result))
            result_row.update()

    page.window_width = 800
    page.window_height = 600

    token_field = flet.Text(height=50, width=150, value="None", size=30)
    generate_token_btn = flet.ElevatedButton(text="Generate token", on_click=token)

    message_field = flet.TextField(width=750, text_size=14, hint_text='WIADOMOŚĆ')

    mails_field = flet.TextField(width=630, text_size=14, hint_text='ADRESY MAILOWE ODZIELONE SPACJĄ')
    invite_btn = flet.ElevatedButton(text="invite guests", on_click=invite)

    result_field = flet.ListView(height=500, width=750, expand=1)

    token_row = flet.Row(alignment=flet.MainAxisAlignment.START, controls=[token_field, generate_token_btn])
    messagge_row = flet.Row(alignment=flet.MainAxisAlignment.START, controls=[message_field])
    mail_row = flet.Row(alignment=flet.MainAxisAlignment.START, controls=[mails_field, invite_btn])
    result_row = flet.Row(alignment=flet.MainAxisAlignment.SPACE_BETWEEN, controls=[result_field])

    page.add(token_row, messagge_row, mail_row, result_row)


flet.app(target=main)
