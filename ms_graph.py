import os
import webbrowser
import msal
from dotenv import load_dotenv

def get_access_token(application_id, client_secret):
    client = msal.ConfidentialClientApplication(
        client_id=application_id,
        client_credential=client_secret,
        authority="https://login.microsoftonline.com/consumers/",
    )

    authorization_url = client.get_authorization_url(scopes)
    webbrowser.open(auth_request_url)
    authorization_code = input("Enter authorization code: ")

    token_response = client.acquire_toek_by_authorization_code(
        code=authorization_code,
        scopes=scopes
    )

    if 'access_token' in token_response:
        return token_response["access_token"]
    else:
        raise Exception("Access Token not found")

def main():
    load_dotenv()
    APPLICATION_ID = os.getenv("APPLICATION_ID")
    CLIENT_SECRET = os.getenv("CLIENT_SECRET")
    SCOPES = ['User.Read', 'Files.ReadWrite.All']

    try:
        access_token = get_access_token(application_id=APPLICATION_ID, client_secret=CLIENT_SECRET)
        headers = {
            "Authorization": 'Bearer' + access_token,
        }
        print(headers)
    except Exception as e:
        print(f'Error: {e}')

main()