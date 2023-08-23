import gspread
from google.auth import exceptions
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

def authenticate():
    creds = None
    token_file = 'token.json'  # Change this to your preferred token file name

    # The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_file, 'w') as token:
            token.write(creds.to_json())
    return creds

def main():
    creds = authenticate()

    client = gspread.authorize(creds)
    spreadsheet = client.open('test-sheet')
    worksheet = spreadsheet.get_worksheet(0)  # 0 represents the first sheet

    new_row = ["Data3", "Data4", "Data5"]
    worksheet.append_row(new_row)

    print("New row added successfully!")

if __name__ == '__main__':
    main()