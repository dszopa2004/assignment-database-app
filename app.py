import gspread
import tkinter as tk
from google.auth import exceptions
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

frm = tk.Tk()
frm.geometry('400x200')

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

    new_row = ["Data5", "Data6", "Data7"]
    worksheet.append_row(new_row)

    print("New row added successfully!")
    

B = tk.Button(frm, text ="Add to Spreadsheet", command = main)
course_name = tk.Text(frm, height = 1,width = 10)
hw_details = tk.Text(frm, height = 1,width = 10)
due_date = tk.Text(frm, height = 1,width = 10)

B.pack(padx=10, pady=10, side='bottom')
course_name.pack(padx=10, pady=10, side='left')
hw_details.pack(padx=10, pady=10, side='left')
due_date.pack(padx=10, pady=10, side='left')
frm.mainloop()