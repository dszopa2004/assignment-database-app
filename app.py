import gspread
import tkinter as tk
from tkinter import ttk
from google.auth import exceptions
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from gspread_formatting import *

frm = tk.Tk()
style = ttk.Style(frm)
frm.geometry('315x200')
frm.title("Assignment Manager")

frm.tk.call('source', 'D:/assignment-database-app/breeze-dark/breeze-dark.tcl')  # Put here the path of your theme file
style.theme_use('breeze-dark')  # Theme files create a ttk theme, here you can put its name


SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

# This function stores the data that the user inputs
# into variables, which then gets pushed to the spreadsheet
def store_input():
    global course, hw, date 
    course = course_name.get("1.0", "end-1c")
    hw = hw_details.get("1.0", "end-1c")
    date = due_date.get("1.0", "end-1c")


# This function adds checkboxes to the 'D' row
def add_checkbox_validation(worksheet, row):
    target_range_of_cells = f'D{row}:D{row}'
    validation_rule = DataValidationRule(
        BooleanCondition('BOOLEAN', []),
        showCustomUi=True
    )
    set_data_validation_for_cell_range(worksheet, target_range_of_cells, validation_rule)


def authenticate():
    creds = None
    token_file = 'token.json'  # Change this to your preferred token file name

    # The file token.json stores the user's access and refresh tokens, 
    # and is created automatically when the authorization flow completes for the first time.
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

    new_row = [course, hw, date]
    worksheet.append_row(new_row)

    add_checkbox_validation(worksheet, new_row)

    print("New row added successfully!")
    

B = tk.Button(frm, text ="Add to Spreadsheet", command=lambda: [store_input(), main()])
course_name = tk.Text(frm, height = 1,width = 10)
hw_details = tk.Text(frm, height = 1,width = 10)
due_date = tk.Text(frm, height = 1,width = 10)

B.pack(padx=10, pady=10, side='bottom')
course_name.pack(padx=10, pady=10, side='left')
hw_details.pack(padx=10, pady=10, side='left')
due_date.pack(padx=10, pady=10, side='left')
frm.mainloop()