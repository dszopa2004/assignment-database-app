import gspread, os
import tkinter as tk
from tkinter import ttk
from google.auth import exceptions
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from gspread_formatting import *

get_working_dir = os.getcwd()

frm = tk.Tk()
style = ttk.Style(frm)
frm.geometry('350x175')
frm.title("Assignment Manager")
frm.resizable(False, False)
p1 = tk.PhotoImage(file = 'icon_image.png')
frm.iconphoto(True, p1)

# Path to theme, including the images
frm.tk.call('source', get_working_dir + r'\breeze-dark\breeze-dark.tcl')  
style.theme_use('breeze-dark')  # Theme name

frm.columnconfigure(0, weight=2)
frm.columnconfigure(1, weight=2)
frm.columnconfigure(2, weight=2)

frm.rowconfigure(0, weight=1)
frm.rowconfigure(1, weight=0)

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

######### FUNCTIONS #########
######### FUNCTIONS #########

# This function stores the data that the user inputs
# into variables, which then gets pushed to the spreadsheet
def store_input():
    global course, hw, date 
    course = course_name.get("1.0", "end-1c")
    hw = hw_details.get("1.0", "end-1c")
    date = due_date.get("1.0", "end-1c")

# This function sorts the sheet
# Sorts the due_date column by date
# Sorts the checkboxes
def sort_sheet():
    creds = authenticate()
    client = gspread.authorize(creds)
    spreadsheet = client.open('Homework Manager')
    worksheet = spreadsheet.get_worksheet(0)  # 0 represents the first sheet

    # Sort the sheet 
    worksheet.sort((3, 'asc')) # Change the num to change columns
    worksheet.sort((4, 'asc')) 

    print("Sheet sorted successfully!")


# This function adds checkboxes to the 'D' row
def add_checkbox_validation(worksheet, row):
    target_range_of_cells = f'D{row}:D{row}'
    validation_rule = DataValidationRule(
        BooleanCondition('BOOLEAN', []),
        showCustomUi=True
    )
    set_data_validation_for_cell_range(worksheet, target_range_of_cells, validation_rule)


# This function creates placeholder text for the textfields
def create_placeholder(event, widget, placeholder_text):
    if widget.get("1.0", "end-1c") == placeholder_text:
        widget.delete("1.0", "end-1c")
        widget.configure(fg="white")  # Change text color to black if it's a placeholder

    if event == "leave" and not widget.get("1.0", "end-1c"):
        widget.insert("1.0", placeholder_text)
        widget.configure(fg="gray")  # Change text color to gray if the field is empty


# Default authentication function from Google Sheet API
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
    spreadsheet = client.open('Homework Manager')
    worksheet = spreadsheet.get_worksheet(0)  # 0 represents the first sheet

    new_row = [course, hw, date]
    worksheet.append_row(new_row)

    add_checkbox_validation(worksheet, new_row)

    print("New row added successfully!")
    
######### FUNCTIONS #########
######### FUNCTIONS #########

######### WIDGETS #########
######### WIDGETS #########
placeholder_course_name = "Course Name"
placeholder_hw_details = "Assignment"
placeholder_due_date = "Due Date"

btn_add = tk.Button(frm, text ="Add to Spreadsheet", relief='ridge', command=lambda: [store_input(), main()])
btn_sort = tk.Button(frm, text ="Sort Spreadsheet", relief='ridge', command=sort_sheet)
course_name = tk.Text(frm, height = 1,width = 15, fg="gray")
hw_details = tk.Text(frm, height = 1,width = 15, fg="gray")
due_date = tk.Text(frm, height = 1,width = 15, fg="gray")

# Course name placeholder
course_name.insert("1.0", placeholder_course_name)
course_name.bind("<FocusIn>", lambda event: create_placeholder("click", course_name, placeholder_course_name))
course_name.bind("<FocusOut>", lambda event: create_placeholder("leave", course_name, placeholder_course_name))
# Due date placeholder
due_date.insert("1.0", placeholder_due_date)
due_date.bind("<FocusIn>", lambda event: create_placeholder("click", due_date, placeholder_due_date))
due_date.bind("<FocusOut>", lambda event: create_placeholder("leave", due_date, placeholder_due_date))
# HW details placeholder
hw_details.insert("1.0", placeholder_hw_details)
hw_details.bind("<FocusIn>", lambda event: create_placeholder("click", hw_details, placeholder_hw_details))
hw_details.bind("<FocusOut>", lambda event: create_placeholder("leave", hw_details, placeholder_hw_details))

btn_add.grid(row=1, column=1, pady=3)
btn_sort.grid(row=2, column=1, pady=3)
course_name.grid(row=0, column=0, padx=1)
hw_details.grid(row=0, column=1, padx=1)
due_date.grid(row=0, column=2, padx=1)
frm.mainloop()

######### WIDGETS #########
######### WIDGETS #########