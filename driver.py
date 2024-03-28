from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os
import time
from googleapiclient.errors import HttpError
# SCOPES defines which permissions to request from the user.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None

# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credential2.json', SCOPES)
        creds = flow.run_local_server(port=0)
    
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

# Build the service object.
service = build('sheets', 'v4', credentials=creds)

def add_row(spreadsheet_id,range_name,list):
    print(list)

    # Name of the sheet and range where you want to append data
    # range_name = 'Sheet2'  # For example, A1 starts appending at the first column

    # How the input data should be interpreted.
    value_input_option = 'USER_ENTERED'  # or 'RAW'

    # How the input data should be inserted.
    insert_data_option = 'INSERT_ROWS'  # or 'OVERWRITE'

    # The new row to append.
    values = list

    body = {
        'values': values
    }

    # Use the Sheets API to append the new row.
    request = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption=value_input_option,
        insertDataOption=insert_data_option,
        body=body
    )
    MAX_RETRIES = 5
    wait_time = 1  # start with 1 second

    for attempt in range(MAX_RETRIES):
        try:
            response = request.execute()
            break
        except HttpError as e:
            if e.resp.status == 429:
                print(f"Waiting {wait_time} seconds before retrying...")
                time.sleep(wait_time)
                wait_time *= 2  # double the wait time for the next attempt
            else:
                raise e
    else:
        print("Request failed after maximum number of retries.")

    # Print response
    print(response)