import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from base64 import urlsafe_b64decode, urlsafe_b64encode
from mimetypes import guess_type as guess_mime_type
from io import BytesIO
from pdfminer.high_level import extract_text
import xlrd
from googleapiclient.http import MediaIoBaseDownload
import driver
import structured1
import structured
import pdfplumber
import io
# Request all access (permission to read/send/receive emails, manage the inbox, and more)

sagicor_id = ''
guardian_id = ''

sagicor_id1 = ''
guardian_id1 = ''
canopy = ''
def structured_function(text):

    data = structured.extract_structured_data_canopy(text)

    
    print(data)
    # sheet=[]
    # for i in range(len(data)):
    #     new_list = []
    #     for key, value in data[i].items():
    #         new_list.append(value)
    #     sheet.append(new_list)

    # driver.add_row(sagicor_id1,'Sheet1',sheet)


def xls_sheet(sheet):


    new_sheet = []
    for i in range(5, len(sheet)):
        list1 = [sheet[0][1],sheet[1][1],sheet[2][1],sheet[3][1],sheet[0][12],sheet[1][12],sheet[2][12]]
        
        list1 += sheet[i][:9]+sheet[i][-5:]
        new_sheet.append(list1)

    driver.add_row(sagicor_id1,'Sheet2', new_sheet)

def parse_parts(service, parts, message):
    """
    Utility function that parses the content of an email partition
    """
    if parts:
        for part in parts:
            filename = part.get("filename")
 
            body = part.get("body")
            data = body.get("data")
        
            part_headers = part.get("headers")
            if part.get("parts"):
                parse_parts(service, part.get("parts"),  message)
            else:
                for part_header in part_headers:
                    part_header_name = part_header.get("name")
                    part_header_value = part_header.get("value")
                    if part_header_name == "Content-Disposition":
                        if "attachment" in part_header_value:
            
                            attachment_id = body.get("attachmentId")
                            attachment = service.users().messages() \
                                        .attachments().get(id=attachment_id, userId='me', messageId=message['id']).execute()
                            data = attachment.get("data")
                  
                            print("attachment file : ", filename)

                            decoded_data = urlsafe_b64decode(data)

                            if filename.__contains__('pdf/') or filename.__contains__('pdf_'):

                                
                                pdf_file = BytesIO(decoded_data)
                                pdf_text = extract_text(pdf_file)

                                with pdfplumber.open(io.BytesIO(decoded_data)) as pdf:
                                    # Loop through each page
                                    for page_number, page in enumerate(pdf.pages, start=1):
                                        # Extract the table from the PDF page
                                        table = page.extract_table()
                                        # If a table is found, process it
                                        if table:
                                            print(f"Table found on page {page_number}:")
                                            # Each row of the table (including the header) is an item in the list
                                            table = table[1:]
                                            driver.add_row(canopy,'Sheet1',table)





                                # structured_function(pdf_text)
                            # if filename.__contains__('.xls'):
                            #     xls_file = BytesIO(decoded_data)

                            #     # Open the workbook
                            #     workbook = xlrd.open_workbook(file_contents=xls_file.getvalue())

                            #     # By default, we will look at the first sheet in the workbook
                            #     sheet = workbook.sheet_by_index(0)
                            #     sheet_list = []
                            #     for row in range(sheet.nrows):
                            #         # Extract the row values
                            #         row_values = sheet.row_values(row)
                            #         # Append the row values to the 'sheet_list'
                            #         sheet_list.append(row_values)
                            #     # print(sheet_list)
                            #     xls_sheet(sheet_list)
                        

      
def get_size_format(b, factor=1024, suffix="B"):
    """
    Scale bytes to its proper byte format
    e.g:
        1253656 => '1.20MB'
        1253656678 => '1.17GB'
    """
    for unit in ["", "K", "M", "G", "T", "P", "E", "Z"]:
        if b < factor:
            return f"{b:.2f}{unit}{suffix}"
        b /= factor
    return f"{b:.2f}Y{suffix}"


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)
def gmail_authenticate(builder, token_pickle,credential_json,vs):
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists(token_pickle):
        with open(token_pickle, "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credential_json, SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open(token_pickle, "wb") as token:
            pickle.dump(creds, token)
    return build(builder, vs, credentials=creds)

def search_messages(service, query):
    result = service.users().messages().list(userId='me',q=query).execute()
    messages = [ ]
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId='me',q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
    return messages

def read_message(service, message):
    """
    This function takes Gmail API `service` and the given `message_id` and does the following:
        - Downloads the content of the email
        - Prints email basic information (To, From, Subject & Date) and plain/text parts
        - Creates a folder for each email based on the subject
        - Downloads text/html content (if available) and saves it under the folder created as index.html
        - Downloads any file that is attached to the email and saves it in the folder created
    """
    msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
    # parts can be the message body, or attachments
    payload = msg['payload']
    headers = payload.get("headers")
    parts = payload.get("parts")
    folder_name = "email"
    has_subject = False
    if headers:
        # this section prints email basic info & creates a folder for the email
        for header in headers:
            name = header.get("name")
            value = header.get("value")
            if name.lower() == 'from':
                # we print the From address
                print("From:", value)
            if name.lower() == "to":
                # we print the To address
                print("To:", value)
            if name.lower() == "subject":
                # make our boolean True, the email has "subject"
                has_subject = True

            if name.lower() == "date":
                # we print the date when the message was sent
                print("Date:", value)
    if not has_subject:
        # if the email does not have a subject, then make a folder with "email" name
        # since folders are created based on subjects
        if not os.path.isdir(folder_name):
            os.mkdir(folder_name)
    parse_parts(service, parts, message)

    print("="*50)

SCOPES = ['https://mail.google.com/']
our_email = 'info@primedicalja.com'


service = gmail_authenticate('gmail', 'token.pickle','credential1.json', vs = 'v1')

results = search_messages(service, "@canopy-insurance.com")
print(f"Found {len(results)} results.")
# for each email matched, read it (output plain/text to console & save HTML and attachments)
for i in range(len(results)):
    read_message(service, results[len(results)-i-1])

# for i in results:
#     read_message(service, i)




