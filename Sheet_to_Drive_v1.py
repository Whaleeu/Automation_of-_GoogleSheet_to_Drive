import os
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import requests
import sys

os.path.exists("token_file.json")
scope = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name("token_file.json", scope)
#Create a Google Sheets API client
service = build('sheets', 'v4', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)
# Google Sheet parameters
sheet_id = ""

file_id = ""


def getGoogleSeet(spreadsheet_id, outDir, outFile):
    url = f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx'
    response = requests.get(url)
    if response.status_code == 200:
        filepath = os.path.join(outDir, outFile)
        with open(filepath, 'wb') as f:
            f.write(response.content)
            print('XLSX file saved to: {}'.format(filepath))
    else:
        print(f'Error downloading Google Sheet: {response.status_code}')
        sys.exit


def upload_excel_to_drive(drive_service, local_file_path, folder_id=None, file_name='Untitled'):
    media = MediaFileUpload(local_file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
    file_metadata = {
        'name': file_name,
    }

    
    # If a folder_id is provided, set it as the parent
    if folder_id:
        file_metadata['addParents'] = folder_id


    file_drive = drive_service.files().update(fileId=file_id,body=file_metadata,media_body=media).execute()
    print('File uploaded to Google Drive')
##############################################

outDir = 'tmp/'
os.makedirs(outDir, exist_ok = True)
filepath = getGoogleSeet(sheet_id, outDir, "TransactiondATA.xlsx")

# Specify the folder ID in Google Drive where you want to upload the Excel file
folder_id = ""  # Replace with the actual folder ID

# Specify the local path to your existing Excel file
local_excel_file_path = 'tmp/TransactiondATA.xlsx'  # Replace with the actual path

# Upload the Excel file to Google Drive
upload_excel_to_drive(drive_service, local_excel_file_path, folder_id, 'TransactiondATA.xlsx')

sys.exit  # success
