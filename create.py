from Google import Create_Service
import pandas as pd

CLIENT_SECRET_FILE_DRIVE = 'client_secret_drive.json'
API_NAME_DRIVE = 'drive'
API_VERSION_DRIVE = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

# Sheet related service parameters
CLIENT_SECRET_FILE_SHEET = 'client_secret_sheets.json'
API_NAME_SHEET = 'sheets'
API_VERSION_SHEET = 'v4'

drive_service = Create_Service(CLIENT_SECRET_FILE_DRIVE, API_NAME_DRIVE, API_VERSION_DRIVE, SCOPES)
sheet_service = Create_Service(CLIENT_SECRET_FILE_SHEET, API_NAME_SHEET, API_VERSION_SHEET, SCOPES)


folder_metadata = {
    'name': 'my_folder_for_sheets',
    'mimeType': 'application/vnd.google-apps.folder'
}
folder_prop = drive_service.files().create(body=folder_metadata).execute()
target_folder_id = folder_prop['id']

sheet01_name = 'Historic Prices'
sheet02_name = 'Returns'
sheet03_name = 'Risk-Returns'
sheet04_name = 'Final Report'

file_metadata = {
    'properties': {
        'title': 'Stock Portfolio Analysis',
        'locale': 'en_US',
        'timeZone': 'America/Los Angeles',
        'autoRecalc': 'ON_CHANGE'
    },
    'sheets': [
        {
            'properties': {
                'title': sheet01_name
            }
        },
        {
            'properties': {
                'title': sheet02_name
            }
        },
        {
            'properties': {
                'title': sheet03_name
            }
        },
        {
            'properties': {
                'title': sheet04_name
            }
        }
    ]
}
file_prop = sheet_service.spreadsheets().create(body=file_metadata).execute()
file_id = file_prop['spreadsheetId']

drive_service.files().update(
    fileId=file_id,
    addParents=target_folder_id
).execute()