import httplib2
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

class GoogleSheetsAPI:
    def __init__(self, credentials_file, spreadsheet_id):
        self.credentials_file = credentials_file
        self.spreadsheet_id = spreadsheet_id
        self.service = self.authorize_google_sheets_api()

    def authorize_google_sheets_api(self):
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
            self.credentials_file,
            ['https://www.googleapis.com/auth/spreadsheets',
             'https://www.googleapis.com/auth/drive']
        )
        httpAuth = credentials.authorize(httplib2.Http())
        service = build('sheets', 'v4', http=httpAuth)
        return service

    def update_google_sheet(self, data, range_name):
        try:
            body = {
                "valueInputOption": "USER_ENTERED",
                "data": [
                    {
                        "range": range_name,
                        "majorDimension": "ROWS",
                        "values": data
                    }
                ]
            }
            result = self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body=body
            ).execute()
            print(f"Updated {result.get('totalUpdatedCells')} cells")
        except HttpError as err:
            print(f"Виникла помилка: {err}")
    
    def read_google_sheet(self, range_name):
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()
            return result.get('values', [])
        except HttpError as err:
            print(f"Виникла помилка: {err}")
            return []