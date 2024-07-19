from google_sheets_api import GoogleSheetsAPI
from googleapiclient.errors import HttpError
import re


CREDENTIALS_FILE = 'stylesheet-project-429720-8bb6dca344c7.json'

SOURCE_SPREADSHEET_ID = '1x4yNdVpfpkfJJXjbKtpV2PjyHL76zkhTYY_5xwhV624'

UPLOAD_SPREADSHEET_ID = '1brK-jpAzLofA2IKRHj66AAEZZRZYBFVmsCvXyPevaGY'

google_sheets_source = GoogleSheetsAPI(CREDENTIALS_FILE, SOURCE_SPREADSHEET_ID)

google_sheets_upload = GoogleSheetsAPI(CREDENTIALS_FILE, UPLOAD_SPREADSHEET_ID)


def clean_phone_number(phone_number):
    return re.sub(r'\D', '', phone_number)


def get_sheet_values(google_sheets, sheet_name, range_index):
    try:
        result = google_sheets.service.spreadsheets().values().get(
            spreadsheetId=SOURCE_SPREADSHEET_ID,
            range=f"{sheet_name}!{range_index}"
        ).execute()
        return result.get('values', [])
    except HttpError as err:
        print(f"An error occurred: {err}")
        return []
    
def find_matching_phone_numbers(google_sheets):
    kwiz_data = get_sheet_values(google_sheets, "Kwiz", "A:E")
    zoho_data = get_sheet_values(google_sheets, "Zoho", "A:D")

    matches = []
    kwiz_phone_column_index = 2  
    zoho_phone_column_index = 1  

    kwiz_phone_dict = {clean_phone_number(row[kwiz_phone_column_index]): row for row in kwiz_data[1:] if len(row) > kwiz_phone_column_index}  
    zoho_phone_dict = {clean_phone_number(row[zoho_phone_column_index]): row for row in zoho_data[1:] if len(row) > zoho_phone_column_index}
    
    for phone, kwiz_row in kwiz_phone_dict.items():
        if phone in zoho_phone_dict:
            zoho_row = zoho_phone_dict[phone]
            matches.append((kwiz_row, zoho_row))

    if matches:
        print("Знайдено відповідні номери телефонів:")
        for kwiz_row, zoho_row in matches:
            print(f"Kwiz: {kwiz_row}, Zoho: {zoho_row}")
    else:
        print("Не знайдено жодного відповідного номера телефону.")

    return matches

def upload_matching_data(google_sheets, matches, sheet_name, start_range):
    data = [["Kwiz ПІП", "Kwiz Телефон", "Zoho ПІП родича", "Zoho Телефон родича"]]
    for kwiz_row, zoho_row in matches:
        data.append([kwiz_row[0], kwiz_row[2], zoho_row[0], zoho_row[1]])
    
    range_name = f"{sheet_name}!{start_range}"
    google_sheets.update_google_sheet(data, range_name)


if __name__ == "__main__":

    if google_sheets_source.service is None or google_sheets_upload.service is None:
        print("Не вдалося ініціалізувати сервіс Google Sheets API.")
    else:
        matches = find_matching_phone_numbers(google_sheets_source)
        upload_matching_data(google_sheets_upload, matches, "users_by_numbers", "A1")

