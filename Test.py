import os.path
import requests
import httplib2
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime



# Посилання на файл, полученный в Google Developer Console
CREDENTIALS_FILE = 'stylesheet-project-429720-24b5e88f67ec.json'

# ID Google Sheets документа (його беремо в URL)
spreadsheet_id = '1brK-jpAzLofA2IKRHj66AAEZZRZYBFVmsCvXyPevaGY'

# Авторизуємося и та отримаємо service — екзепляр доступу до API
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    ['https://www.googleapis.com/auth/spreadsheets',
     'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = build('sheets', 'v4', http=httpAuth)

today = datetime.date.today()

currencies = ["USD", "EUR", "CHF", "GBP", "PLN", "SEK", "ILS", "CAD"]

years = [i for i in range(today.year-4, today.year)]

data = [["Currency"] + years]

#створюємо цикл щоб отримати дані курсів валют за 2020-2023
for currency in currencies:
    row = [currency]
    for year in years:
        date = f"01.12.{year}"
        url = f"https://api.privatbank.ua/p24api/exchange_rates?json&date={date}"
        response = requests.get(url)
        exchange_data = response.json()
        rate = None  # Обнуляюмо змінну для поточного року
        for item in exchange_data['exchangeRate']:
            if item['currency'] == currency:
                rate = item['saleRateNB']
                break
            
        print(f'Processing {currency} for {date}: found rate {rate}')
        row.append(rate)
    print(f'{currency} exchange rate added: {row}')
    data.append(row)
       



# Задаємо команду, додає значення курсів валют відповідно до кількості валют та років
try:
    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body = {
            "valueInputOption": "USER_ENTERED",
            "data": [
                {
                    "range": f"A1:{chr(65+len(years))}{len(currencies)+1}",  # Діапазон клітинок, який ми оновлюємо
                    "majorDimension": "ROWS",  # Обновляем строки
                    "values": data
                }
            ]
        }
    ).execute()
    print(f"Updated {values.get('totalUpdatedCells')} cells")

except HttpError as err:
    print(f"An error occurred: {err}")