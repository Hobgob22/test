import datetime
import requests
from google_sheets_api import GoogleSheetsAPI

CREDENTIALS_FILE = 'stylesheet-project-429720-8bb6dca344c7.json'

SPREADSHEET_ID = '1brK-jpAzLofA2IKRHj66AAEZZRZYBFVmsCvXyPevaGY'

google_sheets = GoogleSheetsAPI(CREDENTIALS_FILE, SPREADSHEET_ID)

today = datetime.date.today()
currencies = ["USD", "EUR", "CHF", "GBP", "PLN", "SEK", "ILS", "CAD"]
years = [i for i in range(today.year-4, today.year)]

data = [["Currency"] + years]

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

sheet_name = "PrivatAPI"


google_sheets.update_google_sheet(data, f"{sheet_name}!A1:{chr(65+len(years))}{len(currencies)+1}")