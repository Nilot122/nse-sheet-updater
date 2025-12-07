import requests
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import StringIO

# Fetch CSV from NSE
URL = "https://nsearchives.nseindia.com/content/equities/symbolchange.csv"

response = requests.get(URL, headers={"User-Agent": "Mozilla/5.0"})
response.raise_for_status()

csv_text = response.text
df = pd.read_csv(StringIO(csv_text))

# Google Sheets auth
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# service_account.json will be created by GitHub Actions from a secret
creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
client = gspread.authorize(creds)

# 🔴 EDIT THESE TWO:
SPREADSHEET_NAME = "Symbol Name Change Data"
WORKSHEET_NAME = "ImportData"   # e.g. "ImportData"

sheet = client.open(SPREADSHEET_NAME).worksheet(WORKSHEET_NAME)

# Clear and update sheet
sheet.clear()
sheet.update([df.columns.tolist()] + df.values.tolist())

print("NSE data updated successfully!")
