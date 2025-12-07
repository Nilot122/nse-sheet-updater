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

# 🔥 FIX: Replace all NaN/inf with empty strings
df = df.replace({pd.NA: "", float("nan"): "", pd.NaT: ""})
df = df.fillna("")  # extra safety

# Google Sheets auth
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
client = gspread.authorize(creds)

SPREADSHEET_NAME = "Symbol Name Change Data"
WORKSHEET_NAME = "ImportData"

sheet = client.open(SPREADSHEET_NAME).worksheet(WORKSHEET_NAME)

# Clear sheet and update
sheet.clear()
sheet.update([df.columns.tolist()] + df.values.tolist())

print("🎉 NSE Symbol Change sheet updated successfully!")
