import requests
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import StringIO
from datetime import datetime
import pytz
from gspread_formatting import sort_range

# Fetch CSV from NSE
URL = "https://nsearchives.nseindia.com/content/equities/symbolchange.csv"
response = requests.get(URL, headers={"User-Agent": "Mozilla/5.0"})
response.raise_for_status()

csv_text = response.text
df = pd.read_csv(StringIO(csv_text))

# Clean NaN / inf values
df = df.replace({pd.NA: "", float("nan"): "", pd.NaT: ""}).fillna("")

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

# Clear entire sheet
sheet.clear()

# Generate timestamp in IST
ist = pytz.timezone("Asia/Kolkata")
timestamp = datetime.now(ist).strftime("%Y-%m-%d %H:%M:%S IST")

# ---------- ROW 1: TIMESTAMP ----------
sheet.update("A1", [[f"Last Updated: {timestamp}"]])

# ---------- ROW 2: HEADER ----------
header = df.columns.tolist()
sheet.update("A2", [header])

# ---------- ROW 3+: DATA ----------
data = df.values.tolist()
sheet.update("A3", data)

print("Data written. Now applying sorting...")

# ---------- SORT BY COLUMN D (4th column) Z → A ----------
# We sort the range A3:E (all data, excluding timestamp & header)
sheet.sort((4, "des"))  # Column 4, descending

print("🎉 NSE Symbol Change sheet updated with timestamp, header, and sorted by column D (Z → A)!")
