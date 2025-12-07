import requests
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import StringIO
from datetime import datetime
import pytz

# Fetch CSV from NSE
URL = "https://nsearchives.nseindia.com/content/equities/symbolchange.csv"
response = requests.get(URL, headers={"User-Agent": "Mozilla/5.0"})
response.raise_for_status()

csv_text = response.text

# CSV has NO HEADER → set header=None
df = pd.read_csv(StringIO(csv_text), header=None)

# Clean NaN / invalid values
df = df.replace({pd.NA: "", float("nan"): "", pd.NaT: ""}).fillna("")

# Google Sheets Auth
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
client = gspread.authorize(creds)

SPREADSHEET_NAME = "Symbol Name Change Data"
WORKSHEET_NAME = "ImportData"

sheet = client.open(SPREADSHEET_NAME).worksheet(WORKSHEET_NAME)

# Clear sheet
sheet.clear()

# IST timestamp
ist = pytz.timezone("Asia/Kolkata")
timestamp = datetime.now(ist).strftime("%Y-%m-%d %H:%M:%S IST")

# -------- ROW 1: Timestamp --------
sheet.update("A1", [[f"Last Updated: {timestamp}"]])

# -------- ROW 2: EMPTY (you will add header manually) --------
sheet.update("A2", [["", "", "", "", ""]])

# -------- ROW 3+: DATA --------
data = df.values.tolist()
sheet.update("A3", data)

print("🎉 Sheet updated with timestamp and data (no header, no sorting).")
