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

# CSV has NO HEADER → treat first row as data
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

# ---- Clear ONLY data rows (Row 3 onward) ----
last_row = len(df) + 5
sheet.batch_clear([f"A3:Z{last_row}"])

# ---- Row 1: Timestamp ----
ist = pytz.timezone("Asia/Kolkata")
timestamp = datetime.now(ist).strftime("%Y-%m-%d %H:%M:%S IST")
sheet.update("A1", [[f"Last Updated: {timestamp}"]])

# ---- Row 2: DO NOTHING (PRESERVE USER HEADER) ----

# ---- Row 3+: Write new data ----
sheet.update("A3", df.values.tolist())

print("🎉 Sheet updated successfully — Row 2 preserved!")
