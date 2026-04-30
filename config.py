# ============================================================
# config.py — Settings for the Email Bot
# ============================================================

# --- Google Sheet Settings ---
# To find your SPREADSHEET_ID, open your Google Sheet and look at the URL:
# https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_IS_HERE/edit
# Copy the long string between /d/ and /edit
SPREADSHEET_ID = "1GMTypqrtNcb5zrkEllEIz3hDvDcoMAlVKqrc65vuOwE"

# The name of the sheet tab (bottom of Google Sheets)
SHEET_NAME = "Sheet4"

# --- Column Mapping (match your sheet columns) ---
# These are the exact header names from row 1 of your sheet
COL_FIRST_NAME = "First Name"
COL_LAST_NAME = "Last Name"
COL_TITLE = "Title"
COL_COMPANY = "Company Name"
COL_EMAIL = "Email"
COL_SENT_STATUS = "Sent status"

# --- Email Settings ---
# Subject is now in template.txt (first line: "Subject: ...")

# Your name and email (used in "From" field)
SENDER_NAME = "Om Varma"
SENDER_EMAIL = "omvarma1731@gmail.com"

# --- Rate Limiting ---
# Delay between each email (in seconds) to avoid Gmail throttling
# Recommended: 5-10 seconds for safety
DELAY_BETWEEN_EMAILS = 5

# --- File Paths ---
TEMPLATE_FILE = "template.txt"
CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"

# --- Gmail API Scopes ---
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.settings.basic",
    "https://www.googleapis.com/auth/spreadsheets",
]
