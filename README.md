# Email Bot — Gmail Mail Merge from Google Sheets

A Python bot that sends personalized emails via Gmail. Reads contacts from Google Sheets, fills in a template, sends emails, and marks sent status back in the sheet. Completely free — no paid tools needed.

---

## Features

- Reads contacts from Google Sheets (Name, Company, Email, Title)
- Sends personalized emails using a local `template.txt` file
- Marks each row as `SENT` in the sheet after sending
- Skips already-sent rows on next run
- 3 sending modes: Sheet, Manual, Quick Fire
- Sends as HTML with proper formatting
- 0 cost — uses Gmail API (free tier: 500 emails/day)

---

## Setup (One Time)

### 1. Google Cloud Setup

1. Go to [https://console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (e.g. `Email Bot`)
3. Enable **Gmail API** and **Google Sheets API**
4. Go to **OAuth consent screen** → External → Add your Gmail as a test user
5. Go to **Credentials** → Create OAuth Client ID → Desktop App
6. Download the JSON → rename to `credentials.json` → place in this folder

> Full step-by-step in `setup_guide.txt`

### 2. Link Your Google Sheet

Open your Google Sheet. Copy the ID from the URL:
```
https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_ID/edit
```
Paste it in `config.py`:
```python
SPREADSHEET_ID = "your-sheet-id-here"
```

Also set the tab name:
```python
SHEET_NAME = "Sheet1"  # change to your tab name
```

Your sheet must have these column headers (row 1):
```
First Name | Last Name | Title | Company Name | Email | Sent status
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Edit Your Email Template

Open `template.txt` and write your email. Use these placeholders:

```
{greeting}       →  Hi Priya,  (or Hi Hiring Team, for generic emails)
{company_name}   →  Google
{first_name}     →  Priya
{last_name}      →  Sharma
{title}          →  Talent Acquisition Partner
```

First line must be the subject:
```
Subject: Exploring SDE Fresher & Internship Opportunities at {company_name}
```

---

## Running the Bot

```bash
start
```

Or if `start` alias isn't set up:
```bash
/usr/bin/python3 mailer.py
```

---

## Modes

When you run the bot, you'll see:

```
What would you like to do?

  [1] Send to specific email IDs (manual entry)
  [2] Send from Google Sheet (bulk mail merge)
  [3] Quick fire (just paste: name, company, email)
  [4] Exit
```

### Mode 1 — Manual Entry
Enter how many recipients, then fill in details one by one:
```
How many recipients? 2

--- Recipient 1/2 ---
  First Name : Priya
  Last Name  : Sharma
  Title      : HR Manager
  Company    : Google
  Email      : priya@google.com
  Added!
```

### Mode 2 — Send from Google Sheet
Bot reads your sheet, shows unsent count, asks how many to send:
```
Total contacts: 204
Already sent:   45
Unsent:         159

How many emails do you want to send? (max 159)
Enter count (or 'all'): 10
```
After sending, each row is marked `SENT (2026-04-30 10:45)` in the sheet.

### Mode 3 — Quick Fire
Fastest mode. Just paste one per line as `FirstName, Company, Email`:
```
> Priya, Google, priya@google.com
> Rahul, Microsoft, rahul@microsoft.com
> done
```

---

## File Structure

```
email-bot/
├── mailer.py          # Main bot script
├── config.py          # Settings (Sheet ID, column names, delay)
├── template.txt       # Your email template (edit anytime)
├── requirements.txt   # Python dependencies
├── setup_guide.txt    # Detailed Google Cloud setup steps
├── start              # Terminal shortcut (run with: start)
├── .gitignore         # Keeps credentials out of git
└── credentials.json   # ← Download from Google Cloud (NOT committed)
```

---

## Notes

- First run opens browser for Google sign-in (one time only)
- Token is saved in `token.json` — delete it to re-authenticate
- Gmail free accounts: 500 emails/day limit
- Delay between emails: 5 seconds (configurable in `config.py`)
- `credentials.json` and `token.json` are in `.gitignore` — never committed
