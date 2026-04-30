#!/usr/bin/env python3
"""
Email Bot — Send personalized emails via Gmail.
Option 1: Manual — enter recipient details directly.
Option 2: Sheet  — read contacts from Google Sheets & track sent status.
"""

import base64
import os
import sys
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

import config


# ─── Authentication ──────────────────────────────────────────────────────────

def authenticate():
    """Authenticate with Google APIs (Gmail + Sheets). Opens browser on first run."""
    creds = None

    if os.path.exists(config.TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(config.TOKEN_FILE, config.SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired token...")
            creds.refresh(Request())
        else:
            if not os.path.exists(config.CREDENTIALS_FILE):
                print(f"ERROR: '{config.CREDENTIALS_FILE}' not found!")
                print("Please download it from Google Cloud Console.")
                print("See setup_guide.txt for instructions.")
                sys.exit(1)
            print("Opening browser for Google sign-in...")
            flow = InstalledAppFlow.from_client_secrets_file(config.CREDENTIALS_FILE, config.SCOPES)
            creds = flow.run_local_server(port=0)

        with open(config.TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
        print("Authentication successful! Token saved.\n")

    return creds


def get_sender_from_header(gmail_service):
    """Fetch the display name from Gmail's sendAs settings.

    Returns the From header string like: Om Varma <omvarma1731@gmail.com>
    This uses Gmail's OWN configured name so it passes authentication.
    """
    try:
        result = gmail_service.users().settings().sendAs().list(userId="me").execute()
        for alias in result.get("sendAs", []):
            if alias.get("isPrimary"):
                name = alias.get("displayName", "")
                email = alias.get("sendAsEmail", "")
                if name:
                    return f"{name} <{email}>"
                return email
    except Exception:
        pass
    return None


# ─── Google Sheets ───────────────────────────────────────────────────────────

def get_sheet_data(creds):
    """Read all rows from the Google Sheet and return (service, headers, rows)."""
    service = build("sheets", "v4", credentials=creds)

    result = service.spreadsheets().values().get(
        spreadsheetId=config.SPREADSHEET_ID,
        range=config.SHEET_NAME,
    ).execute()

    values = result.get("values", [])
    if not values:
        print("ERROR: Sheet is empty!")
        sys.exit(1)

    headers = values[0]
    rows = values[1:]

    return service, headers, rows


def update_sent_status(service, row_index, status_col_index):
    """Mark a row's Sent status as 'SENT' with today's date."""
    cell_range = f"{config.SHEET_NAME}!{col_letter(status_col_index)}{row_index + 2}"
    today = datetime.now().strftime("%Y-%m-%d %H:%M")
    value = f"SENT ({today})"

    service.spreadsheets().values().update(
        spreadsheetId=config.SPREADSHEET_ID,
        range=cell_range,
        valueInputOption="RAW",
        body={"values": [[value]]},
    ).execute()


def col_letter(index):
    """Convert 0-based column index to letter (0=A, 1=B, ..., 25=Z)."""
    return chr(65 + index)


# ─── Gmail ───────────────────────────────────────────────────────────────────

def text_to_html(text):
    """Convert plain text template to formatted HTML with proper paragraphs."""
    paragraphs = text.split("\n\n")
    html_parts = []
    for p in paragraphs:
        p = p.strip()
        if not p:
            continue
        # Convert single newlines within a paragraph to <br>
        p = p.replace("\n", "<br>")
        html_parts.append(f"<p style=\"margin:0 0 12px 0;font-size:14px;line-height:1.6;color:#222;\">{p}</p>")
    return f"""<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;">
{''.join(html_parts)}
</div>"""


def send_email(gmail_service, to_email, subject, body_text, from_header=None):
    """Send an HTML email via Gmail API with proper sender name."""
    html_body = text_to_html(body_text)

    message = MIMEMultipart("alternative")
    message["to"] = to_email
    message["subject"] = subject
    if from_header:
        message["from"] = from_header

    # Attach both plain text (fallback) and HTML
    message.attach(MIMEText(body_text, "plain"))
    message.attach(MIMEText(html_body, "html"))

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return gmail_service.users().messages().send(
        userId="me",
        body={"raw": raw},
    ).execute()


# ─── Template ────────────────────────────────────────────────────────────────

def load_template():
    """Load subject and body from template.txt.

    First line: Subject: ...
    Then blank line, then body.
    """
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.TEMPLATE_FILE)
    if not os.path.exists(template_path):
        print(f"ERROR: '{config.TEMPLATE_FILE}' not found!")
        print("Please create it with a Subject: line and body with placeholders.")
        sys.exit(1)

    with open(template_path, "r") as f:
        content = f.read()

    lines = content.split("\n", 1)
    first_line = lines[0].strip()

    if not first_line.lower().startswith("subject:"):
        print("ERROR: First line of template.txt must start with 'Subject:'")
        print(f"Got: {first_line}")
        sys.exit(1)

    subject_template = first_line[len("Subject:"):].strip()
    body_template = lines[1].lstrip("\n") if len(lines) > 1 else ""

    return subject_template, body_template


GENERIC_NAMES = {
    "hr", "careers", "career", "recruiter", "recruitment", "jobs", "job",
    "talent", "info", "admin", "tech", "hello", "people", "applications",
    "application", "filemakr", "gridebwork", "nezuware", "techersitfirm",
    "businesswithanu", "nityasarga",
}

def build_greeting(first_name, last_name):
    """Return 'Hi Hiring Team,' for generic emails, 'Hi FirstName,' for real names."""
    name = first_name.strip().lower()
    if name in GENERIC_NAMES or not name:
        return "Hi Hiring Team,"
    if last_name and last_name.strip():
        return f"Hi {first_name.strip()} {last_name.strip()},"
    return f"Hi {first_name.strip()},"


def fill_template(template, row_data):
    """Replace placeholders in template with actual values from the row."""
    data = dict(row_data)
    data["greeting"] = build_greeting(data.get("first_name", ""), data.get("last_name", ""))
    return template.format(**data)


# ─── Send Emails Helper ─────────────────────────────────────────────────────

def send_batch(gmail_service, recipients, subject_template, body_template, from_header=None):
    """Send emails to a list of recipients. Returns (success, failed) counts.

    Each recipient is a dict with: first_name, last_name, title, company_name, email
    """
    total = len(recipients)
    success = 0
    failed = 0

    for idx, r in enumerate(recipients):
        row_data = {
            "first_name": r["first_name"],
            "last_name": r["last_name"],
            "title": r["title"],
            "company_name": r["company_name"],
        }

        subject = subject_template.format(**row_data)
        body = fill_template(body_template, row_data)

        print(f"  [{idx + 1}/{total}] Sending to {r['first_name']} {r['last_name']} ({r['email']})...", end=" ")

        try:
            send_email(gmail_service, r["email"], subject, body, from_header)
            print("SENT")
            success += 1
        except Exception as e:
            print(f"FAILED — {e}")
            failed += 1

        # Rate limit delay (skip after last email)
        if idx < total - 1:
            time.sleep(config.DELAY_BETWEEN_EMAILS)

    return success, failed


# ─── Option 1: Manual Entry ─────────────────────────────────────────────────

def option_manual(gmail_service, from_header=None):
    """Send emails to manually entered recipients. Asks count first."""
    print()
    print("=" * 60)
    print("  OPTION 1 — Send to Specific Email IDs")
    print("=" * 60)
    print()

    # Ask how many first
    while True:
        try:
            total = int(input("How many recipients? ").strip())
            if total >= 1:
                break
            print("Enter at least 1.")
        except ValueError:
            print("Enter a number.")

    subject_template, body_template = load_template()
    recipients = []

    print(f"\nEnter details for {total} recipient(s):\n")

    for i in range(total):
        print(f"--- Recipient {i + 1}/{total} ---")
        first_name = input("  First Name : ").strip()
        last_name = input("  Last Name  : ").strip()
        title = input("  Title      : ").strip()
        company_name = input("  Company    : ").strip()
        email = input("  Email      : ").strip()

        if not email or "@" not in email:
            print("  Invalid email, skipping.\n")
            continue

        recipients.append({
            "first_name": first_name,
            "last_name": last_name,
            "title": title,
            "company_name": company_name,
            "email": email,
        })
        print(f"  Added!\n")

    if not recipients:
        print("No valid recipients. Going back.\n")
        return

    # Preview
    print(f"\nRecipients ({len(recipients)}):")
    for i, r in enumerate(recipients, 1):
        print(f"  {i}. {r['first_name']} {r['last_name']} — {r['company_name']} — {r['email']}")

    r0 = recipients[0]
    preview_data = {
        "first_name": r0["first_name"],
        "last_name": r0["last_name"],
        "title": r0["title"],
        "company_name": r0["company_name"],
    }
    print(f"\nSubject: {subject_template.format(**preview_data)}")
    print(f"Body preview:\n{fill_template(body_template, preview_data)[:300]}...")
    print()

    confirm = input("Send emails? (yes/no): ").strip().lower()
    if confirm not in ("yes", "y"):
        print("Cancelled.\n")
        return

    print(f"\nSending {len(recipients)} emails...\n")
    success, failed = send_batch(gmail_service, recipients, subject_template, body_template, from_header)

    print()
    print("=" * 60)
    print(f"  DONE! Sent: {success} | Failed: {failed}")
    print("=" * 60)
    print()


# ─── Option 4: Quick Fire ───────────────────────────────────────────────────

def option_quick(gmail_service, from_header=None):
    """Fast mode — paste firstname, company, email per line separated by comma."""
    print()
    print("=" * 60)
    print("  OPTION 4 — Quick Fire Mode")
    print("=" * 60)
    print()
    print("Paste recipients, one per line:")
    print("  Format: FirstName, CompanyName, Email")
    print("  Example: Priya, Google, priya@google.com")
    print()
    print("Type 'done' when finished.\n")

    subject_template, body_template = load_template()
    recipients = []

    while True:
        line = input("  > ").strip()
        if line.lower() == "done":
            break
        if not line:
            continue

        parts = [p.strip() for p in line.split(",")]
        if len(parts) != 3:
            print("    Invalid format. Use: FirstName, CompanyName, Email")
            continue

        first_name, company_name, email = parts

        if not email or "@" not in email:
            print(f"    Invalid email: {email}")
            continue

        recipients.append({
            "first_name": first_name,
            "last_name": "",
            "title": "",
            "company_name": company_name,
            "email": email,
        })
        print(f"    Added! ({len(recipients)} so far)")

    if not recipients:
        print("No recipients. Going back.\n")
        return

    # Preview
    print(f"\n{len(recipients)} recipient(s):")
    for i, r in enumerate(recipients, 1):
        print(f"  {i}. {r['first_name']} — {r['company_name']} — {r['email']}")

    print()
    confirm = input("Send emails? (yes/no): ").strip().lower()
    if confirm not in ("yes", "y"):
        print("Cancelled.\n")
        return

    print(f"\nSending {len(recipients)} emails...\n")
    success, failed = send_batch(gmail_service, recipients, subject_template, body_template, from_header)

    print()
    print("=" * 60)
    print(f"  DONE! Sent: {success} | Failed: {failed}")
    print("=" * 60)
    print()


# ─── Option 2: From Google Sheet ────────────────────────────────────────────

def option_sheet(gmail_service, creds, from_header=None):
    """Send emails from Google Sheet and update sent status."""
    print()
    print("=" * 60)
    print("  OPTION 2 — Send from Google Sheet")
    print("=" * 60)
    print()

    print("Reading data from Google Sheet...")
    sheets_service, headers, rows = get_sheet_data(creds)

    # Find column indices
    try:
        idx_first_name = headers.index(config.COL_FIRST_NAME)
        idx_last_name = headers.index(config.COL_LAST_NAME)
        idx_title = headers.index(config.COL_TITLE)
        idx_company = headers.index(config.COL_COMPANY)
        idx_email = headers.index(config.COL_EMAIL)
        idx_sent = headers.index(config.COL_SENT_STATUS)
    except ValueError as e:
        print(f"ERROR: Column not found in sheet — {e}")
        print(f"Your sheet headers are: {headers}")
        print("Update config.py column names to match your sheet.")
        return

    # Filter unsent rows
    unsent = []
    for i, row in enumerate(rows):
        while len(row) <= idx_sent:
            row.append("")

        sent_value = row[idx_sent].strip().upper()
        if not sent_value.startswith("SENT"):
            email = row[idx_email].strip() if len(row) > idx_email else ""
            if email:
                unsent.append((i, row))

    total_rows = len(rows)
    sent_count = total_rows - len(unsent)
    print(f"   Total contacts: {total_rows}")
    print(f"   Already sent:   {sent_count}")
    print(f"   Unsent:         {len(unsent)}")
    print()

    if not unsent:
        print("All emails have been sent! Nothing to do.\n")
        return

    # Ask how many to send
    print(f"How many emails do you want to send? (max {len(unsent)})")
    while True:
        try:
            count_input = input("Enter count (or 'all'): ").strip()
            if count_input.lower() == "all":
                count = len(unsent)
            else:
                count = int(count_input)
            if 1 <= count <= len(unsent):
                break
            print(f"Please enter a number between 1 and {len(unsent)}")
        except ValueError:
            print("Invalid input. Enter a number or 'all'.")

    # Load template
    subject_template, body_template = load_template()

    # Preview first email
    first_row = unsent[0][1]
    preview_data = {
        "first_name": first_row[idx_first_name].strip(),
        "last_name": first_row[idx_last_name].strip(),
        "title": first_row[idx_title].strip(),
        "company_name": first_row[idx_company].strip(),
    }
    print(f"\nReady to send {count} emails.")
    print(f"Delay between emails: {config.DELAY_BETWEEN_EMAILS}s")
    print()
    print("--- Preview (first email) ---")
    print(f"To:      {first_row[idx_email].strip()}")
    print(f"Subject: {subject_template.format(**preview_data)}")
    print(f"Body:\n{fill_template(body_template, preview_data)[:300]}...")
    print("-----------------------------\n")

    confirm = input("Send emails? (yes/no): ").strip().lower()
    if confirm not in ("yes", "y"):
        print("Cancelled.\n")
        return

    # Send emails
    print(f"\nSending {count} emails...\n")
    success = 0
    failed = 0

    for idx in range(count):
        row_index, row = unsent[idx]
        first_name = row[idx_first_name].strip()
        last_name = row[idx_last_name].strip()
        title = row[idx_title].strip()
        company = row[idx_company].strip()
        email = row[idx_email].strip()

        row_data = {
            "first_name": first_name,
            "last_name": last_name,
            "title": title,
            "company_name": company,
        }

        subject = subject_template.format(**row_data)
        body = fill_template(body_template, row_data)

        print(f"  [{idx + 1}/{count}] Sending to {first_name} {last_name} ({email})...", end=" ")

        try:
            send_email(gmail_service, email, subject, body, from_header)
            update_sent_status(sheets_service, row_index, idx_sent)
            print("SENT")
            success += 1
        except Exception as e:
            print(f"FAILED — {e}")
            failed += 1

        if idx < count - 1:
            time.sleep(config.DELAY_BETWEEN_EMAILS)

    print()
    print("=" * 60)
    print(f"  DONE! Sent: {success} | Failed: {failed}")
    print("=" * 60)
    print()


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  EMAIL BOT — Mail Merge Tool")
    print("=" * 60)
    print()

    # Authenticate
    print("Authenticating with Google...")
    creds = authenticate()
    gmail_service = build("gmail", "v1", credentials=creds)

    # Fetch sender From header from Gmail's own sendAs settings
    from_header = get_sender_from_header(gmail_service)
    if from_header:
        print(f"Sending as: {from_header}\n")

    # Main loop — keep showing menu until user exits
    while True:
        print("What would you like to do?\n")
        print("  [1] Send to specific email IDs (manual entry)")
        print("  [2] Send from Google Sheet (bulk mail merge)")
        print("  [3] Quick fire (just paste: name, company, email)")
        print("  [4] Exit")
        print()

        choice = input("Enter choice (1/2/3/4): ").strip()

        if choice == "1":
            option_manual(gmail_service, from_header)
        elif choice == "2":
            option_sheet(gmail_service, creds, from_header)
        elif choice == "3":
            option_quick(gmail_service, from_header)
        elif choice == "4":
            print("\nGoodbye!")
            break
        else:
            print("Invalid choice. Enter 1, 2, 3, or 4.\n")


if __name__ == "__main__":
    main()
