import os
import re
import json
from datetime import datetime

import pymupdf
import pytesseract
from PIL import Image
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import gspread
from gspread.exceptions import APIError
from io import BytesIO

# Constants
# DRIVE_FOLDER_ID = '1A_91srZiGU1_VWjgcA3cBLyZElzk40SY'
DRIVE_FOLDER_ID = '1kJVADTGHJ-F4akQivAZoeGPxCdEV_clO'
SHEET_ID = '1SH5bjyi_dAZobEfAP5T8Cxc3B9jnkiFzmIRYGHL89x8'
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'
PORT = 8080


# Authentication
def authenticate_google_services():
    SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=PORT)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    return drive_service, sheets_service


# Fetch files from Google Drive folder
def fetch_drive_files(drive_service):
    results = drive_service.files().list(
        q=f"'{DRIVE_FOLDER_ID}' in parents",
        fields="files(id, name, mimeType)",
        pageSize=100
    ).execute()
    return results.get('files', [])


# Extract data from PDF
def extract_data_from_pdf(file_content):
    pdf_document = pymupdf.open(stream=file_content, filetype="pdf")
    first_page = pdf_document.load_page(0)
    file_text = first_page.get_text()
    return file_text


# Extract data from PNG
def extract_data_from_png(file_content):
    image = Image.open(file_content)
    file_text = pytesseract.image_to_string(image)
    return file_text


# Extract data from the file based on its type
def convert_to_mmddyyyy(text):
    # Regular expressions for different formats with word boundaries
    date_patterns = [
        r'\b(\d{4})-(\d{2})-(\d{2})\b',  # YYYY-MM-DD
        r'\b(\d{4})/(\d{2})/(\d{2})\b',  # YYYY/MM/DD
        r'\b(\d{2})/(\d{2})/(\d{4})\b',  # DD/MM/YYYY
        r'\b(\d{2})-(\d{2})-(\d{4})\b'   # DD-MM-YYYY
    ]

    # Search for all date patterns in the text (handling line breaks)
    for pattern in date_patterns:
        match = re.search(pattern, text, re.DOTALL)  # Using DOTALL to handle line breaks
        if match:
            groups = match.groups()
            # Handling different formats based on matching groups
            if len(groups[0]) == 4:  # YYYY first (formats YYYY-MM-DD and YYYY/MM/DD)
                year, month, day = groups
            else:  # DD first (formats DD/MM/YYYY and DD-MM-YYYY)
                day, month, year = groups
            # Convert and return in MM/DD/YYYY format
            return datetime(int(year), int(month), int(day)).strftime('%m/%d/%Y')

    return None  # Return None if no pattern matches


def extract_data_from_file(file_id, mime_type, drive_service):
    request = drive_service.files().get_media(fileId=file_id).execute()
    file_content = BytesIO(request)
    USD = False
    duties = False
    refund = False

    if mime_type == 'application/pdf':
        file_text = extract_data_from_pdf(file_content)
    elif mime_type == 'image/png':
        file_text = extract_data_from_png(file_content)
    else:
        raise ValueError(f"Unsupported file type: {mime_type}")

    print("Extracted Text:\n", file_text)

    formatted_date = convert_to_mmddyyyy(file_text)

    cost_match = re.search(r"\b(?:Total:?\s+%?|INR:?\s)(\d+,?\d+.?\d+|\d+,?\d+.?\d+\s\w+)\b", file_text)

    refund_match = re.search(r"Refund\s*([0-9,.]+)", file_text)


    duties_match = re.search(r"Duties,\s*", file_text)

    if duties_match:
        duties = True

    if refund_match:
        refund = True

    if cost_match is None:
        cost_match = re.search(r"order\samount\sUSD\s*(\d+,?\d+.?\d+|\d+,?\d+.?\d+\s\w+)", file_text)
        if cost_match:
            USD = True

    if cost_match is None:
        cost = ""
    else:
        cost = float(cost_match.group(1).replace(',', ''))
    return formatted_date, cost, refund, USD, duties


# Fetch existing entries from the Google Sheet
def fetch_existing_entries(sheets_service):
    sheet = sheets_service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range="expenses!A:C"
    ).execute()
    rows = sheet.get('values', [])
    # Return a set of invoice links already in the sheet
    return set(row[2] for row in rows if len(row) >= 3)


# Update Google Sheet
def update_google_sheet(sheets_service, data):
    sheet = sheets_service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range="expenses!A:D",
        valueInputOption="RAW",
        body={"values": data}
    ).execute()
    return sheet


# Main function
def main():
    drive_service, sheets_service = authenticate_google_services()

    # Fetch existing invoice links from the sheet
    existing_entries = fetch_existing_entries(sheets_service)

    files = fetch_drive_files(drive_service)
    data_to_insert = []

    for file in files:
        file_id = file.get('id')
        file_name = file.get('name')
        mime_type = file.get('mimeType')
        invoice_link = f"https://drive.google.com/file/d/{file_id}/view"

        # Skip if this file is already in the Google Sheet
        if invoice_link in existing_entries:
            print(f"Skipping file {file_name} as it is already in the sheet.")
            continue

        try:
            date, cost, refund, usd, duties_match = extract_data_from_file(file_id, mime_type, drive_service)
            data_to_insert.append([date, file_name, cost, invoice_link, refund, usd, duties_match])
        except Exception as e:
            print(f"Error processing file {file_name}: {e}")

    if data_to_insert:
        update_google_sheet(sheets_service, data_to_insert)
        print("Data has been updated in Google Sheets.")
    else:
        print("No data to update.")


if __name__ == "__main__":
    main()
