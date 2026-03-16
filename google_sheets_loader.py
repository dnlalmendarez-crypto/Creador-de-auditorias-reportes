"""
Google Sheets integration module.
Connects to Google Sheets via service account or share link for reading audit data.
"""

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
import re
from typing import Optional


SCOPES_READ = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

SCOPES_READWRITE = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]


def connect_with_service_account(credentials_info: dict, write_access: bool = False) -> gspread.Client:
    """Connect to Google Sheets using a service account JSON credentials dict."""
    scopes = SCOPES_READWRITE if write_access else SCOPES_READ
    creds = Credentials.from_service_account_info(credentials_info, scopes=scopes)
    return gspread.authorize(creds)


def extract_sheet_id(url_or_id: str) -> str:
    """Extract the spreadsheet ID from a Google Sheets URL or return as-is if already an ID."""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url_or_id)
    if match:
        return match.group(1)
    return url_or_id.strip()


def load_sheet_as_dict(client: gspread.Client, spreadsheet_id: str) -> dict[str, pd.DataFrame]:
    """
    Load all worksheets from a Google Spreadsheet into a dict of DataFrames.
    Keys are sheet names, values are DataFrames.
    """
    spreadsheet = client.open_by_key(spreadsheet_id)
    result = {}
    for ws in spreadsheet.worksheets():
        data = ws.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            result[ws.title] = df
        elif len(data) == 1:
            df = pd.DataFrame(columns=data[0])
            result[ws.title] = df
        else:
            result[ws.title] = pd.DataFrame()
    return result


def load_sheet_first(client: gspread.Client, spreadsheet_id: str) -> pd.DataFrame | None:
    """Load only the first worksheet as a single DataFrame."""
    spreadsheet = client.open_by_key(spreadsheet_id)
    ws = spreadsheet.sheet1
    data = ws.get_all_values()
    if len(data) > 1:
        return pd.DataFrame(data[1:], columns=data[0])
    return None


def load_clasificacion_from_sheets(client: gspread.Client, spreadsheet_id: str) -> pd.DataFrame | None:
    """Load clasificacion de no conformidades from Google Sheets."""
    sid = extract_sheet_id(spreadsheet_id)
    return load_sheet_first(client, sid)


def load_tipificaciones_from_sheets(client: gspread.Client, spreadsheet_id: str) -> pd.DataFrame | None:
    """Load tipificaciones from Google Sheets."""
    sid = extract_sheet_id(spreadsheet_id)
    return load_sheet_first(client, sid)


def load_graficas_from_sheets(client: gspread.Client, spreadsheet_id: str) -> dict | None:
    """Load graficas de cumplimiento (all sheets) from Google Sheets."""
    sid = extract_sheet_id(spreadsheet_id)
    return load_sheet_as_dict(client, sid)


def load_base_datos_from_sheets(client: gspread.Client, spreadsheet_id: str) -> dict | None:
    """Load base de datos x cita (all sheets) from Google Sheets."""
    sid = extract_sheet_id(spreadsheet_id)
    return load_sheet_as_dict(client, sid)


def upload_file_to_drive(
    credentials_info: dict,
    file_bytes: bytes,
    filename: str,
    mime_type: str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    folder_id: str | None = None,
) -> str:
    """
    Upload a file to Google Drive using service account credentials.
    Returns the web view link of the uploaded file.
    """
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaInMemoryUpload

    creds = Credentials.from_service_account_info(credentials_info, scopes=SCOPES_READWRITE)
    service = build("drive", "v3", credentials=creds)

    file_metadata = {"name": filename}
    if folder_id:
        file_metadata["parents"] = [folder_id]

    media = MediaInMemoryUpload(file_bytes, mimetype=mime_type, resumable=False)

    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, webViewLink",
    ).execute()

    return uploaded.get("webViewLink", f"https://drive.google.com/file/d/{uploaded['id']}/view")


def extract_folder_id(url_or_id: str) -> str:
    """Extract folder ID from a Google Drive folder URL or return as-is."""
    match = re.search(r"/folders/([a-zA-Z0-9-_]+)", url_or_id)
    if match:
        return match.group(1)
    return url_or_id.strip()
