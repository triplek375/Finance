from google.oauth2 import service_account
import config
from googleapiclient.discovery import build
import gspread
import io
from googleapiclient.http import MediaIoBaseDownload

def get_services():
    creds = service_account.Credentials.from_service_account_file(
        config.KEY_FILE_PATH, scopes=config.SCOPES)
    drive_service = build('drive', 'v3', credentials=creds)
    sheet_service = build('sheets', 'v4', credentials=creds)
    gc = gspread.authorize(creds)
    return drive_service, sheet_service, gc

def get_folder_id(drive_service, parent_id, folder_name):
    """Finds a folder ID by name within a parent folder."""
    query = f"'{parent_id}' in parents and name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    if not files:
        raise FileNotFoundError(f"Folder '{folder_name}' not found inside parent {parent_id}")
    return files[0]['id']

def get_file_id(drive_service, parent_id, file_name):
    """Finds a file ID by name within a parent folder."""
    query = f"'{parent_id}' in parents and name = '{file_name}' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    if not files:
        raise FileNotFoundError(f"File '{file_name}' not found inside parent {parent_id}")
    return files[0]['id']

def list_files_in_folder(drive_service, folder_id):
    """Returns a list of files in a folder."""
    query = f"'{folder_id}' in parents and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    return results.get('files', [])

def download_file_content(drive_service, file_id):
    """Downloads file content to a BytesIO object (memory)."""
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh