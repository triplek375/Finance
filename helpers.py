import io
from googleapiclient.http import MediaIoBaseDownload

def get_folder_id(parent_id, folder_name):
    """Finds a folder ID by name within a parent folder."""
    query = f"'{parent_id}' in parents and name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    if not files:
        raise FileNotFoundError(f"Folder '{folder_name}' not found inside parent {parent_id}")
    return files[0]['id']

def get_file_id(parent_id, file_name):
    """Finds a file ID by name within a parent folder."""
    query = f"'{parent_id}' in parents and name = '{file_name}' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    if not files:
        raise FileNotFoundError(f"File '{file_name}' not found inside parent {parent_id}")
    return files[0]['id']

def list_files_in_folder(folder_id):
    """Returns a list of files in a folder."""
    query = f"'{folder_id}' in parents and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    return results.get('files', [])

def download_file_content(file_id):
    """Downloads file content to a BytesIO object (memory)."""
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def get_root_finance_structure():
    """
    Traverses from 'Shared with me' to find /Finance/Zerodha paths.
    You might need to adjust the logic if 'Finance' is in the root or shared.
    """
    # Find 'Finance' folder. If it's shared, we search without a parent first
    # or assume it is in the root accessible to the service account.
    # Note: Service accounts have their own 'root'. You must share the folder with the SA email.
    
    # Search for "Finance" globally
    results = drive_service.files().list(
        q="name = 'Finance' and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
        fields="files(id, name)"
    ).execute()
    
    if not results.get('files'):
        raise Exception("Could not find 'Finance' folder. Did you share it with the Service Account email?")
    
    finance_id = results['files'][0]['id']
    zerodha_id = get_folder_id(finance_id, 'Zerodha')
    
    return {
        'finance': finance_id,
        'zerodha': zerodha_id,
        'dividend': get_folder_id(zerodha_id, 'Dividend Statement'),
        'contract': get_folder_id(zerodha_id, 'Contract Note')
    }