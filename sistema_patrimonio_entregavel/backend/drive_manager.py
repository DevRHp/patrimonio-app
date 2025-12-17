import os
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'credentials.json'

def authenticate():
    creds = None
    if os.path.exists(SERVICE_ACCOUNT_FILE):
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return creds

def get_service():
    creds = authenticate()
    if not creds:
        return None
    return build('drive', 'v3', credentials=creds)

def create_folder(service, name, parent_id=None):
    file_metadata = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    if parent_id:
        file_metadata['parents'] = [parent_id]
    
    file = service.files().create(body=file_metadata, fields='id, webViewLink').execute()
    return file

def find_folder(service, name, parent_id=None):
    query = f"mimeType='application/vnd.google-apps.folder' and name='{name}' and trashed=false"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    
    results = service.files().list(q=query, fields="files(id, webViewLink)").execute()
    files = results.get('files', [])
    if files:
        return files[0]
    return None

def upload_file(service, filename, filepath, folder_id):
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    media = MediaFileUpload(filepath, resumable=True)
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return file

def upload_audit_results(analyst_name, room_name, files_map):
    """
    files_map: { 'filename': 'absolute_filepath' }
    Returns: webViewLink of the folder
    """
    service = get_service()
    if not service:
        # If no credentials, return None to signal failure/skip
        print("No Drive credentials found.")
        return None

    # 1. Root Folder
    root_name = "Auditorias_Patrimonio"
    root_folder = find_folder(service, root_name)
    if not root_folder:
        root_folder = create_folder(service, root_name)
    
    # 2. Audit Folder
    date_str = datetime.datetime.now().strftime("%Y-%m-%d")
    safe_room = "".join([c for c in room_name if c.isalnum() or c in (' ','-','_')]).strip()
    audit_folder_name = f"{date_str} - {analyst_name} - {safe_room}"
    
    # Check if exists (optional), or just create new one
    audit_folder = create_folder(service, audit_folder_name, parent_id=root_folder['id'])
    
    # 3. Upload Files
    for filename, filepath in files_map.items():
        if os.path.exists(filepath):
            upload_file(service, filename, filepath, audit_folder['id'])
            
    return audit_folder.get('webViewLink')
