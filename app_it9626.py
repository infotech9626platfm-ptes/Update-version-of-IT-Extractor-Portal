import io
import streamlit as st
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

def get_drive_service():
    """
    Builds and returns a Google Drive API service object 
    using credentials stored in Streamlit secrets.
    """
    creds = Credentials(
        token=None,  # Access token auto-refreshes using the refresh_token
        refresh_token=st.secrets["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=st.secrets["client_id"],
        client_secret=st.secrets["client_secret"]
    )
    return build('drive', 'v3', credentials=creds)

st.title("📁 Google Drive File Uploader")

# File uploader widget in Streamlit
uploaded_file = st.file_uploader("Choose a file to upload to Google Drive")

if uploaded_file is not None:
    if st.button("Upload to Drive"):
        try:
            st.info("Uploading file to Google Drive...")
            
            # Initialize Drive API service
            service = get_drive_service()
            
            # Prepare file metadata
            file_metadata = {'name': uploaded_file.name}
            
            # Convert file bytes stream for Google API
            media = MediaIoBaseUpload(
                io.BytesIO(uploaded_file.getvalue()), 
                mimetype=uploaded_file.type,
                resumable=True
            )
            
            # Upload file
            uploaded_drive_file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, webViewLink'
            ).execute()
            
            st.success(f"✅ Upload Successful! File ID: {uploaded_drive_file.get('id')}")
            st.markdown(f"[View File on Google Drive]({uploaded_drive_file.get('webViewLink')})")
            
        except Exception as e:
            st.error(f"❌ Upload Failed: {e}")
