import io
import streamlit as st
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

def build_drive_service():
    """
    Constructs an authenticated Google Drive API service client 
    using the refresh token stored inside Streamlit Secrets.
    """
    creds = Credentials(
        token=None,  # Access token auto-refreshes using refresh_token
        refresh_token=st.secrets["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=st.secrets["client_id"],
        client_secret=st.secrets["client_secret"]
    )
    return build('drive', 'v3', credentials=creds)

# --- STREAMLIT UI ---
st.set_page_config(page_title="IT PYP Portal", page_icon="📁")
st.title("📁 A-Level IT PYP Portal - File Uploader")
st.write("Upload past year papers and solutions directly to Google Drive.")

# File uploader widget
uploaded_file = st.file_uploader("Select a file to upload", type=["pdf", "docx", "zip", "png", "jpg"])

if uploaded_file is not None:
    if st.button("Upload to Google Drive"):
        try:
            st.info("Uploading file to Google Drive...")
            
            # Connect to Google Drive API
            service = build_drive_service()
            
            # Set metadata for uploaded file
            file_metadata = {'name': uploaded_file.name}
            
            # Convert uploaded bytes into media stream
            media = MediaIoBaseUpload(
                io.BytesIO(uploaded_file.getvalue()), 
                mimetype=uploaded_file.type,
                resumable=True
            )
            
            # Execute upload
            file_result = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, webViewLink'
            ).execute()
            
            st.success(f"✅ Success! File uploaded with ID: `{file_result.get('id')}`")
            st.markdown(f"[🔗 View Uploaded File in Google Drive]({file_result.get('webViewLink')})")
            
        except Exception as error:
            st.error(f"❌ Upload failed: {error}")
