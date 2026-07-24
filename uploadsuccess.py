import io
import re
import streamlit as st
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ---------------------------------------------------------
# GOOGLE DRIVE FOLDER IDS CONFIGURATION
# Your live folder IDs mapped directly from Google Drive
# ---------------------------------------------------------
FOLDER_IDS = {
    "theory": "1T1sIqRKxF5aO_r0sCyIVxidt0TyXOCcB",     # 9626_theory (Papers 1 & 3)
    #1T1sIqRKxF5aO_r0sCyIVxidt0TyXOCcB
    "practical": "1EWBiwjvTc12LVtyNi2V9P9RSr8d2vgq7",  # 9626_practical (Papers 2 & 4)
    #1EWBiwjvTc12LVtyNi2V9P9RSr8d2vgq7
    "zips": "1AsXq8TktyqajB7XTa9SQ5f85Pr6CQcFJ"          # 9626_zips (Source/Zip files)
    #1AsXq8TktyqajB7XTa9SQ5f85Pr6CQcFJ
}

def determine_target_folder(filename: str) -> tuple[str, str]:
    """
    Analyzes the filename and returns a tuple of (folder_key, folder_display_name).
    
    Sorting Rules:
    - Ends with .zip OR contains '_sf_' -> zips folder
    - Contains _qp_02, _ms_02, _qp_04, _ms_04 -> practical folder
    - Contains _qp_11, _ms_12, _qp_31, _ms_33, etc. -> theory folder
    """
    filename_lower = filename.lower()
    
    # 1. Rule for ZIP / Source Files
    if filename_lower.endswith(".zip") or "_sf_" in filename_lower:
        return "zips", "9626_zips"
        
    # 2. Rule for Practical Papers (Paper 02 and Paper 04)
    if re.search(r'_(qp|ms)_0[24]\b', filename_lower):
        return "practical", "9626_practical"
        
    # 3. Rule for Theory Papers (Paper 11, 12, 13, 31, 32, 33)
    if re.search(r'_(qp|ms)_(1[123]|3[123])\b', filename_lower):
        return "theory", "9626_theory"
        
    # Default fallback if the filename format is unrecognized
    return None, None

def build_drive_service():
    """
    Constructs an authenticated Google Drive API client 
    using the OAuth secrets stored inside Streamlit Secrets.
    """
    creds = Credentials(
        token=None,  # Auto-refreshes using refresh_token
        refresh_token=st.secrets["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=st.secrets["client_id"],
        client_secret=st.secrets["client_secret"]
    )
    return build('drive', 'v3', credentials=creds)

# --- STREAMLIT USER INTERFACE ---
st.set_page_config(page_title="IT PYP Portal", page_icon="📁")
st.title("📁 A-Level IT PYP Portal - Auto Uploader")
st.write("Upload past year papers and source files. The system automatically routes the file to the correct folder based on its Cambridge code.")

# File uploader widget accepting PDF and ZIP formats
uploaded_file = st.file_uploader("Select a file to upload", type=["pdf", "zip"])

if uploaded_file is not None:
    # Run the filename analysis function
    folder_key, folder_name = determine_target_folder(uploaded_file.name)
    
    if folder_key is None:
        st.warning(
            f"⚠️ Filename `{uploaded_file.name}` does not match the standard pattern. "
            "Example expected formats: `9626_m19_qp_12.pdf`, `9626_s20_qp_02.pdf`, or `9626_s20_sf_02.zip`."
        )
    else:
        st.info(f"🎯 Target Folder Detected: **{folder_name}**")
        
        if st.button("Upload to Google Drive"):
            try:
                st.info("Uploading file to Google Drive...")
                
                # Fetch target folder ID using the key returned by the matcher
                target_folder_id = FOLDER_IDS[folder_key]
                
                # Build Google Drive API client
                service = build_drive_service()
                
                # Define file metadata with the parent folder destination
                file_metadata = {
                    'name': uploaded_file.name,
                    'parents': [target_folder_id]
                }
                
                # Prepare binary stream for upload
                media = MediaIoBaseUpload(
                    io.BytesIO(uploaded_file.getvalue()), 
                    mimetype=uploaded_file.type,
                    resumable=True
                )
                
                # Execute the upload request
                file_result = service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id, webViewLink'
                ).execute()
                
                st.success(f"✅ Successfully uploaded to **{folder_name}**!")
                st.markdown(f"[🔗 View File in Google Drive]({file_result.get('webViewLink')})")
                
            except Exception as error:
                st.error(f"❌ Upload failed: {error}")

