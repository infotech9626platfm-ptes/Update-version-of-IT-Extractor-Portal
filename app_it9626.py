import io
import os
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ==========================================
# 1. TARGET GOOGLE DRIVE FOLDERS
# ==========================================
# Target Google Drive folder IDs
FOLDER_IDS = {
    "Theory": "1T1sIqRKxF5aO_r0sCyIVxidt0TyXOCcB",     # Theory Papers (P1 & P3)
    "Practical": "1EWBiwjvTc12LVtyNi2V9P9RSr8d2vgq7"   # Practical Papers (P2 & P4)
}

SCOPES = ['https://www.googleapis.com/auth/drive']


# ==========================================
# 2. AUTHENTICATION & UPLOAD FUNCTION
# ==========================================
def get_drive_service():
    """
    Loads Service Account credentials from st.secrets or local service_account.json.
    """
    if "gcp_service_account" in st.secrets:
        info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return build('drive', 'v3', credentials=creds)
    elif os.path.exists("service_account.json"):
        creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
        return build('drive', 'v3', credentials=creds)
    else:
        st.error("❌ Credentials missing! Ensure `service_account.json` exists locally or `gcp_service_account` is in Streamlit secrets.")
        return None


def upload_file_test(file_bytes, filename, folder_id):
    """
    Uploads a file directly to Google Drive using the Service Account.
    """
    service = get_drive_service()
    if not service:
        return None

    try:
        # Wrap raw file bytes into binary stream
        file_stream = io.BytesIO(file_bytes)
        
        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }
        
        # Set MIME type
        mime_type = 'application/pdf' if filename.endswith('.pdf') else 'application/zip'
        
        media = MediaIoBaseUpload(
            file_stream, 
            mimetype=mime_type, 
            resumable=False
        )

        # Send request with supportsAllDrives=True
        response = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, name, owners',
            supportsAllDrives=True
        ).execute()

        return response

    except Exception as error:
        st.error(f"❌ Upload Failed! Detailed Error:\n\n`{error}`")
        return None


# ==========================================
# 3. SIMPLE TEST USER INTERFACE
# ==========================================
st.set_page_config(page_title="Upload Tester", page_icon="🧪")
st.title("🧪 Isolated Upload Tester")
st.caption("Use this interface to test file uploads with your Service Account JSON credentials.")

st.markdown("---")

# Form inputs
col1, col2 = st.columns(2)

with col1:
    category = st.radio(
        "1. Select Paper Category:",
        ["Theory", "Practical"],
        help="Theory goes to P1/P3 folder. Practical goes to P2/P4 folder."
    )

with col2:
    doc_type = st.radio(
        "2. Select Document Type:",
        ["Question Paper (qp)", "Marking Scheme (ms)"]
    )

st.markdown("---")

# File browser
uploaded_file = st.file_uploader("3. Browse File from Hard Drive:", type=["pdf", "zip"])

# Action Button
if st.button("🚀 Upload to Google Drive", type="primary"):
    if uploaded_file:
        target_folder = FOLDER_IDS[category]
        st.info(f"Target Folder ID: `{target_folder}` ({category} Folder)")
        
        with st.spinner("Executing Google Drive API Upload request..."):
            raw_bytes = uploaded_file.read()
            result = upload_file_test(raw_bytes, uploaded_file.name, target_folder)

            if result:
                st.success(f"✅ SUCCESS! File uploaded successfully.")
                st.json(result)  # Displays file ID and owner details returned by Google
    else:
        st.warning("Please browse and select a file first.")
