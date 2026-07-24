import io
import os
import re
import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from docx.shared import Inches

# Google API Libraries
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ==========================================
# 1. CONFIGURATION & DRIVE FOLDER MAPPING
# ==========================================
SYLLABUS_CODE = "9626"

# Google Drive Folder IDs mapped to your live Google Drive folders
FOLDER_IDS = {
    "theory": "1T1sIqRKxF5aO_r0sCyIVxidt0TyXOCcB",     # Theory Papers (P1 & P3)
    "practical": "1EWBiwjvTc12LVtyNi2V9P9RSr8d2vgq7",  # Practical Papers (P2 & P4)
    "zips": "1AsXq8TktyqajB7XTa9SQ5f85Pr6CQcFJ"          # Source Files (.zip)
}

# Local directories for mirroring files locally on the server
LOCAL_FOLDERS = {
    "theory": "9626_theory",
    "practical": "9626_practical",
    "zips": "9626_zips"
}

# Ensure local storage directories exist on server startup
for folder_path in LOCAL_FOLDERS.values():
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


# ==========================================
# 2. AUTOMATIC ROUTING & GOOGLE DRIVE API
# ==========================================
def determine_target_folder(filename: str) -> tuple[str, str]:
    """
    Analyzes the filename using Regular Expressions to determine target folder.
    Returns a tuple of (folder_key, folder_display_name).
    """
    filename_lower = filename.lower()
    
    # 1. Zip / Source Files
    if filename_lower.endswith(".zip") or "_sf_" in filename_lower:
        return "zips", "9626_zips (Source Files)"
        
    # 2. Practical Papers (Papers 02 and 04)
    if re.search(r'_(qp|ms)_0[24]\b', filename_lower):
        return "practical", "9626_practical (Papers 2 & 4)"
        
    # 3. Theory Papers (Papers 11, 12, 13, 31, 32, 33)
    if re.search(r'_(qp|ms)_(1[123]|3[123])\b', filename_lower):
        return "theory", "9626_theory (Papers 1 & 3)"
        
    return None, None


def build_drive_service():
    """
    Authenticates with Google Drive API using OAuth 2.0 User Credentials stored in Streamlit Secrets.
    """
    try:
        creds = Credentials(
            token=None,  # Auto-refreshes token when needed
            refresh_token=st.secrets["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=st.secrets["client_id"],
            client_secret=st.secrets["client_secret"]
        )
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"❌ Authentication Configuration Error: {e}")
        return None


def upload_file_to_drive(file_bytes, filename, folder_id, mime_type):
    """
    Uploads file binary stream directly to Google Drive using OAuth 2.0 User Credentials.
    """
    service = build_drive_service()
    if not service:
        return None

    try:
        file_stream = io.BytesIO(file_bytes)
        
        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }
        
        media = MediaIoBaseUpload(
            file_stream, 
            mimetype=mime_type, 
            resumable=True
        )

        uploaded_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()

        return uploaded_file

    except Exception as error:
        st.error(f"❌ Drive API Upload Failed: {error}")
        return None


def sync_drive_folder_to_local(folder_key: str) -> tuple[int, str]:
    """
    Queries Google Drive for a specific folder and downloads any files missing locally.
    Returns (downloaded_count, status_message).
    """
    service = build_drive_service()
    if not service:
        return 0, "Failed to authenticate with Google Drive."

    drive_folder_id = FOLDER_IDS[folder_key]
    local_path = LOCAL_FOLDERS[folder_key]

    try:
        # Query Google Drive for all non-trashed files in this target folder
        query = f"'{drive_folder_id}' in parents and trashed = false"
        results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
        drive_files = results.get('files', [])

        downloaded_count = 0

        for file_info in drive_files:
            file_name = file_info['name']
            file_id = file_info['id']
            local_file_path = os.path.join(local_path, file_name)

            # Download only if the file does not exist locally yet
            if not os.path.exists(local_file_path):
                request = service.files().get_media(fileId=file_id)
                with open(local_file_path, "wb") as f:
                    downloader = MediaIoBaseDownload(f, request)
                    done = False
                    while not done:
                        status, done = downloader.next_chunk()
                
                downloaded_count += 1

        return downloaded_count, f"Synced {downloaded_count} new file(s) for folder `{folder_key}`."

    except Exception as e:
        return 0, f"Sync error on folder `{folder_key}`: {e}"


# ==========================================
# 3. ADVANCED FLEXIBLE SEARCH ENGINE
# ==========================================
def search_pdfs(keyword_list, folder_path, allowed_variants):
    """
    Scans local PDF files for keywords using flexible, case-insensitive matching.
    Handles variations like 'VECTOR', 'Vector', 'vector', and plurals ('vectors').
    """
    results = []
    if not os.path.exists(folder_path):
        return results

    # Clean and prepare keyword list
    cleaned_keywords = [k.strip().lower() for k in keyword_list if k.strip()]
    if not cleaned_keywords:
        return results

    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            base_name = os.path.splitext(file)[0]
            
            # Filter allowed paper variants (e.g., _11, _12, _13, _02, _31, _04)
            is_valid_variant = any(base_name.endswith(f"_{variant}") for variant in allowed_variants)
            if not is_valid_variant:
                continue

            filepath = os.path.join(folder_path, file)
            try:
                doc = fitz.open(filepath)
                for page_num in range(len(doc)):
                    page_text = doc[page_num].get_text()
                    
                    # Check if ALL entered keywords match on this page
                    matches_all = True
                    for kw in cleaned_keywords:
                        escaped_kw = re.escape(kw)
                        
                        # Regex pattern: matches exact word OR common plurals (e.g. vector / vectors)
                        pattern = r'\b' + escaped_kw + r'(s|es)?\b'
                        
                        # Fallback to general substring match if regex boundary misses non-standard words
                        if not re.search(pattern, page_text, re.IGNORECASE) and kw not in page_text.lower():
                            matches_all = False
                            break
                    
                    if matches_all:
                        results.append({
                            "file": file,
                            "page": page_num,
                            "path": filepath,
                            "type": "QP" if "_qp_" in file else "MS"
                        })
                doc.close()
            except Exception:
                continue
                
    return results


# ==========================================
# 4. APP STATE INITIALIZATION
# ==========================================
if 'handout_basket' not in st.session_state:
    st.session_state.handout_basket = []
if 'theory_results' not in st.session_state:
    st.session_state.theory_results = []
if 'practical_results' not in st.session_state:
    st.session_state.practical_results = []


# ==========================================
# 5. STREAMLIT UI LAYOUT
# ==========================================
st.set_page_config(page_title="9626 IT Resource Platform", layout="wide")
st.title("PUSAT TINGKATAN ENAM SENGKURONG")
st.subheader("💻 9626 Information Technology PYP Resources")

# Sidebar - Handout Basket Status
with st.sidebar:
    st.header("Handout Basket Summary")
    st.metric(label="Saved Pages in Basket", value=len(st.session_state.handout_basket))
    if st.button("🗑️ Clear Basket"):
        st.session_state.handout_basket = []
        st.rerun()

# Application Navigation Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔍 Theory Search (P1 & P3)", 
    "⚙️ Practical Search (P2 & P4)", 
    "🛒 Collection of Added Handout", 
    "📦 Source Files (ZIP)", 
    "🔒 Admin & Sync Panel"
])


# --- TAB 1: THEORY SEARCH (P1 & P3) ---
with tab1:
    st.header("Search Theory Papers (Paper 1 & Paper 3)")
    st.caption("Variants: Paper 1 (11, 12, 13) | Paper 3 (31, 32, 33)")
    keyword_t1 = st.text_input("Enter Theory Keywords (e.g., 'Normalized', 'Relational Database', 'CSS')", key="t1_kw")

    if st.button("Search Theory Papers", type="primary"):
        if keyword_t1:
            with st.spinner("Scanning Theory PDFs..."):
                keywords = [k.strip() for k in keyword_t1.split(",") if k.strip()]
                theory_variants = ["11", "12", "13", "31", "32", "33"]
                st.session_state.theory_results = search_pdfs(keywords, LOCAL_FOLDERS["theory"], theory_variants)
        else:
            st.warning("Please enter a keyword.")

    if st.session_state.theory_results:
        st.write(f"Found **{len(st.session_state.theory_results)}** matching pages:")
        for idx, item in enumerate(st.session_state.theory_results):
            col1, col2 = st.columns([4, 1])
            doc_kind = "📝 Question Paper" if item["type"] == "QP" else "🔑 Marking Scheme"
            col1.write(f"📄 **{item['file']}** | {doc_kind} | (Page {item['page'] + 1})")
            if col2.button("➕ Add", key=f"add_t1_{idx}"):
                st.session_state.handout_basket.append(item)
                st.toast("Added to basket!")


# --- TAB 2: PRACTICAL SEARCH (P2 & P4) ---
with tab2:
    st.header("Search Practical Papers (Paper 2 & Paper 4)")
    st.caption("Variants: Paper 2 (02) | Paper 4 (04)")
    keyword_t2 = st.text_input("Enter Practical Keywords (e.g., 'Mail Merge', 'JavaScript', 'Vector Graphics')", key="t2_kw")

    if st.button("Search Practical Papers", type="primary"):
        if keyword_t2:
            with st.spinner("Scanning Practical PDFs..."):
                keywords = [k.strip() for k in keyword_t2.split(",") if k.strip()]
                practical_variants = ["02", "04"]
                st.session_state.practical_results = search_pdfs(keywords, LOCAL_FOLDERS["practical"], practical_variants)
        else:
            st.warning("Please enter a keyword.")

    if st.session_state.practical_results:
        st.write(f"Found **{len(st.session_state.practical_results)}** matching pages:")
        for idx, item in enumerate(st.session_state.practical_results):
            col1, col2 = st.columns([4, 1])
            doc_kind = "📝 Question Paper" if item["type"] == "QP" else "🔑 Marking Scheme"
            col1.write(f"📄 **{item['file']}** | {doc_kind} | (Page {item['page'] + 1})")
            if col2.button("➕ Add", key=f"add_t2_{idx}"):
                st.session_state.handout_basket.append(item)
                st.toast("Added to basket!")


# --- TAB 3: HANDOUT BASKET ---
with tab3:
    st.header("Worksheet / Handout Builder")
    if st.session_state.handout_basket:
        st.subheader(f"Selected Question/Answer Pages: {len(st.session_state.handout_basket)}")

        for idx, item in enumerate(st.session_state.handout_basket):
            st.write(f"{idx+1}. **{item['file']}** (Page {item['page'] + 1})")

        if st.button("🪄 Export Handout to Word Document", type="primary"):
            doc = Document()
            doc.add_heading(f'PTES {SYLLABUS_CODE} IT Handout', 0)

            for item in st.session_state.handout_basket:
                doc.add_heading(f"Source: {item['file']} (Page {item['page'] + 1})", level=2)
                pdf_doc = fitz.open(item['path'])
                page = pdf_doc.load_page(item['page'])
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_data = io.BytesIO(pix.tobytes("png"))
                doc.add_picture(img_data, width=Inches(6.5))
                doc.add_page_break()
                pdf_doc.close()

            target_filename = f"{SYLLABUS_CODE}_IT_Handout.docx"
            doc.save(target_filename)

            with open(target_filename, "rb") as f:
                st.download_button("📥 Click for final Download to Local Drive", f, file_name=target_filename)
    else:
        st.info("Your basket is empty. Add pages from Tab 1 or Tab 2.")


# --- TAB 4: SOURCE FILES (ZIP) ---
with tab4:
    st.header("Download Practical Source Files (ZIP)")
    c1, c2, c3 = st.columns(3)
    with c1:
        z_year = st.selectbox("Select Year", [str(y) for y in range(2026, 2018, -1)])
    with c2:
        z_session = st.selectbox("Select Session", ["March (m)", "June (s)", "Nov (w)"])
        session_code = z_session.split("(")[1].replace(")", "")
    with c3:
        z_paper = st.selectbox("Select Paper Component", ["02 (Paper 2)", "04 (Paper 4)"])
        paper_code = z_paper.split()[0]

    short_year = z_year[-2:]
    expected_zip_name = f"9626_{session_code}{short_year}_sf_{paper_code}.zip"
    zip_path = os.path.join(LOCAL_FOLDERS["zips"], expected_zip_name)

    st.markdown("---")
    if os.path.exists(zip_path):
        st.success(f"Found Source File: `{expected_zip_name}`")
        with open(zip_path, "rb") as zf:
            st.download_button(
                label=f"📦 Download {expected_zip_name}",
                data=zf,
                file_name=expected_zip_name,
                mime="application/zip"
            )
    else:
        st.warning(f"Source file `{expected_zip_name}` is not available locally in `{LOCAL_FOLDERS['zips']}`. Use the Admin Sync button to pull newly uploaded files from Drive.")


# --- TAB 5: ADMIN & SYNC PANEL ---
with tab5:
    st.header("Admin & Google Drive Sync Panel")

    # Fetch admin password securely from secrets (without hardcoded defaults)
    admin_password = st.secrets.get("ADMIN_PASSWORD")

    if not admin_password:
        st.error(
            "🚨 `ADMIN_PASSWORD` is not configured in your Streamlit Secrets. "
            "Please add `ADMIN_PASSWORD = 'your_password'` to your secrets configuration."
        )
    else:
        pwd = st.text_input("Enter Your Admin Password", type="password")

        if pwd == admin_password:
            st.success("Admin Access Granted")
            
            # --- BULK SYNC FROM GOOGLE DRIVE ---
            st.subheader("🔄 Bulk Sync with Google Drive")
            st.caption("Click below if tutors uploaded files directly into your Google Drive folders.")
            
            if st.button("🔄 Sync All Files from Google Drive", type="primary"):
                with st.spinner("Scanning Google Drive folders and downloading new files..."):
                    total_synced = 0
                    for f_key in ["theory", "practical", "zips"]:
                        count, msg = sync_drive_folder_to_local(f_key)
                        total_synced += count
                        st.info(msg)
                    
                    st.success(f"🎉 Sync Complete! **{total_synced}** new file(s) downloaded and ready for search!")

            st.markdown("---")
            
            # --- MANUAL SINGLE FILE UPLOAD ---
            st.subheader("📤 Single File Direct Upload")
            uploaded_file = st.file_uploader("Browse Past Paper PDF or Source File (ZIP)", type=["pdf", "zip"])

            if uploaded_file is not None:
                folder_key, folder_name = determine_target_folder(uploaded_file.name)

                if folder_key is None:
                    st.warning(
                        f"⚠️ Filename `{uploaded_file.name}` does not match Cambridge naming conventions. "
                        "Expected formats: `9626_m19_qp_12.pdf`, `9626_s20_qp_02.pdf`, or `9626_s20_sf_02.zip`."
                    )
                else:
                    st.info(f"🎯 Target Destination Detected: **{folder_name}**")

                    if st.button("🚀 Upload File to Drive & Mirror Locally"):
                        with st.spinner("Processing file..."):
                            file_bytes = uploaded_file.read()

                            # 1. Save locally for instant search access
                            local_dest_dir = LOCAL_FOLDERS[folder_key]
                            local_save_path = os.path.join(local_dest_dir, uploaded_file.name)
                            with open(local_save_path, "wb") as f:
                                f.write(file_bytes)
                            st.info(f"📁 Mirrored file locally to `{local_dest_dir}/{uploaded_file.name}`.")

                            # 2. Upload to Google Drive
                            target_drive_folder_id = FOLDER_IDS[folder_key]
                            drive_result = upload_file_to_drive(
                                file_bytes, 
                                uploaded_file.name, 
                                target_drive_folder_id, 
                                uploaded_file.type
                            )

                            if drive_result:
                                st.success(f"✅ Successfully uploaded `{uploaded_file.name}` to Google Drive!")
                                st.markdown(f"[🔗 View File in Google Drive]({drive_result.get('webViewLink')})")
                            else:
                                st.error("❌ Failed to upload to Google Drive. Check your connection or secrets.")


# ==========================================
# 6. FOOTER
# ==========================================
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; width: 100%;">
        <p style="font-size: 20px; font-weight: bold; margin-bottom: 5px;">✨ PTES 9626 Information Technology Resource Portal ✨</p>
        <p style="color: gray; font-size: 14px;">Creator: Miss Hajah Nurul Haziqah HN</p>
    </div>
    """,
    unsafe_allow_html=True
)
