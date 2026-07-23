import os
import io
import json
import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO

# Google Drive API Libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

# --- 1. IT CONFIGURATION (9626) ---
SYLLABUS_CODE = "9626"

# Folder directory mapping
FOLDERS = {
    "theory": "9626_theory",
    "practical": "9626_practical",
    "zips": "9626_zips"
}

# Ensure local directories exist
for folder_path in FOLDERS.values():
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

# --- 2. GOOGLE DRIVE API HELPER FUNCTIONS ---
SCOPES = ['https://www.googleapis.com/auth/drive']

def get_drive_service():
    """Authenticates using Streamlit Secrets or local service_account.json."""
    if "gcp_service_account" in st.secrets:
        info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    elif os.path.exists("service_account.json"):
        creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
    else:
        return None
    return build('drive', 'v3', credentials=creds)

def upload_to_drive(file_bytes, filename, folder_id):
    """Uploads a file directly to a specific Google Drive folder."""
    service = get_drive_service()
    if not service:
        st.error("Google Drive API Credentials not found.")
        return False
    
    # Save temporary file locally for upload stream
    temp_path = os.path.join("temp_upload", filename)
    os.makedirs("temp_upload", exist_ok=True)
    with open(temp_path, "wb") as f:
        f.write(file_bytes)

    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    media = MediaFileUpload(temp_path, resumable=True)
    
    uploaded_file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    if os.path.exists(temp_path):
        os.remove(temp_path)
        
    return uploaded_file.get('id')

# --- 3. SEARCH & HELPER FUNCTIONS ---
def search_pdfs(keyword_list, folder_path, allowed_papers):
    """Scans PDFs in a folder matching specified keywords and paper filters."""
    results = []
    if not os.path.exists(folder_path):
        return results
        
    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            # Check paper filtering based on filename (e.g., 9626_s22_qp_12.pdf)
            is_valid_paper = any(f"_{p}" in file or f"_{p}." in file for p in allowed_papers)
            if not is_valid_paper:
                continue

            filepath = os.path.join(folder_path, file)
            try:
                doc = fitz.open(filepath)
                for page_num in range(len(doc)):
                    text = doc[page_num].get_text().lower()
                    if all(k.lower() in text for k in keyword_list if k.strip()):
                        results.append({
                            "file": file,
                            "page": page_num,
                            "path": filepath,
                            "type": "QP" if "_qp_" in file else "MS"
                        })
                doc.close()
            except Exception as e:
                continue
    return results

# --- 4. APP STATE INITIALIZATION ---
if 'handout_basket' not in st.session_state:
    st.session_state.handout_basket = []
if 'theory_results' not in st.session_state:
    st.session_state.theory_results = []
if 'practical_results' not in st.session_state:
    st.session_state.practical_results = []

# --- 5. UI LAYOUT ---
st.set_page_config(page_title="9626 IT Resource Platform", layout="wide")
st.title("PUSAT TINGKATAN ENAM SENGKURONG")
st.subheader("💻 9626 Information Technology PYP Platform")

# Sidebar
with st.sidebar:
    st.header("Basket Summary")
    st.metric(label="Saved Pages in Basket", value=len(st.session_state.handout_basket))
    if st.button("🗑️ Clear Basket"):
        st.session_state.handout_basket = []
        st.rerun()

# Tabs Definition
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔍 Theory Search (P1 & P3)", 
    "⚙️ Practical Search (P2 & P4)", 
    "🛒 Handout Basket", 
    "📦 Source Files (ZIP)", 
    "🔒 Admin Panel"
])

# --- TAB 1: THEORY SEARCH (P1 & P3) ---
with tab1:
    st.header("Search Theory Papers (Paper 1 & Paper 3)")
    keyword_t1 = st.text_input("Enter Theory Keywords (e.g., 'Normalized', 'Relational Database', 'CSS')", key="t1_kw")
    
    if st.button("Search Theory Papers", type="primary"):
        if keyword_t1:
            with st.spinner("Scanning Theory PDFs..."):
                keywords = [k.strip() for k in keyword_t1.split(",") if k.strip()]
                # Allowed components for P1 and P3
                p1_p3_papers = ["11", "12", "13", "31", "32", "33"]
                st.session_state.theory_results = search_pdfs(keywords, FOLDERS["theory"], p1_p3_papers)
        else:
            st.warning("Please enter a search keyword.")

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
    keyword_t2 = st.text_input("Enter Practical Keywords (e.g., 'Mail Merge', 'JavaScript', 'Vector Graphics')", key="t2_kw")
    
    if st.button("Search Practical Papers", type="primary"):
        if keyword_t2:
            with st.spinner("Scanning Practical PDFs..."):
                keywords = [k.strip() for k in keyword_t2.split(",") if k.strip()]
                # Allowed components for P2 and P4
                p2_p4_papers = ["21", "22", "23", "41", "42", "43"]
                st.session_state.practical_results = search_pdfs(keywords, FOLDERS["practical"], p2_p4_papers)
        else:
            st.warning("Please enter a search keyword.")

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

        if st.button("🪄 Export Handout to Word (.docx)", type="primary"):
            doc = Document()
            doc.add_heading(f'PTES {SYLLABUS_CODE} IT Handout', 0)
            
            for item in st.session_state.handout_basket:
                doc.add_heading(f"Source: {item['file']} (Page {item['page'] + 1})", level=2)
                pdf_doc = fitz.open(item['path'])
                page = pdf_doc.load_page(item['page'])
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_data = BytesIO(pix.tobytes("png"))
                doc.add_picture(img_data, width=Inches(6.5))
                doc.add_page_break()
                pdf_doc.close()

            target_filename = f"{SYLLABUS_CODE}_IT_Handout.docx"
            doc.save(target_filename)
            
            with open(target_filename, "rb") as f:
                st.download_button("📥 Click to Download Word Document", f, file_name=target_filename)
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
        z_paper = st.selectbox("Select Paper Component", ["21", "22", "23", "41", "42", "43"])

    # Target file naming standard: e.g., 9626_s22_sf_21.zip
    short_year = z_year[-2:]
    expected_zip_name = f"9626_{session_code}{short_year}_sf_{z_paper}.zip"
    zip_path = os.path.join(FOLDERS["zips"], expected_zip_name)

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
        st.warning(f"Source file `{expected_zip_name}` is not available locally in `{FOLDERS['zips']}`.")

# --- TAB 5: ADMIN PANEL ---
with tab5:
    st.header("Admin Upload to Google Drive")
    pwd = st.text_input("Enter Admin Password", type="password")
    
    # Check secrets or environment variable for password
    admin_password = st.secrets.get("ADMIN_PASSWORD", "ptes123")
    
    if pwd == admin_password:
        st.success("Admin Access Granted")
        
        st.subheader("Upload Paper to Google Drive")
        
        col_a, col_b = st.columns(2)
        with col_a:
            up_year = st.number_input("Year", min_value=2018, max_value=2030, value=2026)
            up_paper_num = st.selectbox("Paper Component", ["11", "12", "13", "21", "22", "23", "31", "32", "33", "41", "42", "43"])
            up_file_type = st.radio("File Type", ["Question Paper (qp)", "Marking Scheme (ms)", "Source File (sf/zip)"])
        
        with col_b:
            target_drive_folder_id = st.text_input("Target Google Drive Folder ID")
            uploaded_file = st.file_uploader("Browse File", type=["pdf", "zip"])

        if st.button("🚀 Upload File to Drive", type="primary"):
            if uploaded_file and target_drive_folder_id:
                file_bytes = uploaded_file.read()
                file_id = upload_to_drive(file_bytes, uploaded_file.name, target_drive_folder_id)
                
                if file_id:
                    st.success(f"Successfully uploaded `{uploaded_file.name}` to Drive! (File ID: {file_id})")
                    
                    # Save locally as well to mirror immediately
                    if "qp" in uploaded_file.name or "ms" in uploaded_file.name:
                        dest_folder = FOLDERS["theory"] if up_paper_num[0] in ["1", "3"] else FOLDERS["practical"]
                    else:
                        dest_folder = FOLDERS["zips"]
                        
                    local_save_path = os.path.join(dest_folder, uploaded_file.name)
                    with open(local_save_path, "wb") as f:
                        f.write(file_bytes)
                    st.info(f"Mirrored file to local directory `{dest_folder}`.")
            else:
                st.error("Please provide both a file and a valid Google Drive Folder ID.")

# --- FOOTER ---
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
