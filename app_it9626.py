import os
import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO

# Google Drive API Libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- 1. CONFIGURATION & GOOGLE DRIVE FOLDER MAPPING ---
SYLLABUS_CODE = "9626"

# Google Drive Folder IDs mapped to your specific folder links
FOLDER_IDS = {
    "theory": "1T1sIqRKxF5aO_r0sCyIVxidt0TyXOCcB",      # Theory Papers (P1 & P3)
    "practical": "1EWBiwjvTc12LVtyNi2V9P9RSr8d2vgq7",   # Practical Papers (P2 & P4)
    "zips": "1AsXq8TktyqajB7XTa9SQ5f85Pr6CQcFJ"          # Source Files (.zip)
}

# Local storage directories
LOCAL_FOLDERS = {
    "theory": "9626_theory",
    "practical": "9626_practical",
    "zips": "9626_zips"
}

# Ensure local directories exist
for folder_path in LOCAL_FOLDERS.values():
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

SCOPES = ['https://www.googleapis.com/auth/drive']

# --- 2. GOOGLE DRIVE AUTHENTICATION & UPLOAD ---
def get_drive_service():
    """Authenticates with Google Drive API via Secrets or local service account file."""
    if "gcp_service_account" in st.secrets:
        info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return build('drive', 'v3', credentials=creds)
    elif os.path.exists("service_account.json"):
        creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
        return build('drive', 'v3', credentials=creds)
    else:
        st.error("Google Drive Service Account credentials not found!")
        return None

def upload_file_to_drive(file_bytes, filename, folder_id):
    """Uploads a file directly to the specified Google Drive folder."""
    service = get_drive_service()
    if not service:
        return None

    os.makedirs("temp_upload", exist_ok=True)
    temp_path = os.path.join("temp_upload", filename)
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

# --- 3. HELPER & SEARCH FUNCTIONS ---
def search_pdfs(keyword_list, folder_path, allowed_papers):
    """
    Scans local PDF files for matching keywords for specific paper components.
    e.g., allowed_papers = ["11", "12", "13", "31", "32", "33"] or ["02", "04"]
    """
    results = []
    if not os.path.exists(folder_path):
        return results

    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            # Ensure filename ends with one of the allowed variant codes (e.g., '_02.pdf')
            base_name = os.path.splitext(file)[0]
            is_valid_variant = any(base_name.endswith(f"_{p}") for p in allowed_papers)
            
            if not is_valid_variant:
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
            except Exception:
                continue
    return results

# --- 4. APP STATE INITIALIZATION ---
if 'handout_basket' not in st.session_state:
    st.session_state.handout_basket = []
if 'theory_results' not in st.session_state:
    st.session_state.theory_results = []
if 'practical_results' not in st.session_state:
    st.session_state.practical_results = []

# --- 5. USER INTERFACE ---
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

# Application Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔍 Theory Search (P1 & P3)", 
    "⚙️ Practical Search (P2 & P4)", 
    "🛒 Handout Basket", 
    "📦 Source Files (ZIP)", 
    "🔒 Admin Panel"
])

# TAB 1: THEORY SEARCH (P1 & P3)
with tab1:
    st.header("Search Theory Papers (Paper 1 & Paper 3)")
    st.caption("Variants: Paper 1 (11, 12, 13) | Paper 3 (31, 32, 33)")
    keyword_t1 = st.text_input("Enter Theory Keywords (e.g., 'Normalized', 'Relational Database', 'CSS')", key="t1_kw")

    if st.button("Search Theory Papers", type="primary"):
        if keyword_t1:
            with st.spinner("Scanning Theory PDFs..."):
                keywords = [k.strip() for k in keyword_t1.split(",") if k.strip()]
                p1_p3_variants = ["11", "12", "13", "31", "32", "33"]
                st.session_state.theory_results = search_pdfs(keywords, LOCAL_FOLDERS["theory"], p1_p3_variants)
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

# TAB 2: PRACTICAL SEARCH (P2 & P4)
with tab2:
    st.header("Search Practical Papers (Paper 2 & Paper 4)")
    st.caption("Variants: Paper 2 (02) | Paper 4 (04)")
    keyword_t2 = st.text_input("Enter Practical Keywords (e.g., 'Mail Merge', 'JavaScript', 'Vector Graphics')", key="t2_kw")

    if st.button("Search Practical Papers", type="primary"):
        if keyword_t2:
            with st.spinner("Scanning Practical PDFs..."):
                keywords = [k.strip() for k in keyword_t2.split(",") if k.strip()]
                p2_p4_variants = ["02", "04"]
                st.session_state.practical_results = search_pdfs(keywords, LOCAL_FOLDERS["practical"], p2_p4_variants)
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

# TAB 3: HANDOUT BASKET
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

# TAB 4: SOURCE FILES (ZIP)
with tab4:
    st.header("Download Practical Source Files (ZIP)")
    c1, c2, c3 = st.columns(3)
    with c1:
        z_year = st.selectbox("Select Year", [str(y) for y in range(2026, 2018, -1)])
    with c2:
        z_session = st.selectbox("Select Session", ["March (m)", "June (s)", "Nov (w)"])
        session_code = z_session.split("(")[1].replace(")", "")
    with c3:
        # Paper 2 is variant '02', Paper 4 is variant '04'
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
        st.warning(f"Source file `{expected_zip_name}` is not available locally in `{LOCAL_FOLDERS['zips']}`. Upload it via Admin Panel first.")

# TAB 5: ADMIN PANEL
with tab5:
    st.header("Admin Upload Panel")
    pwd = st.text_input("Enter Admin Password", type="password")

    admin_password = st.secrets.get("ADMIN_PASSWORD", "ptes123")

    if pwd == admin_password:
        st.success("Admin Access Granted")

        col_a, col_b = st.columns(2)
        with col_a:
            up_year = st.number_input("Year", min_value=2018, max_value=2030, value=2026)
            # Aligned component list to match Cambridge variants
            up_paper_num = st.selectbox(
                "Paper Component / Variant", 
                ["11", "12", "13", "02", "31", "32", "33", "04"]
            )
            up_file_type = st.radio("File Type", ["Question Paper (qp)", "Marking Scheme (ms)", "Source File (sf/zip)"])

        with col_b:
            # Automated target folder selection
            if "Source File" in up_file_type:
                target_folder_id = FOLDER_IDS["zips"]
                target_folder_name = "Source Files (ZIP) Folder"
                local_dest = LOCAL_FOLDERS["zips"]
            elif up_paper_num in ["11", "12", "13", "31", "32", "33"]:
                target_folder_id = FOLDER_IDS["theory"]
                target_folder_name = "Theory Papers (P1 & P3) Folder"
                local_dest = LOCAL_FOLDERS["theory"]
            else:
                target_folder_id = FOLDER_IDS["practical"]
                target_folder_name = "Practical Papers (P2 & P4) Folder"
                local_dest = LOCAL_FOLDERS["practical"]

            st.info(f"📁 **Target Folder:** `{target_folder_name}`")
            uploaded_file = st.file_uploader("Browse File", type=["pdf", "zip"])

        if st.button("🚀 Upload File to Drive", type="primary"):
            if uploaded_file:
                file_bytes = uploaded_file.read()
                file_id = upload_file_to_drive(file_bytes, uploaded_file.name, target_folder_id)

                if file_id:
                    st.success(f"Uploaded `{uploaded_file.name}` to Drive! (File ID: `{file_id}`)")

                    # Mirror locally for search indexing
                    local_save_path = os.path.join(local_dest, uploaded_file.name)
                    with open(local_save_path, "wb") as f:
                        f.write(file_bytes)
                    st.info(f"Saved locally to `{local_dest}/{uploaded_file.name}`.")
            else:
                st.error("Please select a file to upload.")

# FOOTER
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
