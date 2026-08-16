import os
import io
import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==============================================================================
# 1. PAGE CONFIGURATION & STYLING (Matching 9699 Sociology Palette)
# ==============================================================================
st.set_page_config(
    page_title="9626 IT PYP Portal",
    page_icon="💻",
    layout="wide"
)

CUSTOM_CSS = """
<style>
    /* Main Background Color */
    .stApp {
        background-color: #63D0F8;
    }
    
    /* Headers Styling */
    h1, h2, h3 {
        color: #0E2F56;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Card/Container Backgrounds */
    div[data-testid="stVerticalBlock"] > div {
        background-color: #FFFFFF;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1);
    }
    
    /* Accent Buttons */
    .stButton>button {
        background-color: #CCFF00;
        color: #0E2F56;
        font-weight: bold;
        border-radius: 8px;
        border: none;
        padding: 8px 16px;
    }
    .stButton>button:hover {
        background-color: #B3E600;
        color: #000000;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ==============================================================================
# 2. LOCAL FOLDER & DRIVE CONFIGURATION
# ==============================================================================
PARENT_DIR = "9626ITMaterials"
FOLDERS = {
    "p1_it": os.path.join(PARENT_DIR, "p1_it"),
    "p2_it": os.path.join(PARENT_DIR, "p2_it"),
    "p3_it": os.path.join(PARENT_DIR, "p3_it"),
    "p4_it": os.path.join(PARENT_DIR, "p4_it"),
    "ms_p1_p2": os.path.join(PARENT_DIR, "ms_p1_p2"),
    "ms_p3_p4": os.path.join(PARENT_DIR, "ms_p3_p4")
}

# Ensure local directories exist
for folder_path in FOLDERS.values():
    os.makedirs(folder_path, exist_ok=True)

# Initialize Session State for PYP Cart
if "cart" not in st.session_state:
    st.session_state.cart = []

# ==============================================================================
# 3. HELPER FUNCTIONS
# ==============================================================================
def get_drive_service():
    """Authenticates and returns the Google Drive API service using Streamlit secrets."""
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        return build('drive', 'v3', credentials=creds)
    return None

def sync_folder_from_drive(drive_folder_id, local_folder_path):
    """Syncs files from a Google Drive folder to a local directory."""
    service = get_drive_service()
    if not service or not drive_folder_id:
        return 0
    
    query = f"'{drive_folder_id}' in parents and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    
    downloaded_count = 0
    for file in files:
        file_path = os.path.join(local_folder_path, file['name'])
        if not os.path.exists(file_path):
            request = service.files().get_media(fileId=file['id'])
            with open(file_path, 'wb') as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
            downloaded_count += 1
    return downloaded_count

def search_pdf_keywords(folder_path, keyword):
    """Searches for a keyword across all PDFs in a specified folder."""
    results = []
    if not os.path.exists(folder_path) or not keyword:
        return results
    
    for filename in sorted(os.listdir(folder_path)):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            try:
                doc = fitz.open(pdf_path)
                for page_num in range(len(doc)):
                    text = doc[page_num].get_text("text")
                    if keyword.lower() in text.lower():
                        results.append({
                            "file_name": filename,
                            "page_num": page_num + 1,
                            "pdf_path": pdf_path
                        })
                doc.close()
            except Exception as e:
                st.error(f"Error reading {filename}: {e}")
    return results

def render_pdf_page_as_image(pdf_path, page_num, zoom=2.0):
    """Renders a specific PDF page as an image bytes object."""
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_num - 1)
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img_data = pix.tobytes("png")
    doc.close()
    return img_data

def generate_word_worksheet(cart_items):
    """Generates a Word Document (.docx) containing selected PYP pages."""
    doc = Document()
    
    # Header Title
    title = doc.add_heading('9626 Information Technology Worksheet', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for idx, item in enumerate(cart_items, start=1):
        doc.add_heading(f"Question {idx}: {item['file_name']} (Page {item['page_num']})", level=2)
        
        # Render image of the page
        img_bytes = render_pdf_page_as_image(item['pdf_path'], item['page_num'], zoom=1.5)
        image_stream = io.BytesIO(img_bytes)
        doc.add_picture(image_stream, width=Inches(6.0))
        doc.add_page_break()
        
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==============================================================================
# 4. MAIN USER INTERFACE
# ==============================================================================
st.title("💻 9626 Information Technology PYP Portal")
st.write("Welcome to the **A-Level 9626 IT** Past Year Paper portal. Search topics, view mark schemes, and build custom worksheets.")

tabs = st.tabs([
    "📘 AS Level IT (P1 & P2)", 
    "📙 A Level IT (P3 & P4)", 
    "🛒 PYP Cart & Worksheet Builder", 
    "🔍 Mark Scheme Search", 
    "⚙️ Admin & Sync"
])

# ------------------------------------------------------------------------------
# TAB 1: AS LEVEL IT
# ------------------------------------------------------------------------------
with tabs[0]:
    st.header("AS Level Information Technology")
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Paper 1: Theory")
        p1_keyword = st.text_input("Search P1 Topics (e.g., Data, Database, Network):", key="p1_search")
        if st.button("Search Paper 1", key="btn_p1"):
            results = search_pdf_keywords(FOLDERS["p1_it"], p1_keyword)
            st.write(f"Found **{len(results)}** match(es).")
            for res in results:
                st.write(f"📄 **{res['file_name']}** - Page {res['page_num']}")
                img = render_pdf_page_as_image(res['pdf_path'], res['page_num'])
                st.image(img, use_container_width=True)
                if st.button(f"Add {res['file_name']} (P.{res['page_num']}) to Cart", key=f"add_p1_{res['file_name']}_{res['page_num']}"):
                    st.session_state.cart.append(res)
                    st.success("Added to cart!")

    with col2:
        st.subheader("Paper 2: Practical")
        p2_keyword = st.text_input("Search P2 Topics (e.g., Spreadsheets, Database, Sound):", key="p2_search")
        if st.button("Search Paper 2", key="btn_p2"):
            results = search_pdf_keywords(FOLDERS["p2_it"], p2_keyword)
            st.write(f"Found **{len(results)}** match(es).")
            for res in results:
                st.write(f"📄 **{res['file_name']}** - Page {res['page_num']}")
                img = render_pdf_page_as_image(res['pdf_path'], res['page_num'])
                st.image(img, use_container_width=True)
                if st.button(f"Add {res['file_name']} (P.{res['page_num']}) to Cart", key=f"add_p2_{res['file_name']}_{res['page_num']}"):
                    st.session_state.cart.append(res)
                    st.success("Added to cart!")

# ------------------------------------------------------------------------------
# TAB 2: A LEVEL IT
# ------------------------------------------------------------------------------
with tabs[1]:
    st.header("A Level Information Technology")
    col3, col4 = st.columns(2)
    
    with col3:
        st.subheader("Paper 3: Advanced Theory")
        p3_keyword = st.text_input("Search P3 Topics (e.g., Project Management, JavaScript, Encryption):", key="p3_search")
        if st.button("Search Paper 3", key="btn_p3"):
            results = search_pdf_keywords(FOLDERS["p3_it"], p3_keyword)
            st.write(f"Found **{len(results)}** match(es).")
            for res in results:
                st.write(f"📄 **{res['file_name']}** - Page {res['page_num']}")
                img = render_pdf_page_as_image(res['pdf_path'], res['page_num'])
                st.image(img, use_container_width=True)
                if st.button(f"Add {res['file_name']} (P.{res['page_num']}) to Cart", key=f"add_p3_{res['file_name']}_{res['page_num']}"):
                    st.session_state.cart.append(res)
                    st.success("Added to cart!")

    with col4:
        st.subheader("Paper 4: Advanced Practical")
        p4_keyword = st.text_input("Search P4 Topics (e.g., Vector Graphics, Animation, Web Authoring):", key="p4_search")
        if st.button("Search Paper 4", key="btn_p4"):
            results = search_pdf_keywords(FOLDERS["p4_it"], p4_keyword)
            st.write(f"Found **{len(results)}** match(es).")
            for res in results:
                st.write(f"📄 **{res['file_name']}** - Page {res['page_num']}")
                img = render_pdf_page_as_image(res['pdf_path'], res['page_num'])
                st.image(img, use_container_width=True)
                if st.button(f"Add {res['file_name']} (P.{res['page_num']}) to Cart", key=f"add_p4_{res['file_name']}_{res['page_num']}"):
                    st.session_state.cart.append(res)
                    st.success("Added to cart!")

# ------------------------------------------------------------------------------
# TAB 3: PYP CART & WORKSHEET BUILDER
# ------------------------------------------------------------------------------
with tabs[2]:
    st.header("🛒 Selected Questions Cart")
    
    if not st.session_state.cart:
        st.info("Your cart is currently empty. Search papers and add questions to build a worksheet.")
    else:
        st.write(f"Total questions selected: **{len(st.session_state.cart)}**")
        
        if st.button("Clear Cart"):
            st.session_state.cart = []
            st.rerun()
            
        for i, item in enumerate(st.session_state.cart):
            st.write(f"**Item {i+1}:** {item['file_name']} (Page {item['page_num']})")
        
        # Download Word Worksheet
        doc_file = generate_word_worksheet(st.session_state.cart)
        st.download_button(
            label="📄 Export Worksheet as Word Document (.docx)",
            data=doc_file,
            file_name="9626_IT_Custom_Worksheet.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# ------------------------------------------------------------------------------
# TAB 4: MARK SCHEME SEARCH
# ------------------------------------------------------------------------------
with tabs[3]:
    st.header("🔍 Mark Scheme Search")
    ms_choice = st.radio("Select Level:", ["AS Level (P1 & P2 MS)", "A Level (P3 & P4 MS)"])
    ms_keyword = st.text_input("Enter Mark Scheme Keyword/Term:", key="ms_search")
    
    if st.button("Search Mark Schemes", key="btn_ms"):
        target_folder = FOLDERS["ms_p1_p2"] if "AS Level" in ms_choice else FOLDERS["ms_p3_p4"]
        results = search_pdf_keywords(target_folder, ms_keyword)
        st.write(f"Found **{len(results)}** match(es).")
        for res in results:
            st.write(f"📝 **{res['file_name']}** - Page {res['page_num']}")
            img = render_pdf_page_as_image(res['pdf_path'], res['page_num'])
            st.image(img, use_container_width=True)

# ------------------------------------------------------------------------------
# TAB 5: ADMIN & DRIVE SYNC
# ------------------------------------------------------------------------------
with tabs[4]:
    st.header("⚙️ Admin Dashboard & Google Drive Sync")
    st.write("Sync your local folder structure with Google Drive cloud storage.")
    
    if st.button("Sync All Folders from Google Drive"):
        if "gdrive_folders" in st.secrets:
            total_synced = 0
            for key, folder_path in FOLDERS.items():
                drive_id = st.secrets["gdrive_folders"].get(key)
                if drive_id:
                    count = sync_folder_from_drive(drive_id, folder_path)
                    st.write(f"Synced `{key}`: **{count}** new file(s) downloaded.")
                    total_synced += count
            st.success(f"Sync complete! Total new files downloaded: {total_synced}")
        else:
            st.error("Google Drive folder configuration missing in `.streamlit/secrets.toml`!")
