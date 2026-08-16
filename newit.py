import os
import io
import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==============================================================================
# 1. PAGE CONFIGURATION & STYLING (Matching 9699 Sociology Theme Exactly)
# ==============================================================================
st.set_page_config(
    page_title="9626 IT PYP Portal",
    page_icon="💻",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Sociology Color Scheme (#63D0F8 background, #CCFF00 accents)
CUSTOM_CSS = """
<style>
    /* Main Background Color */
    .stApp {
        background-color: #63D0F8;
    }
    
    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #0E2F56;
        color: #FFFFFF;
    }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] label {
        color: #FFFFFF !important;
    }
    
    /* Headers Styling */
    h1, h2, h3 {
        color: #0E2F56;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Content Cards/Containers */
    div[data-testid="stVerticalBlock"] > div {
        background-color: #FFFFFF;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1);
    }
    
    /* Buttons */
    .stButton>button {
        background-color: #CCFF00;
        color: #0E2F56;
        font-weight: bold;
        border-radius: 8px;
        border: none;
        padding: 8px 16px;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #B3E600;
        color: #000000;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ==============================================================================
# 2. LOCAL DIRECTORY & DRIVE CONFIGURATION
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

# Initialize Session State
if "cart" not in st.session_state:
    st.session_state.cart = []
if "search_query" not in st.session_state:
    st.session_state.search_query = ""

# ==============================================================================
# 3. HELPER FUNCTIONS
# ==============================================================================
def get_drive_service():
    """Authenticates with Google Drive API using Streamlit secrets."""
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        return build('drive', 'v3', credentials=creds)
    return None

def sync_folder_from_drive(drive_folder_id, local_folder_path):
    """Downloads missing files from Google Drive folder to local folder."""
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
    """Searches PDF files in a folder for matching text."""
    results = []
    if not os.path.exists(folder_path) or not keyword:
        return results
    
    for filename in sorted(os.listdir(folder_path)):
        if filename.lower().endswith(".pdf"):
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
                st.error(f"Error loading {filename}: {e}")
    return results

def render_pdf_page_as_image(pdf_path, page_num, zoom=2.0):
    """Converts a PDF page into a PNG image byte array."""
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_num - 1)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    img_data = pix.tobytes("png")
    doc.close()
    return img_data

def generate_word_worksheet(cart_items):
    """Compiles selected PYP pages into an editable Word document."""
    doc = Document()
    title = doc.add_heading('9626 Information Technology Custom Worksheet', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for idx, item in enumerate(cart_items, start=1):
        doc.add_heading(f"Question {idx}: {item['file_name']} (Page {item['page_num']})", level=2)
        img_bytes = render_pdf_page_as_image(item['pdf_path'], item['page_num'], zoom=1.5)
        doc.add_picture(io.BytesIO(img_bytes), width=Inches(6.0))
        doc.add_page_break()
        
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==============================================================================
# 4. SIDEBAR PANEL (Preserving Sociology Portal Structure)
# ==============================================================================
with st.sidebar:
    st.title("💻 IT 9626 Control Panel")
    st.markdown("---")
    
    # Navigation / View Selection
    view_option = st.radio(
        "Select Portal Mode:",
        ["📘 AS Level (P1 & P2)", "📙 A Level (P3 & P4)", "🔍 Mark Schemes", "🛒 PYP Cart & Export"]
    )
    
    st.markdown("---")
    st.subheader("⚡ Quick IT Preset Search")
    preset_topic = st.selectbox(
        "Select Syllabus Topic:",
        ["", "Data Processing", "Database", "Spreadsheets", "Networks", "Sound Editing", 
         "Video Editing", "Project Management", "Encryption", "JavaScript", "Vector Graphics"]
    )
    if preset_topic:
        st.session_state.search_query = preset_topic
        
    st.markdown("---")
    st.subheader("☁️ Google Drive Sync")
    if st.button("Sync Local Folders"):
        if "gdrive_folders" in st.secrets:
            with st.spinner("Syncing files from Google Drive..."):
                total_files = 0
                for key, path in FOLDERS.items():
                    drive_id = st.secrets["gdrive_folders"].get(key)
                    if drive_id:
                        total_files += sync_folder_from_drive(drive_id, path)
                st.success(f"Synced {total_files} new files!")
        else:
            st.error("Missing Google Drive secrets!")

    st.markdown("---")
    st.write(f"🛒 **Items in Cart:** {len(st.session_state.cart)}")

# ==============================================================================
# 5. MAIN CONTENT AREA
# ==============================================================================
st.title("💻 9626 Information Technology PYP Portal")

# VIEW 1: AS LEVEL (P1 & P2)
if view_option == "📘 AS Level (P1 & P2)":
    st.header("AS Level Information Technology")
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Paper 1: Theory")
        search_p1 = st.text_input("Search P1 Keyword:", value=st.session_state.search_query, key="p1_input")
        if st.button("Search P1", key="btn_p1"):
            results = search_pdf_keywords(FOLDERS["p1_it"], search_p1)
            st.write(f"Found **{len(results)}** match(es).")
            for res in results:
                st.write(f"📄 **{res['file_name']}** - Page {res['page_num']}")
                st.image(render_pdf_page_as_image(res['pdf_path'], res['page_num']), use_container_width=True)
                if st.button("Add to Cart", key=f"add_p1_{res['file_name']}_{res['page_num']}"):
                    st.session_state.cart.append(res)
                    st.success("Added to cart!")

    with col2:
        st.subheader("Paper 2: Practical")
        search_p2 = st.text_input("Search P2 Keyword:", value=st.session_state.search_query, key="p2_input")
        if st.button("Search P2", key="btn_p2"):
            results = search_pdf_keywords(FOLDERS["p2_it"], search_p2)
            st.write(f"Found **{len(results)}** match(es).")
            for res in results:
                st.write(f"📄 **{res['file_name']}** - Page {res['page_num']}")
                st.image(render_pdf_page_as_image(res['pdf_path'], res['page_num']), use_container_width=True)
                if st.button("Add to Cart", key=f"add_p2_{res['file_name']}_{res['page_num']}"):
                    st.session_state.cart.append(res)
                    st.success("Added to cart!")

# VIEW 2: A LEVEL (P3 & P4)
elif view_option == "📙 A Level (P3 & P4)":
    st.header("A Level Information Technology")
    col3, col4 = st.columns(2)
    
    with col3:
        st.subheader("Paper 3: Advanced Theory")
        search_p3 = st.text_input("Search P3 Keyword:", value=st.session_state.search_query, key="p3_input")
        if st.button("Search P3", key="btn_p3"):
            results = search_pdf_keywords(FOLDERS["p3_it"], search_p3)
            st.write(f"Found **{len(results)}** match(es).")
            for res in results:
                st.write(f"📄 **{res['file_name']}** - Page {res['page_num']}")
                st.image(render_pdf_page_as_image(res['pdf_path'], res['page_num']), use_container_width=True)
                if st.button("Add to Cart", key=f"add_p3_{res['file_name']}_{res['page_num']}"):
                    st.session_state.cart.append(res)
                    st.success("Added to cart!")

    with col4:
        st.subheader("Paper 4: Advanced Practical")
        search_p4 = st.text_input("Search P4 Keyword:", value=st.session_state.search_query, key="p4_input")
        if st.button("Search P4", key="btn_p4"):
            results = search_pdf_keywords(FOLDERS["p4_it"], search_p4)
            st.write(f"Found **{len(results)}** match(es).")
            for res in results:
                st.write(f"📄 **{res['file_name']}** - Page {res['page_num']}")
                st.image(render_pdf_page_as_image(res['pdf_path'], res['page_num']), use_container_width=True)
                if st.button("Add to Cart", key=f"add_p4_{res['file_name']}_{res['page_num']}"):
                    st.session_state.cart.append(res)
                    st.success("Added to cart!")

# VIEW 3: MARK SCHEMES
elif view_option == "🔍 Mark Schemes":
    st.header("🔍 Mark Scheme Search")
    ms_level = st.radio("Select Target Level:", ["AS Level (MS P1 & P2)", "A Level (MS P3 & P4)"])
    ms_query = st.text_input("Search Mark Schemes:", value=st.session_state.search_query, key="ms_input")
    
    if st.button("Search Mark Schemes", key="btn_ms"):
        target_folder = FOLDERS["ms_p1_p2"] if "AS Level" in ms_level else FOLDERS["ms_p3_p4"]
        results = search_pdf_keywords(target_folder, ms_query)
        st.write(f"Found **{len(results)}** match(es).")
        for res in results:
            st.write(f"📝 **{res['file_name']}** - Page {res['page_num']}")
            st.image(render_pdf_page_as_image(res['pdf_path'], res['page_num']), use_container_width=True)

# VIEW 4: CART & EXPORT
elif view_option == "🛒 PYP Cart & Export":
    st.header("🛒 PYP Cart & Worksheet Generator")
    if not st.session_state.cart:
        st.info("Your cart is empty. Use the sidebar to navigate to AS or A Level papers and search for questions.")
    else:
        st.write(f"Total Selected Questions: **{len(st.session_state.cart)}**")
        if st.button("Clear Cart", key="clear_cart"):
            st.session_state.cart = []
            st.rerun()
            
        for i, item in enumerate(st.session_state.cart):
            st.write(f"**Item {i+1}:** {item['file_name']} (Page {item['page_num']})")
            
        doc_file = generate_word_worksheet(st.session_state.cart)
        st.download_button(
            label="📄 Download Custom Worksheet (.docx)",
            data=doc_file,
            file_name="9626_IT_Worksheet.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
