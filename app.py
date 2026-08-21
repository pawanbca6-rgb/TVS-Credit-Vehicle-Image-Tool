import streamlit as st
import time
import os
import re
import requests
import pandas as pd
import gc
import zipfile
import io
from datetime import datetime, timedelta, timezone
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from PIL import Image as PILImage
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# --- STRICT SORTING ORDER CONFIGURATION (For Gallery Mode) ---
DESIRED_ORDER = [
    "Front Side", "Right Side", "Back Side", "Left Side", "Engine",
    "Engine & Plate Photo", "Chassis Plate", "Odometer", "Dashboard",
    "Seat", "Gear and Pedal", "Wind Shield", "All Boot", "Tyre 1",
    "Tyre 2", "Tyre 3", "Tyre 4", "Chassis Number", "Chassis Print",
    "Selfie with Vehicle", "RC Front"
]

DISALLOWED_KEYWORDS = ['score', 'qr', 'logo', 'icon', 'pdf', 'verified', 'badge', 'banner']

def sanitize_name(name):
    """Normalize label string for exact fuzzy matching."""
    cleaned = re.sub(r'[^a-zA-Z0-9]', '', str(name)).lower()
    if cleaned in ['repofront', 'front', 'frontimage', 'frontview']:
        cleaned = 'frontside'
    return cleaned

ORDER_MAP = {sanitize_name(label): idx for idx, label in enumerate(DESIRED_ORDER, start=1)}

# --- STATE INITIALIZATION ---
if "processing_complete" not in st.session_state:
    st.session_state.processing_complete = False
if "zip_name" not in st.session_state:
    st.session_state.zip_name = ""
if "total_extracted" not in st.session_state:
    st.session_state.total_extracted = 0
if "success_count" not in st.session_state:
    st.session_state.success_count = 0
if "failure_count" not in st.session_state:
    st.session_state.failure_count = 0
if "final_df_summary" not in st.session_state:
    st.session_state.final_df_summary = None

# --- UI CONFIGURATION ---
st.set_page_config(
    page_title="TVS Auto-Detect Image Tool", 
    page_icon="📸",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] { gap: 0.5rem !important; }
    section[data-testid="stSidebar"] { overflow-y: hidden !important; overflow-x: hidden !important; }
    .main-header { font-size: 34px; font-weight: 800; color: #1E3A8A; margin-bottom: 2px; }
    .metric-box { background-color: #F3F4F6; padding: 18px; border-radius: 10px; border-left: 6px solid #1E40AF; box-shadow: 0 2px 4px rgba(0,0,0,0.04); }
    .footer-text { text-align: center; font-size: 14px; color: #9CA3AF; margin-top: 50px; padding-top: 20px; border-top: 1px solid #E5E7EB; }
    .note-box { background-color: #EFF6FF; color: #1E40AF; padding: 12px; border-radius: 6px; border-left: 4px solid #2563EB; margin-bottom: 8px; font-size: 14px; font-weight: 500; }
    .premium-card { background: #FFFFFF; padding: 24px; border-radius: 14px; border-bottom: 5px solid #CBD5E1; box-shadow: 0 4px 20px rgba(0,0,0,0.03); text-align: center; transition: transform 0.2s; }
    .premium-card:hover { transform: translateY(-2px); }
    .card-title { font-size: 12px; color: #64748B; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
    .card-value { font-size: 38px; font-weight: 800; margin-top: 8px; }
    
    div[data-testid="stFileUploaderDropzoneInstructions"] small,
    div[data-testid="stFileUploaderDropzoneInstructions"] div,
    div[data-testid="stWidgetFormInstructions"],
    .stWidgetFormInstructions { display: none !important; }
    div[data-testid="stFileUploaderDropzoneInstructions"]::after {
        content: "⚠️ Please upload valid Excel (.xlsx) file";
        font-size: 13px; color: #DC2626; font-weight: 600; display: block; margin-top: 12px; text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🎛️ Control & Template Center")
    st.write("---")
    
    sample_data = {
        "AGREEMENT NO": ["TN3006CA0024047", "TN3006TW0162293"],
        "VALUATION REPORT LINK": ["https://valuation.mytvs.in/report?id=13a2a699...", "https://valuation.mytvs.in/..."],
        "REPO FRONT": ["https://icms.tvscredit.com/Vehical_Images.aspx?name=...", "https://icms.tvscredit.com/..."]
    }
    sample_df = pd.DataFrame(sample_data)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        sample_df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    st.download_button(
        label="📥 Download TVS Excel Template",
        data=buffer.getvalue(),
        file_name="tvs_input_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.write("---")
    
    st.markdown("##### ⚙️ Size Filter Configuration")
    min_size_kb = st.number_input("Minimum Image Size Filter (KB):", min_value=0, value=10, step=10)
    st.write("---")
    st.caption("Recommended: Maximum 100 loans per batch for best execution speed.")

# --- MAIN UI HEADER ---
st.markdown("## 📸 TVS Credit Unified Auto-Detect Tool")

st.markdown('<div class="note-box">📌 <strong>This Tool Work Only For TVS Credit links.</strong></div>', unsafe_allow_html=True)
st.markdown("🔗 **Sample Link Format 1 (Standard):** `https://icms.tvscredit.com/Vehical_Images.aspx?name=...`")
st.markdown("🔗 **Sample Link Format 2 (Gallery):** `https://valuation.mytvs.in/report?id=...`")
st.markdown("<br>", unsafe_allow_html=True)

layout_left, layout_right = st.columns([2, 1])

with layout_left:
    st.markdown("### 📂 File Upload")
    uploaded_file = st.file_uploader("Drop xlsx file here", type=["xlsx"], label_visibility="collapsed")

# --- DATASET VALIDATION ---
valid_rows = []
total_loans_count = 0
link_columns = []

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    agreement_col = next((c for c in df.columns if c.upper() in ["AGREEMENT NO", "AGREEMENTNO", "LOAN NO"]), None)
    
    if not agreement_col:
        st.error("❌ Column Validation Error: Sheet inside uploaded file must contain 'AGREEMENT NO' header.")
    else:
        valid_rows = df[df[agreement_col].notna()]
        total_loans_count = len(valid_rows)
        link_columns = [col for col in df.columns if col != agreement_col and col.upper() != "AUCTION ID"]
        
        with layout_left:
            st.markdown("### 🔍 Input Preview")
            preview_df = df.copy()
            preview_df.index = preview_df.index + 1
            st.dataframe(preview_df, use_container_width=True, height=180)
        
        with layout_right:
            st.markdown("### 📊 Queue Status")
            st.markdown(f"""
                <div class="metric-box">
                    <span style='font-size:13px; color:#6B7280; text-transform: uppercase; font-weight:bold;'>Total Loans Queue</span><br>
                    <span style='font-size:32px; font-weight:bold; color:#1E40AF;'>{total_loans_count}</span>
                </div>
            """, unsafe_allow_html=True)

with st.sidebar:
    st.write(" ")
    run_engine = st.button(
        "🚀 Start Downloading", 
        type="primary", 
        use_container_width=True,
        disabled=(uploaded_file is None)
    )

# --- CORE NETWORK & PARSING FUNCTIONS ---
def create_fast_session():
    session = requests.Session()
    retries = Retry(total=3, backoff_factor=0.5, status_forcelist=[500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retries, pool_connections=20, pool_maxsize=20)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    
    # ADVANCED BROWSER HEADERS TO BYPASS 403 CLOUD BLOCKS
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "none",
        "Sec-Fetch-User": "?1",
        "Cache-Control": "max-age=0"
    })
    return session

def download_and_verify(img_url, save_path, session, min_kb):
    try:
        r = session.get(img_url, stream=True, timeout=12)
        if r.status_code == 200:
            content_length = r.headers.get('content-length')
            if content_length and (int(content_length) / 1024) < min_kb:
                return False, f"Skipped (<{min_kb}KB)"
            
            img_data = r.content
            try:
                check_img = PILImage.open(io.BytesIO(img_data))
                check_img.verify()
            except (IOError, SyntaxError):
                return False, "Corrupted Data"
            
            with open(save_path, 'wb') as f:
                f.write(img_data)
            return True, "Downloaded"
        return False, f"HTTP Error {r.status_code}"
    except Exception as e:
        return False, str(e)

# --- SCRAPER LOGIC 1: ICMS TVS CREDIT ---
def get_real_image_url(webpage_url, session):
    try:
        response = session.get(webpage_url, timeout=12)
        if response.status_code != 200:
            return None, f"HTTP Status {response.status_code}"
        
        soup = BeautifulSoup(response.text, 'html.parser')
        img_tag = soup.find('img', id='imgDisplay')
        
        if img_tag and img_tag.get('src'):
            src = img_tag.get('src')
            full_img_url = urljoin(webpage_url, src)
            return full_img_url, "Success"
        return None, "Secure tag 'imgDisplay' not found"
    except Exception as e:
        return None, str(e)

# --- SCRAPER LOGIC 2: MYTVS GALLERY ---
def extract_gallery_images_only(url, col_name, session):
    results = []
    try:
        response = session.get(url, timeout=12)
        if response.status_code != 200:
            return results, f"HTTP Status {response.status_code}"
        
        content_type = response.headers.get('Content-Type', '')
        if 'image' in content_type:
            if not any(k in url.lower() for k in DISALLOWED_KEYWORDS):
                results.append((col_name, url))
            return results, "Success (Direct Image)"

        soup = BeautifulSoup(response.text, 'html.parser')
        gallery_section = None
        for header in soup.find_all(['h2', 'h3', 'h4', 'div']):
            if 'vehicle gallery' in header.get_text(strip=True).lower():
                gallery_section = header.find_parent(['section', 'div', 'main']) or header.parent
                break

        imgs_to_parse = gallery_section.find_all('img') if gallery_section else soup.find_all('img')
        
        if not imgs_to_parse:
            return results, "No images found on page"

        for img in imgs_to_parse:
            src = img.get('src') or img.get('data-src')
            if not src:
                continue
            
            full_src = urljoin(url, src)
            alt = img.get('alt', '').strip()
            
            check_str = f"{alt} {full_src}".lower()
            if any(k in check_str for k in DISALLOWED_KEYWORDS):
                continue

            label = alt
            if not label:
                parent = img.find_parent()
                if parent:
                    label = parent.get_text(strip=True)
            if not label:
                label = col_name

            results.append((label, full_src))

        if not results:
            return results, "Images found but rejected by keyword filter"
            
        return results, "Success"
    except Exception as e:
        return results, f"Scrape Error: {str(e)}"

# --- AUTO-DETECT PIPELINE LOGIC ---
def process_loan_auto_detect(agreement_no, row_data, columns, min_kb, base_dir):
    agreement_no = str(agreement_no).strip()
    indian_offset = timezone(timedelta(hours=5, minutes=30))
    timestamp = datetime.now(indian_offset).strftime("%d-%m-%Y %H:%M:%S")
    
    loan_folder = os.path.join(base_dir, agreement_no)
    os.makedirs(loan_folder, exist_ok=True)
    
    session = create_fast_session()
    
    success_downloads = 0
    failed_downloads = 0
    details = []
    
    standard_tasks = []
    gallery_raw = []
    
    for col in columns:
        url = str(row_data[col]).strip()
        if pd.isna(url) or not url.startswith('http'):
            continue
            
        if 'icms.tvscredit.com' in url.lower():
            img_url, err_msg = get_real_image_url(url, session)
            if img_url:
                clean_name = re.sub(r'[^a-zA-Z0-9_]', '', col.replace(" ", "_"))
                standard_tasks.append((img_url, f"{clean_name}.jpg", col))
            else:
                failed_downloads += 1
                details.append(f"{col}: Scrape error ({err_msg})")
                
        elif 'mytvs.in' in url.lower():
            extracted, err_msg = extract_gallery_images_only(url, col, session)
            if extracted:
                gallery_raw.extend(extracted)
            else:
                failed_downloads += 1
                details.append(f"{col}: Gallery Error ({err_msg})")
        else:
            details.append(f"{col}: Skipped unrecognized domain")

    for img_url, filename, col in standard_tasks:
        save_path = os.path.join(loan_folder, filename)
        ok, status_str = download_and_verify(img_url, save_path, session, min_kb)
        if ok:
            success_downloads += 1
        else:
            failed_downloads += 1
        details.append(f"{col}: {status_str}")

    if gallery_raw:
        slot_allocations = {}
        for label, img_url in gallery_raw:
            clean_lbl = sanitize_name(label)
            seq_num = ORDER_MAP.get(clean_lbl)
            if seq_num and seq_num not in slot_allocations:
                slot_allocations[seq_num] = (label, img_url)

        for seq_num in sorted(slot_allocations.keys()):
            label, img_url = slot_allocations[seq_num]
            clean_title = DESIRED_ORDER[seq_num - 1]
            filename = f"{seq_num}_{re.sub(r'[^a-zA-Z0-9_]', '', clean_title.replace(' ', '_'))}.jpg"
            
            save_path = os.path.join(loan_folder, filename)
            ok, status_str = download_and_verify(img_url, save_path, session, min_kb)
            if ok:
                success_downloads += 1
            else:
                failed_downloads += 1
            details.append(f"{filename}: {status_str}")

    session.close()
    
    return {
        "Agreement_No": agreement_no,
        "Timestamp": timestamp,
        "Total_Links_Processed": success_downloads + failed_downloads,
        "Success_Downloads": success_downloads,
        "Failed_Downloads": failed_downloads,
        "Summary_Logs": " | ".join(details)
    }

# --- ENGINE FLOW ---
if run_engine and uploaded_file is not None:
    st.session_state.processing_complete = False
    st.session_state.total_extracted = 0
    st.session_state.success_count = 0
    st.session_state.failure_count = 0
    st.session_state.final_df_summary = None
    
    CURRENT_BATCH_DIR = os.path.abspath("TVS_Downloaded_Images")
    if os.path.exists(CURRENT_BATCH_DIR):
        import shutil
        shutil.rmtree(CURRENT_BATCH_DIR)
    os.makedirs(CURRENT_BATCH_DIR, exist_ok=True)
    
    try:
        report_data = []
        st.write("---")
        st.markdown("### ⚙️ Processing Pipeline Logs")
        
        engine_progressbar = st.progress(0)
        percentage_text = st.empty()  
        completed_tasks = 0
        
        agreement_col = next((c for c in df.columns if c.upper() in ["AGREEMENT NO", "AGREEMENTNO", "LOAN NO"]), None)
        
        with st.status("Auto-Detecting URLs and Downloading Assets...", expanded=True) as log_context:
            workers = min(10, len(valid_rows)) if len(valid_rows) > 0 else 1
            with ThreadPoolExecutor(max_workers=workers) as executor:
                future_to_loan = {
                    executor.submit(
                        process_loan_auto_detect, row[agreement_col], row, link_columns, min_size_kb, CURRENT_BATCH_DIR
                    ): row[agreement_col] for _, row in valid_rows.iterrows()
                }
                
                for future in as_completed(future_to_loan):
                    loan_id = future_to_loan[future]
                    try:
                        res = future.result()
                        completed_tasks += 1
                        
                        st.session_state.total_extracted += res["Success_Downloads"]
                        if res["Success_Downloads"] > 0:
                            st.session_state.success_count += 1
                        else:
                            st.session_state.failure_count += 1
                            
                        report_data.append(res)
                        log_context.write(f"✅ **Done ({completed_tasks}/{total_loans_count}):** `{res['Agreement_No']}` -> Saved {res['Success_Downloads']} assets.")
                    except Exception as e:
                        completed_tasks += 1
                        st.session_state.failure_count += 1
                        log_context.write(f"🔴 **Error on {loan_id}:** {e}")
                        
                    ratio = completed_tasks / total_loans_count
                    percentage_text.markdown(f"**📊 Total Download Progress: {int(ratio * 100)}%**")
                    engine_progressbar.progress(ratio)
                    gc.collect()
            
            log_context.update(label="🚀 Execution Completed Successfully!", state="complete", expanded=False)
            
        if report_data:
            report_df = pd.DataFrame(report_data)
            st.session_state.final_df_summary = report_df
            report_df.to_csv(os.path.join(CURRENT_BATCH_DIR, "Batch_Download_Report.csv"), index=False)
            
            master_zip = "TVS_Master_Package.zip"
            with zipfile.ZipFile(master_zip, 'w', zipfile.ZIP_DEFLATED) as z:
                for root, dirs, files in os.walk(CURRENT_BATCH_DIR):
                    for f in files:
                        full_p = os.path.join(root, f)
                        z.write(full_p, os.path.relpath(full_p, CURRENT_BATCH_DIR))
            
            st.balloons()
            st.session_state.zip_name = master_zip
            st.session_state.processing_complete = True
            
    except Exception as outer_err:
        st.error(f"Critical execution error: {outer_err}")

# --- METRIC DASHBOARD OUTCOMES ---
if st.session_state.processing_complete and os.path.exists(st.session_state.zip_name):
    st.write("---")
    st.markdown("### 🏁 Execution Dashboard Analytics")
    
    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f'<div class="premium-card" style="border-bottom: 5px solid #10B981;"><span class="card-title">Success Loans</span><div class="card-value" style="color:#10B981;">{st.session_state.success_count}</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="premium-card" style="border-bottom: 5px solid #EF4444;"><span class="card-title">Failed Loans</span><div class="card-value" style="color:#EF4444;">{st.session_state.failure_count}</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="premium-card" style="border-bottom: 5px solid #F59E0B;"><span class="card-title">Total Extracted Images</span><div class="card-value" style="color:#F59E0B;">{st.session_state.total_extracted}</div></div>', unsafe_allow_html=True)
        
    tab_audit, tab_visuals = st.tabs(["📋 Process Log Report", "📊 Metrics Visualization"])
    
    with tab_audit:
        if st.session_state.final_df_summary is not None:
            st.dataframe(st.session_state.final_df_summary, use_container_width=True, hide_index=True)
            
    with tab_visuals:
        health_chart_data = pd.DataFrame({
            "Status": ["Success Pipeline", "Failed Pipeline"],
            "Values": [st.session_state.success_count, st.session_state.failure_count]
        })
        st.data_editor(
            health_chart_data,
            column_config={
                "Values": st.column_config.ProgressColumn(
                    "Ratio Scale", format="%d", min_value=0, max_value=total_loans_count if total_loans_count > 0 else 10
                ),
            },
            hide_index=True, use_container_width=True, key="metric_visual_editor"
        )
        
    st.write(" ")
    with open(st.session_state.zip_name, "rb") as fp:
        indian_offset = timezone(timedelta(hours=5, minutes=30))
        custom_name = f"TVS_Unified_Images_{datetime.now(indian_offset).strftime('%d-%m-%Y_%H-%M-%S')}.zip"
        
        st.download_button(
            label="📥 DOWNLOAD ZIP FILE AND COMPLETION REPORT",
            data=fp,
            file_name=custom_name,
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

st.markdown('<div class="footer-text">🛠️ TVS Credit Automated Pipeline Tool</div>', unsafe_allow_html=True)
