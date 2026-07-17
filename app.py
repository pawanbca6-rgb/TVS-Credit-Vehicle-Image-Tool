import os
import re
import io
import time
import datetime
import shutil
import zipfile
import pandas as pd
import requests
import fitz  # PyMuPDF
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from pdf2image import convert_from_bytes, convert_from_path

# --- SELENIUM INFRASTRUCTURE CONTROLLERS ---
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

app = Flask(__name__)

# ==========================================
# ⚙️ CONFIGURATION SYSTEM GLOBAL PARAMETERS
# ==========================================
UPLOAD_FOLDER = 'uploads'
OUTPUT_BASE_FOLDER = 'Unified_Inspection_Downloads'
POPPLER_PATH = r"D:\HTML Data\Other\Tool\Mercedes_autoinspect\Mercedes_autoinspect\poppler-25.11.0\Library\bin"
SELENIUM_WAIT_TIME = 12

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_BASE_FOLDER, exist_ok=True)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
}
http_session = requests.Session()
http_session.headers.update(HEADERS)

# ==========================================
# 🛠️ GLOBAL SHARED PIPELINE UTILITIES
# ==========================================

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()

def smart_get(url, retries=3):
    for i in range(1, retries + 1):
        try:
            response = http_session.get(url, timeout=30)
            if response.status_code == 200:
                return response, i
        except Exception:
            pass
        if i < retries:
            time.sleep(1.0)
    return None, retries

def get_next_z_number(folder):
    existing = [f for f in os.listdir(folder) if f.lower().startswith("z") and f.lower().endswith(".jpg")]
    if not existing: return 1
    numbers = [int(n) for f in existing for n in re.findall(r'\d+', f)]
    return max(numbers) + 1 if numbers else 1

def process_pdf_bytes_to_images(content, folder):
    try:
        images = convert_from_bytes(content, poppler_path=POPPLER_PATH, dpi=200)
        z_counter = get_next_z_number(folder)
        for image in images:
            filename = f"Z{str(z_counter).zfill(3)}.jpg"
            image.save(os.path.join(folder, filename), "JPEG")
            z_counter += 1
        return "Success"
    except Exception as e:
        return f"Error: {str(e)}"

def process_pdf_path_to_images(pdf_path, folder):
    try:
        images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH, dpi=200)
        z_counter = get_next_z_number(folder)
        for image in images:
            filename = f"Z{str(z_counter).zfill(3)}.jpg"
            image.save(os.path.join(folder, filename), "JPEG")
            z_counter += 1
        return "Success"
    except Exception as e:
        return f"Error: {str(e)}"

def extract_all_zips(folder):
    extracting = True
    while extracting:
        extracting = False
        for file in os.listdir(folder):
            if file.endswith(".zip"):
                zip_path = os.path.join(folder, file)
                try:
                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                        zip_ref.extractall(folder)
                    os.remove(zip_path)
                    extracting = True
                except Exception:
                    pass

# ==========================================
# 🎛️ CORE ROUTING & LOGICAL CHANNELS
# ==========================================

def execute_autoinspekt_pipeline(link, car_folder, status_entry):
    parts = link.split("/")
    encoded_id = parts[5] if len(parts) > 5 else None
    if not encoded_id:
        status_entry["Pipeline Channel"] = "AutoInspekt (Invalid Token Mapping ID)"
        return
    
    status_entry["Pipeline Channel"] = "AutoInspekt Engine"
    pdf_res, attempts = smart_get(link)
    status_entry["Attempts"] = attempts
    if pdf_res:
        status_entry["PDF Outcome"] = process_pdf_bytes_to_images(pdf_res.content, car_folder)

    zip_link = f"https://aiv2client.autoinspekt.com/download/LeadImages/{encoded_id}"
    zip_res, _ = smart_get(zip_link)
    if zip_res:
        rename_map = {"front": "1_front", "left": "2_left", "right": "3_Right", "rear": "4_Rear"}
        try:
            with zipfile.ZipFile(io.BytesIO(zip_res.content)) as zip_ref:
                for file in zip_ref.namelist():
                    if file.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
                        orig_name = os.path.basename(file)
                        if not orig_name: continue
                        new_name = orig_name.replace(" ", "_")
                        for key, val in rename_map.items():
                            if key.lower() in new_name.lower():
                                new_name = re.compile(re.escape(key), re.IGNORECASE).sub(val, new_name)
                                break
                        with open(os.path.join(car_folder, new_name), "wb") as f:
                            f.write(zip_ref.read(file))
            status_entry["ZIP Outcome"] = "Success"
        except Exception as e:
            status_entry["ZIP Outcome"] = f"Error: {str(e)}"

def execute_tvs_credit_pipeline(row, link_columns, car_folder, status_entry):
    status_entry["Pipeline Channel"] = "TVS Credit / Adroit DOM Parser"
    log_details = []
    success_downloads = 0
    
    for col in link_columns:
        webpage_url = str(row.get(col, "")).strip()
        if not webpage_url or webpage_url == "nan" or not webpage_url.startswith("http"):
            continue
        try:
            res = http_session.get(webpage_url, timeout=15)
            if res.status_code == 200:
                soup = BeautifulSoup(res.text, 'html.parser')
                img_tag = soup.find('img', id='imgDisplay')
                if img_tag and img_tag.get('src'):
                    full_img_url = urljoin(webpage_url, img_tag.get('src'))
                    img_res = http_session.get(full_img_url, timeout=15)
                    if img_res.status_code == 200:
                        clean_col = re.sub(r'[^a-zA-Z0-9_]', '', str(col).replace(" ", "_"))
                        with open(os.path.join(car_folder, f"{clean_col}.jpg"), 'wb') as f:
                            f.write(img_res.content)
                        success_downloads += 1
                        log_details.append(f"[{col}]: Success")
                    else: log_details.append(f"[{col}]: Binary Fetch Error")
                else: log_details.append(f"[{col}]: imgDisplay tag missing")
            else: log_details.append(f"[{col}]: HTTP Page Failure {res.status_code}")
        except Exception as e:
            log_details.append(f"[{col}]: Failure ({str(e)})")
            
    status_entry["PDF Outcome"] = f"Downloaded {success_downloads} explicit web assets"
    status_entry["ZIP Outcome"] = "N/A"
    status_entry["Pipeline Log Summary"] = " | ".join(log_details)

def wait_for_selenium_files(download_dir, before_files, timeout=SELENIUM_WAIT_TIME):
    start = time.time()
    while time.time() - start < timeout:
        current_files = set(os.listdir(download_dir))
        new_files = current_files - before_files
        if new_files:
            if any(f.endswith(".crdownload") or f.endswith(".tmp") for f in current_files):
                time.sleep(0.5)
                continue
            valid_files = {f for f in new_files if not f.endswith(".crdownload") and not f.endswith(".tmp")}
            if valid_files:
                time.sleep(0.5)
                return valid_files
        time.sleep(0.5)
    return set()

def execute_selenium_headless_pipeline(link, file_name, car_folder, run_path, status_entry):
    status_entry["Pipeline Channel"] = "Selenium Headless Browser Engine"
    
    options = Options()
    options.add_experimental_option("prefs", {
        "download.default_directory": os.path.abspath(car_folder),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    })
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    
    driver = webdriver.Chrome(options=options)
    try:
        before_files = set(os.listdir(car_folder))
        driver.get(link)
        new_files = wait_for_selenium_files(car_folder, before_files)
        
        # Branch directly if a target file payload drops instantly
        if new_files:
            for file in new_files:
                src_path = os.path.join(car_folder, file)
                ext = os.path.splitext(file)[1]
                standardized_name = f"{file_name}{ext}"
                dest_path = os.path.join(car_folder, standardized_name)
                shutil.move(src_path, dest_path)
                
                if standardized_name.lower().endswith(".pdf"):
                    # Scan internal page structures via PyMuPDF vector extraction maps
                    doc = fitz.open(dest_path)
                    download_url = None
                    for page in doc:
                        for link_dict in page.get_links():
                            uri = link_dict.get("uri", "")
                            if uri and "action=download_case_images" in uri:
                                download_url = uri
                                break
                        if download_url: break
                    doc.close()
                    
                    if download_url:
                        before_zip = set(os.listdir(car_folder))
                        driver.get(download_url)
                        wait_for_selenium_files(car_folder, before_zip)
                    else:
                        # Falling back directly onto visible page container elements
                        driver.get(link)
                        try:
                            btn = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Download Images')]"))
                            )
                            before_zip = set(os.listdir(car_folder))
                            btn.click()
                            wait_for_selenium_files(car_folder, before_zip)
                        except Exception: pass
                    
                    status_entry["PDF Outcome"] = process_pdf_path_to_images(dest_path, car_folder)
                    os.remove(dest_path)
                
            extract_all_zips(car_folder)
            status_entry["ZIP Outcome"] = "Success (Extracted Framework Assets)"
        else:
            # Fallback workflow directly executing from base valuation domain landing maps
            driver.get(link)
            try:
                btn = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Download Images')]"))
                )
                before_zip = set(os.listdir(car_folder))
                btn.click()
                downloaded = wait_for_selenium_files(car_folder, before_zip)
                extract_all_zips(car_folder)
                
                for file in os.listdir(car_folder):
                    if file.lower().endswith(".pdf"):
                        p_path = os.path.join(car_folder, file)
                        process_pdf_path_to_images(p_path, car_folder)
                        os.remove(p_path)
                status_entry["PDF Outcome"] = "Success (Web Fallback)"
                status_entry["ZIP Outcome"] = "Success (Web Extracted)"
            except Exception as inner:
                status_entry["PDF Outcome"] = "Failed"
                status_entry["ZIP Outcome"] = f"No interactive triggers captured: {str(inner)}"
    except Exception as outer:
        status_entry["PDF Outcome"] = "Crash"
        status_entry["ZIP Outcome"] = f"Selenium Pipeline Error: {str(outer)}"
    finally:
        driver.quit()

# ==========================================
# 📊 CONTROL INTERFACE ENDPOINTS
# ==========================================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file_web():
    file = request.files.get('excel_file')
    if not file: return "No file selected."

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    run_folder_name = f"Unified_Run_{timestamp}"
    run_path = os.path.join(OUTPUT_BASE_FOLDER, run_folder_name)
    os.makedirs(run_path, exist_ok=True)

    excel_path = os.path.join(UPLOAD_FOLDER, secure_filename(file.filename))
    file.save(excel_path)

    df = pd.read_excel(excel_path)
    report_data = []

    # Map target rows gracefully across any custom layout variations
    cols_upper = [str(c).upper() for c in df.columns]
    name_col = next((df.columns[i] for i, c in enumerate(cols_upper) if c in ["FILE NAME", "CAR/FILE NAME", "AGREEMENT NO", "LOAN NO"]), "File Name")
    link_col = next((df.columns[i] for i, c in enumerate(cols_upper) if c in ["LINK", "INPUT LINK"]), "Link")
    link_columns = [col for col in df.columns if col != name_col and str(col).upper() not in ["AUCTION ID", "SR NO", "SRNO", link_col.upper()]]

    for index, row in df.iterrows():
        raw_name = row.get(name_col, f"Item_{index}")
        if pd.isna(raw_name): continue
        
        file_name = sanitize_filename(str(raw_name))
        primary_link = str(row.get(link_col, "")).strip()
        
        status_entry = {
            "Car/File Name": file_name,
            "Pipeline Channel": "Unknown",
            "PDF Outcome": "N/A",
            "ZIP Outcome": "N/A",
            "Attempts": 1,
            "Pipeline Log Summary": "Execution Completed",
            "Processed At": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        car_folder = os.path.join(run_path, file_name)
        os.makedirs(car_folder, exist_ok=True)
        
        # 🔀 ARCHITECTURE PLATFORM DYNAMIC ROUTING TREE
        if "autoinspekt.com" in primary_link:
            execute_autoinspekt_pipeline(primary_link, car_folder, status_entry)
            
        elif "adroitauto.in" in primary_link or any(str(row.get(c, "")).startswith("http") for c in link_columns):
            # Fallback path executing explicit DOM scraping if landing links maps onto TVS profiles
            execute_tvs_credit_pipeline(row, link_columns, car_folder, status_entry)
            
        elif primary_link.startswith("http"):
            # Triggering headless automated worker browsers for Bajaj Finserv / Adroit secure sessions
            execute_selenium_headless_pipeline(primary_link, file_name, car_folder, run_path, status_entry)
            
        else:
            status_entry["Pipeline Log Summary"] = "Skipped: Invalid or empty URI link path parameters."

        report_data.append(status_entry)

    if report_data:
        report_df = pd.DataFrame(report_data)
        report_df.to_excel(os.path.join(run_path, f"Global_Completion_Report_{timestamp}.xlsx"), index=False)

    return render_template('index.html', finished=True, run_id=run_folder_name, folder=os.path.abspath(run_path))

if __name__ == '__main__':
    app.run(debug=True)
