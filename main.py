# --- KONFIGURASI UTAMA ---
import json
import re
import time
import pandas as pd
import base64
import random
import logging
from logging.handlers import RotatingFileHandler
import os
import requests
import sys
import shutil
import glob
from datetime import datetime, timedelta
from dotenv import load_dotenv
import pyotp

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# --- LOAD ENVIRONMENT VARIABLES ---
load_dotenv()

USERNAME = os.getenv("BPS_USERNAME")
PASSWORD = os.getenv("BPS_PASSWORD")
OTP_SECRET = os.getenv("BPS_OTP_SECRET")
# Default True jika tidak ada setting
USE_SESSION_CACHE = os.getenv("USE_SESSION_CACHE", "true").lower() == "true"
HEADLESS_MODE = os.getenv("HEADLESS", "true").lower() == "true"

if not USERNAME or not PASSWORD:
    print("ERROR: Kredensial (BPS_USERNAME, BPS_PASSWORD) tidak ditemukan di file .env")
    sys.exit(1)

# --- SETUP LOGGING ---
# Get the root logger
logger = logging.getLogger()
logger.setLevel(logging.DEBUG) # Set overall level to DEBUG to capture all messages for file

# Clear existing handlers if any (important for re-running in IDE or interactive sessions)
if logger.handlers:
    for handler in logger.handlers:
        logger.removeHandler(handler)

# File Handler (detailed)
# Menggunakan RotatingFileHandler: Max 5MB per file, simpan 3 file backup terakhir
file_handler = RotatingFileHandler("app.log", maxBytes=5*1024*1024, backupCount=3, encoding='utf-8')
file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(file_formatter)
file_handler.setLevel(logging.DEBUG) # File handler captures DEBUG and above
logger.addHandler(file_handler)

# Stream Handler (console - concise with timestamp)
console_handler = logging.StreamHandler(sys.stdout)
console_formatter = logging.Formatter('[%(asctime)s] %(message)s', datefmt='%H:%M:%S') # Only show the message with timestamp
console_handler.setFormatter(console_formatter)
console_handler.setLevel(logging.INFO) # Console handler only shows INFO and above
logger.addHandler(console_handler)


CUSTOM_USER_AGENT = 'Mozilla/5.0 (Linux; Android 13; itel A666LN Build/TP1A.220624.014; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/143.0.7499.192 Mobile Safari/537.36'

DIR_URL = "https://matchapro.web.bps.go.id/dirgc"
POST_URL = 'https://matchapro.web.bps.go.id/dirgc/konfirmasi-user'

SESSION_FILE = 'session.json'
INPUT_DIR = 'input'
BACKUP_DIR = 'backup'
PROCESSED_DIR = 'processed'
BOUNDING_BOX_FILE = 'bounding_boxes.json'


def get_driver():
    """Menginisialisasi dan mengembalikan driver Selenium."""
    logging.info("Menginisialisasi Chrome Driver...")
    chrome_options = Options()
    chrome_options.add_argument(f'user-agent={CUSTOM_USER_AGENT}')
    if HEADLESS_MODE:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--log-level=3")
    # Anti-detection
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def save_session_data(driver):
    """Menyimpan data sesi (cookies & csrf) dan mengembalikan gc_token dari driver yang aktif."""
    logging.debug("Mengambil cookie dan token CSRF dari browser...") # Changed to debug
    time.sleep(2)
    cookies = driver.get_cookies()
    page_source = driver.page_source

    # Mencari CSRF Token
    match = re.search(r'<meta name="csrf-token" content="([^"]+)">', page_source)
    csrf_token = match.group(1) if match else None

    if not csrf_token:
        logging.warning("Tidak dapat menemukan token CSRF di halaman.")

    # Mencari gcSubmitToken
    logging.debug("Mencari gcSubmitToken di source code halaman...") # Changed to debug
    gc_token_match = re.search(r"gcSubmitToken\s*=\s*['\"]([^'\"]+)['\"]", page_source)
    gc_token = None
    if gc_token_match:
        gc_token = gc_token_match.group(1)
        logging.debug(f"Ditemukan gcSubmitToken: {gc_token}") # Changed to debug
    else:
        logging.warning("gcSubmitToken tidak ditemukan di halaman.")

    session_data = None
    if csrf_token:
        session_data = {'cookies': cookies, 'csrf_token': csrf_token}

        if USE_SESSION_CACHE:
            with open(SESSION_FILE, 'w') as f:
                json.dump(session_data, f)
            logging.info(f"Sesi berhasil diperbarui dan disimpan di '{SESSION_FILE}'.")
        else:
            logging.info("Sesi diperbarui (Tidak disimpan ke file karena USE_SESSION_CACHE=false).")

    return session_data, gc_token


def login_selenium(driver):
    """Melakukan proses login."""
    logging.info("Membuka halaman login...")
    
    while True:
        # Retry logic for connection reset
        max_retries = 3
        success = False
        for attempt in range(max_retries):
            try:
                driver.get(DIR_URL)
                success = True
                break
            except WebDriverException as e:
                if "ERR_CONNECTION_RESET" in str(e) or "ERR_CONNECTION_CLOSED" in str(e):
                    logging.warning(f"Koneksi terputus (ERR_CONNECTION_RESET). Mencoba ulang... ({attempt + 1}/{max_retries})")
                    time.sleep(5)
                else:
                    raise e
        
        if success:
            break
            
        print("\n" + "!" * 50)
        print("GAGAL MENGHUBUNGI SERVER (ERR_CONNECTION_RESET)")
        print("Silakan cek koneksi internet Anda atau pastikan VPN FortiClient sudah terhubung.")
        print("!" * 50 + "\n")
        # Bunyikan beep
        print('\a')
        
        choice = input("Ketik 'y' untuk mencoba lagi, atau 'n' untuk keluar: ").strip().lower()
        if choice == 'n':
            logging.error("User memilih untuk keluar aplikasi karena masalah koneksi.")
            sys.exit(1)
        else:
            logging.info("User memilih untuk mencoba koneksi lagi...")

    # Cek jika sudah login
    if driver.current_url == DIR_URL or "Sign in" not in driver.page_source:
        try:
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Sign in with SSO BPS')]")))
        except:
            logging.info("Terdeteksi sudah dalam keadaan login.")
            return

    logging.info("Melakukan klik tombol Sign in with SSO BPS...")
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Sign in with SSO BPS')]"))).click()

    logging.info("Memasukkan kredensial...")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username"))).send_keys(USERNAME)
    driver.find_element(By.ID, "password").send_keys(PASSWORD)
    driver.find_element(By.XPATH, "//input[@type='submit']").click()

    # --- OTP HANDLING ---
    logging.info("Menunggu respons login (OTP atau Redirect)...")
    
    otp_xpath = "//input[contains(@name, 'token') or contains(@id, 'token') or contains(@name, 'otp') or contains(@id, 'otp')]"
    
    try:
        # Tunggu sampai URL adalah DIR_URL ATAU elemen OTP ditemukan
        WebDriverWait(driver, 10).until(
            lambda d: d.current_url == DIR_URL or len(d.find_elements(By.XPATH, otp_xpath)) > 0
        )
    except Exception:
        logging.warning("Timeout menunggu transisi halaman setelah login. Memeriksa kondisi terakhir...")

    if driver.current_url == DIR_URL:
        logging.info("Login berhasil tanpa OTP.")
    else:
        otp_elements = driver.find_elements(By.XPATH, otp_xpath)
        if otp_elements:
            logging.info("Halaman OTP terdeteksi!")
            otp_field = otp_elements[0]

            otp_code = None
            if OTP_SECRET:
                try:
                    totp = pyotp.TOTP(OTP_SECRET)
                    otp_code = totp.now()
                    logging.info("OTP dihasilkan otomatis dari secret key.")
                except Exception as e:
                    logging.error(f"Gagal generate OTP: {e}")

            if not otp_code:
                print("\n" + "!" * 50)
                print("MASUKKAN KODE OTP SECARA MANUAL!")
                print("!" * 50 + "\n")
                # Bunyikan beep sistem agar user sadar (opsional, hanya work di beberapa terminal)
                print('\a')
                otp_code = input("Masukkan Kode OTP: ").strip()

            logging.info("Menginput kode OTP...")
            otp_field.send_keys(otp_code)

            # Cari tombol submit OTP (biasanya type submit atau button dengan text Sign in/Verifikasi)
            # Kita coba enter saja di field atau cari tombol
            try:
                otp_field.submit()
            except:
                driver.find_element(By.XPATH, "//input[@type='submit'] or //button[@type='submit']").click()

            WebDriverWait(driver, 20).until(EC.url_to_be(DIR_URL))
            logging.info("Login berhasil setelah OTP.")
        else:
            logging.warning(f"Tidak terdeteksi OTP dan belum masuk ke halaman utama. URL saat ini: {driver.current_url}")


def get_authenticated_session_selenium():
    """Fungsi wrapper untuk login penuh. Mengembalikan (session_data, gc_token)."""
    logging.info("--- MEMULAI OTENTIKASI BARU DENGAN SELENIUM ---")
    driver = get_driver()
    try:
        login_selenium(driver)
        return save_session_data(driver)
    finally:
        if driver:
            driver.quit()


def refresh_gc_token_selenium():
    """Mencoba refresh halaman untuk dapat token baru. Login ulang jika perlu. Mengembalikan (session_data, gc_token)."""
    logging.info("--- REFRESH TOKEN DENGAN SELENIUM ---")
    driver = get_driver()
    try:
        if os.path.exists(SESSION_FILE):
            try:
                with open(SESSION_FILE, 'r') as f:
                    old_session = json.load(f)

                driver.get(DIR_URL)
                for cookie in old_session.get('cookies', []):
                    try:
                        driver.add_cookie(cookie)
                    except:
                        pass
            except Exception as e:
                logging.error(f"Gagal load cookie lama: {e}")

        logging.debug(f"Membuka {DIR_URL} untuk cek token...") # Changed to debug
        
        while True:
            # Retry logic for connection reset
            max_retries = 3
            success = False
            for attempt in range(max_retries):
                try:
                    driver.get(DIR_URL)
                    success = True
                    break
                except WebDriverException as e:
                    if "ERR_CONNECTION_RESET" in str(e) or "ERR_CONNECTION_CLOSED" in str(e):
                        logging.warning(f"Koneksi terputus saat refresh token. Mencoba ulang... ({attempt + 1}/{max_retries})")
                        time.sleep(5)
                    else:
                        logging.error(f"Error Selenium tak terduga: {e}")
                        return None, None
            
            if success:
                break
                
            print("\n" + "!" * 50)
            print("GAGAL MENGHUBUNGI SERVER SAAT REFRESH TOKEN")
            print("Silakan cek koneksi internet Anda atau pastikan VPN FortiClient sudah terhubung.")
            print("!" * 50 + "\n")
            print('\a')
            
            choice = input("Ketik 'y' untuk mencoba lagi, atau 'n' untuk membatalkan refresh: ").strip().lower()
            if choice == 'n':
                logging.error("User membatalkan refresh token.")
                return None, None
            else:
                logging.info("User memilih untuk mencoba refresh lagi...")

        time.sleep(3)

        page_source = driver.page_source
        if "gcSubmitToken" in page_source:
            logging.debug("gcSubmitToken ditemukan tanpa perlu login ulang.") # Changed to debug
            return save_session_data(driver)
        else:
            logging.warning("gcSubmitToken tidak ditemukan. Kemungkinan sesi habis. Melakukan login ulang...")
            login_selenium(driver)
            return save_session_data(driver)

    finally:
        if driver:
            driver.quit()


def load_session_from_file():
    """Mencoba memuat sesi dari file."""
    if not USE_SESSION_CACHE:
        logging.info("USE_SESSION_CACHE=false, melewati pemuatan sesi dari file.")
        return None

    if os.path.exists(SESSION_FILE):
        logging.info(f"Mencoba memuat sesi dari file '{SESSION_FILE}'...")
        with open(SESSION_FILE, 'r') as f:
            return json.load(f)
    return None


def create_backup(file_path):
    """Membuat backup file Excel ke folder backup/."""
    try:
        # Buat folder backup jika belum ada
        if not os.path.exists(BACKUP_DIR):
            os.makedirs(BACKUP_DIR)
            logging.info(f"Folder backup dibuat: {BACKUP_DIR}")

        # Nama file backup dengan timestamp
        filename = os.path.basename(file_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"{os.path.splitext(filename)[0]}_{timestamp}{os.path.splitext(filename)[1]}"
        backup_path = os.path.join(BACKUP_DIR, backup_filename)

        shutil.copy2(file_path, backup_path)
        logging.info(f"Backup file dibuat: {backup_path}")
    except Exception as e:
        logging.error(f"Gagal membuat backup untuk {file_path}: {e}")


def load_bounding_boxes():
    """Memuat data bounding box dari file JSON."""
    if not os.path.exists(BOUNDING_BOX_FILE):
        logging.warning(f"File '{BOUNDING_BOX_FILE}' tidak ditemukan. Validasi lokasi dilewati.")
        return {}
    
    try:
        with open(BOUNDING_BOX_FILE, 'r') as f:
            data = json.load(f)
        
        # Mapping kode kabupaten (2 digit) ke bounding box
        bbox_map = {}
        for key, bbox in data.items():
            # Ekstrak kode lengkap (misal: 13301 dari final_desa_202413301.geojson)
            # Regex diperbaiki untuk menangkap semua digit setelah '2024' dan sebelum '.geojson'
            match = re.search(r'2024(\d+)\.geojson', key) 
            if match:
                full_code = match.group(1) # e.g., "13301"
                if len(full_code) >= 2:
                    kab_code = full_code[-2:] # Take last 2 digits, e.g., "01" from "13301"
                    bbox_map[kab_code] = bbox
                else:
                    logging.warning(f"Kode wilayah '{full_code}' dari '{key}' terlalu pendek, dilewati.")
            else:
                logging.warning(f"Tidak dapat mengekstrak kode wilayah dari '{key}', dilewati.")
        
        logging.info(f"Berhasil memuat {len(bbox_map)} data bounding box wilayah (2 digit).")
        return bbox_map
    except Exception as e:
        logging.error(f"Gagal memuat file bounding box: {e}")
        return {}


def print_validation_rules():
    """Menampilkan aturan validasi ke console/log."""
    rules = """
    =======================================================
    ATURAN VALIDASI DATA:
    1. perusahaan_id : Wajib terisi.
    2. kdkab         : Wajib terisi (2 digit).
    3. hasilgc       : Harus salah satu dari ['1', '3', '4', '99'].
    4. edit_nama     : Harus '0' atau '1'.
    5. edit_alamat   : Harus '0' atau '1'.
    6. Konsistensi   :
       - Jika nama_usaha terisi, edit_nama harus '1'.
       - Jika nama_usaha kosong, edit_nama harus '0'.
       - Jika alamat_usaha terisi, edit_alamat harus '1'.
       - Jika alamat_usaha kosong, edit_alamat harus '0'.
    7. Lokasi        : Latitude & Longitude harus berada dalam
                       wilayah kabupaten (berdasarkan kolom kdkab).
    =======================================================
    """
    print(rules)
    input("Tekan Enter untuk melanjutkan...")
    logging.info("Aturan validasi ditampilkan dan disetujui user.")


def get_input_files():
    """Mencari semua file Excel di folder input."""
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)
        logging.info(f"Folder '{INPUT_DIR}' dibuat. Silakan letakkan file Excel di dalamnya.")
        return []

    files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx")) + glob.glob(os.path.join(INPUT_DIR, "*.xls"))
    # Filter file temporary (yang dimulai dengan ~$)
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    return files


def validate_row_data(row, bbox_map):
    """Melakukan validasi logika bisnis pada satu baris data."""
    perusahaan_id_val = str(row['perusahaan_id']).strip()
    kdkab_val = str(row['kdkab']).strip()
    hasilgc_val = str(row['hasilgc']).replace('.0', '').strip()
    edit_nama_val = str(row['edit_nama']).replace('.0', '').strip()
    edit_alamat_val = str(row['edit_alamat']).replace('.0', '').strip()
    nama_usaha_val = str(row['nama_usaha']).strip()
    alamat_usaha_val = str(row['alamat_usaha']).strip()
    lat_val = str(row['latitude']).strip()
    long_val = str(row['longitude']).strip()

    validation_errors = []

    if not perusahaan_id_val:
        validation_errors.append("perusahaan_id kosong")
    if not kdkab_val:
        validation_errors.append("kdkab kosong")
    elif len(kdkab_val) != 2:
        # Opsional: Jika ingin strict 2 digit, uncomment baris bawah. 
        # Saat ini hanya warning jika tidak 2 digit tapi tetap lanjut cek bbox
        validation_errors.append(f"kdkab harus 2 digit (ditemukan: {kdkab_val})")
        pass

    valid_hasilgc = ['1', '3', '4', '99']
    if hasilgc_val not in valid_hasilgc:
        validation_errors.append(f"hasilgc invalid ({hasilgc_val}), harus {valid_hasilgc}")

    valid_flag = ['0', '1']
    if edit_nama_val not in valid_flag:
        validation_errors.append(f"edit_nama invalid ({edit_nama_val}), harus {valid_flag}")
    if edit_alamat_val not in valid_flag:
        validation_errors.append(f"edit_alamat invalid ({edit_alamat_val}), harus {valid_flag}")

    if nama_usaha_val and edit_nama_val != '1':
        validation_errors.append("nama_usaha terisi tapi edit_nama bukan 1")
    elif not nama_usaha_val and edit_nama_val != '0':
        validation_errors.append("nama_usaha kosong tapi edit_nama bukan 0")

    if alamat_usaha_val and edit_alamat_val != '1':
        validation_errors.append("alamat_usaha terisi tapi edit_alamat bukan 1")
    elif not alamat_usaha_val and edit_alamat_val != '0':
        validation_errors.append("alamat_usaha kosong tapi edit_alamat bukan 0")

    # --- VALIDASI LOKASI ---
    if bbox_map and kdkab_val:
        try:
            # Pastikan format kdkab 2 digit (misal '1' jadi '01') jika perlu, 
            # tapi asumsi data excel sudah string '01'
            kab_code = kdkab_val.zfill(2) # Memastikan 2 digit, misal "1" jadi "01"

            if kab_code in bbox_map:
                bbox = bbox_map[kab_code]
                min_long, min_lat, max_long, max_lat = bbox
                curr_lat = float(lat_val)
                curr_long = float(long_val)

                if not (min_lat <= curr_lat <= max_lat):
                    validation_errors.append(f"Lat ({curr_lat}) di luar {kab_code}")
                if not (min_long <= curr_long <= max_long):
                    validation_errors.append(f"Long ({curr_long}) di luar {kab_code}")
            else:
                validation_errors.append(f"Kode kab {kab_code} tidak ada di bbox map.")
        except ValueError:
            validation_errors.append("Format Lat/Long invalid.")
        except Exception as e:
            validation_errors.append(f"Error validasi lokasi: {str(e)}")

    return validation_errors

def process_file(file_path, session, post_headers, gc_token, csrf_token, bbox_map):
    """Memproses satu file Excel."""
    filename_short = os.path.basename(file_path)
    logging.info(f"Memproses file: {filename_short}") # Concise for console

    df = None
    while True:
        try:
            logging.debug(f"Membaca file data: {file_path}") # Changed to debug
            df = pd.read_excel(file_path, dtype=str)
            df = df.fillna('')
            df = df.replace('nan', '')
            
            # Bersihkan nama kolom (hapus spasi di awal/akhir)
            df.columns = df.columns.str.strip()

            if 'status_upload' not in df.columns:
                df['status_upload'] = ''
            
            # Jika berhasil baca, keluar dari loop
            break

        except PermissionError:
            print("\n" + "!" * 50)
            print(f"ERROR: File '{os.path.basename(file_path)}' sedang dibuka!")
            print("Mohon TUTUP file Excel tersebut agar program bisa membacanya.")
            print("!" * 50 + "\n")
            print('\a') # Beep
            input("Tekan ENTER jika sudah menutup file untuk mencoba lagi...")
        except Exception as e:
            logging.error(f"Gagal membaca file Excel {file_path}: {e}", exc_info=True)
            return gc_token, None  # Return token yang ada dan None untuk stats

    # 1. Buat Backup (Hanya jika file berhasil dibaca)
    create_backup(file_path)

    required_columns = [
        'perusahaan_id', 'kdkab', 'latitude', 'longitude', 'hasilgc',
        'edit_nama', 'edit_alamat', 'nama_usaha', 'alamat_usaha'
    ]

    if not all(col in df.columns for col in required_columns):
        logging.error(f"Kolom di Excel {file_path} tidak lengkap. Harus ada: {', '.join(required_columns)}")
        return gc_token, None

    total_data = len(df)
    logging.info(f"[{filename_short}] Total {total_data} baris data.") # Concise for console
    
    # --- STATISTIK ---
    stats = {
        'filename': filename_short,
        'total': total_data,
        'success': 0,
        'failed': 0,
        'skipped': 0,
        'start_time': datetime.now()
    }

    SAVE_BATCH_SIZE = 10  # Simpan ke Excel setiap 10 baris untuk performa

    try:
        for index, row in df.iterrows():
            current_num = index + 1
            progress_pct = (current_num / total_data) * 100

            status_upload_lower = str(row.get('status_upload', '')).lower()
            if status_upload_lower == 'berhasil' or 'sudah diground check oleh user lain' in status_upload_lower:
                logging.info(f"[{filename_short}] Baris {current_num}/{total_data} ({progress_pct:.2f}%) - Status '{row.get('status_upload', '')}', dilewati.")
                stats['skipped'] += 1
                continue

            logging.info(f"[{filename_short}] Baris {current_num}/{total_data} ({progress_pct:.2f}%) - Memproses...") # Concise for console

            # --- VALIDASI DATA (Refactored) ---
            validation_errors = validate_row_data(row, bbox_map)
            
            if validation_errors:
                error_msg = "Invalid: " + "; ".join(validation_errors)
                logging.warning(f"[{filename_short}] Baris {current_num}/{total_data} ({progress_pct:.2f}%) - Gagal Validasi: {error_msg}") # Concise for console

                df.at[index, 'status_upload'] = error_msg
                stats['failed'] += 1
                # Tidak langsung save setiap error, tunggu batch
                continue

            # Persiapan Data untuk Request
            perusahaan_id_val = str(row['perusahaan_id']).strip()
            hasilgc_val = str(row['hasilgc']).replace('.0', '').strip()
            edit_nama_val = str(row['edit_nama']).replace('.0', '').strip()
            edit_alamat_val = str(row['edit_alamat']).replace('.0', '').strip()
            nama_usaha_val = str(row['nama_usaha']).strip()
            alamat_usaha_val = str(row['alamat_usaha']).strip()

            if edit_nama_val == '1':
                logging.debug(f"Encoding nama_usaha '{nama_usaha_val}' ke Base64...") # Changed to debug
                nama_usaha_val = base64.b64encode(nama_usaha_val.encode('utf-8')).decode('utf-8')

            if edit_alamat_val == '1':
                logging.debug(f"Encoding alamat_usaha '{alamat_usaha_val}' ke Base64...") # Changed to debug
                alamat_usaha_val = base64.b64encode(alamat_usaha_val.encode('utf-8')).decode('utf-8')

            time_on_page_val = str(random.randint(30, 120))

            data = {
                'perusahaan_id': perusahaan_id_val,
                'latitude': str(row['latitude']),
                'longitude': str(row['longitude']),
                'hasilgc': hasilgc_val,
                'gc_token': gc_token,
                'edit_nama': edit_nama_val,
                'edit_alamat': edit_alamat_val,
                'nama_usaha': nama_usaha_val,
                'alamat_usaha': alamat_usaha_val,
                'time_on_page': time_on_page_val,
                '_token': csrf_token,
            }

            logging.debug(f"Mengirim data: perusahaan_id={data['perusahaan_id']}, time_on_page={time_on_page_val}") # Changed to debug

            retry_count = 0
            max_retries = 1
            status_akhir = "gagal"
            response = None # Initialize response

            while retry_count <= max_retries:
                try:
                    response = session.post(POST_URL, headers=post_headers, data=data, timeout=30)
                    logging.debug(f"Status Code: {response.status_code}") # Changed to debug

                    if response.status_code == 200:
                        try:
                            response_json = response.json()
                            msg = response_json.get('message', 'No message')
                            logging.debug(f"Response Message: {msg}") # Changed to debug

                            if response_json.get('status') == 'success' and 'new_gc_token' in response_json:
                                new_token = response_json['new_gc_token']
                                gc_token = new_token
                                logging.debug(f"[SUCCESS] Token diperbarui: {new_token[:10]}...") # Changed to debug
                                status_akhir = "berhasil"
                                break
                            else:
                                logging.warning(f"[INFO] Gagal: {msg}")
                                status_akhir = f"gagal - {msg}"
                                logging.debug(f"Full Response: {response.text}") # Changed to debug
                                break
                        except json.JSONDecodeError:
                            logging.error("Gagal memparsing respons sebagai JSON.")
                            logging.debug(f"Response Text: {response.text}") # Changed to debug
                            break
                    elif response.status_code == 429:
                        logging.warning("Rate limit (429) terdeteksi. Mencoba menunggu...")
                        retry_after = response.headers.get('Retry-After')
                        wait_time = 30  # Fallback jika header tidak ada
                        if retry_after and retry_after.isdigit():
                            wait_time = int(retry_after) + 1 # Tambah 1 detik buffer
                        
                        logging.info(f"[{filename_short}] Baris {current_num}/{total_data} ({progress_pct:.2f}%) - Rate Limit (429). Menunggu {wait_time} detik...") # Concise for console
                        time.sleep(wait_time)
                        # Coba lagi request yang sama tanpa menambah retry_count
                        continue 
                    elif response.status_code == 400:
                        try:
                            response_json = response.json()
                            msg = response_json.get('message', '')

                            if "Token invalid atau sudah terpakai" in msg:
                                logging.warning("Token invalid. Mencoba refresh token dengan Selenium...")

                                session_data, new_gc_token = refresh_gc_token_selenium()

                                if session_data and new_gc_token:
                                    csrf_token = session_data['csrf_token']
                                    session = requests.Session()
                                    for cookie in session_data['cookies']:
                                        session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'])

                                    gc_token = new_gc_token
                                    data['gc_token'] = gc_token
                                    data['_token'] = csrf_token

                                    # Jangan increment retry_count di sini agar request ulang dianggap percobaan baru
                                    logging.info("Mencoba mengirim ulang request dengan token baru...")
                                    continue
                                else:
                                    logging.error("Gagal mendapatkan sesi atau token baru.")
                                    status_akhir = "gagal - Refresh token error"
                                    break
                            else:
                                logging.error(f"Request gagal (400): {msg}")
                                status_akhir = f"gagal - {msg}"
                                break
                        except:
                            logging.error("Request gagal (400).")
                            logging.debug(response.text) # Changed to debug
                            status_akhir = "gagal - 400 Bad Request"
                            break
                    else:
                        logging.error(f"Request gagal dengan status {response.status_code}.")
                        logging.debug(response.text) # Changed to debug
                        status_akhir = f"gagal - HTTP {response.status_code}"
                        break

                except requests.exceptions.Timeout:
                    logging.error("Request Timeout (30s). Server tidak merespons.")
                    retry_count += 1 # Increment retry count on timeout
                    if retry_count > max_retries:
                        status_akhir = "gagal - Timeout"
                        break
                except Exception as e:
                    logging.error(f"Terjadi kesalahan saat melakukan request: {e}", exc_info=True)
                    retry_count += 1 # Increment retry count on other exceptions
                    if retry_count > max_retries:
                        status_akhir = f"gagal - Error: {str(e)}"
                        break

            df.at[index, 'status_upload'] = status_akhir
            
            if status_akhir == "berhasil":
                stats['success'] += 1
            else:
                stats['failed'] += 1

            # --- BATCH SAVING ---
            # Hanya simpan jika kelipatan batch atau ini adalah data terakhir
            if current_num % SAVE_BATCH_SIZE == 0 or current_num == total_data:
                save_success = False
                for attempt in range(3):
                    try:
                        df.to_excel(file_path, index=False)
                        logging.info(f"[{filename_short}] Menyimpan progress batch ke Excel...")
                        save_success = True
                        break
                    except PermissionError:
                        logging.warning(f"File Excel terkunci. Retry save ({attempt + 1}/3)...")
                        time.sleep(2)
                    except Exception as e:
                        logging.error(f"Gagal menyimpan file Excel: {e}")
                        break
                if not save_success:
                    logging.error("GAGAL MENYIMPAN BATCH KE EXCEL.")
            else:
                logging.info(f"[{filename_short}] Status: {status_akhir} (Menunggu batch save)")

            # Hanya tidur jika request sebelumnya berhasil atau gagal permanen
            # Tidak perlu tidur jika baru saja menunggu karena 429
            if response and response.status_code != 429:
                sleep_time = random.uniform(1, 3)
                time.sleep(sleep_time)

    except KeyboardInterrupt:
        logging.warning("\n!!! PROSES DIHENTIKAN OLEH PENGGUNA (Ctrl+C) !!!")
        logging.info("Menyimpan data terakhir sebelum keluar...")
        try:
            df.to_excel(file_path, index=False)
            logging.info("Data berhasil disimpan.")
        except Exception as e:
            logging.error(f"Gagal menyimpan data saat exit: {e}")
        
        # Tetap return stats agar laporan bisa dibuat
        stats['end_time'] = datetime.now()
        return gc_token, stats
    
    stats['end_time'] = datetime.now()
    return gc_token, stats

def generate_summary_report(all_stats):
    """Membuat dan menampilkan laporan ringkasan."""
    if not all_stats:
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_filename = f"summary_report_{timestamp}.txt"
    
    lines = []
    lines.append("=" * 60)
    lines.append(f"RINGKASAN EKSEKUSI MATCHAIN GC - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 60)
    lines.append("")

    total_all_files = 0
    total_all_success = 0
    total_all_failed = 0
    total_all_skipped = 0

    for stats in all_stats:
        duration = stats['end_time'] - stats['start_time']
        duration_str = str(duration).split('.')[0] # Remove microseconds

        lines.append(f"FILE: {stats['filename']}")
        lines.append(f"  - Total Data    : {stats['total']}")
        lines.append(f"  - Berhasil      : {stats['success']}")
        lines.append(f"  - Gagal         : {stats['failed']}")
        lines.append(f"  - Dilewati      : {stats['skipped']}")
        lines.append(f"  - Durasi        : {duration_str}")
        lines.append("-" * 40)

        total_all_files += 1
        total_all_success += stats['success']
        total_all_failed += stats['failed']
        total_all_skipped += stats['skipped']

    lines.append("")
    lines.append("=" * 60)
    lines.append("TOTAL KESELURUHAN")
    lines.append(f"  - Jumlah File   : {total_all_files}")
    lines.append(f"  - Total Berhasil: {total_all_success}")
    lines.append(f"  - Total Gagal   : {total_all_failed}")
    lines.append(f"  - Total Dilewati: {total_all_skipped}")
    lines.append("=" * 60)

    report_content = "\n".join(lines)

    # Tampilkan di console
    print("\n" + report_content + "\n")

    # Simpan ke file
    try:
        with open(report_filename, "w") as f:
            f.write(report_content)
        logging.info(f"Laporan ringkasan disimpan di: {report_filename}")
    except Exception as e:
        logging.error(f"Gagal menyimpan laporan ringkasan: {e}")


def main():
    """Fungsi utama untuk menjalankan scraper."""
    print("\n" + "=" * 50)
    print("   MatchaIn GC (Matcha Input Gak Culun)")
    print("=" * 50 + "\n")
    logging.info("Aplikasi dimulai.")

    # Load bounding boxes
    bbox_map = load_bounding_boxes()

    print_validation_rules()

    # Bersihkan sesi lama jika cache dimatikan
    if not USE_SESSION_CACHE and os.path.exists(SESSION_FILE):
        try:
            os.remove(SESSION_FILE)
            logging.info("Sesi lama dihapus karena USE_SESSION_CACHE=false.")
        except:
            pass

    input_files = get_input_files()
    if not input_files:
        logging.warning(f"Tidak ada file Excel (.xlsx/.xls) ditemukan di folder '{INPUT_DIR}'.")
        return

    # Inisialisasi variabel
    session_data = load_session_from_file()
    gc_token = None

    if not session_data:
        session_data, gc_token = get_authenticated_session_selenium()
    else:
        logging.info("Sesi dimuat dari file, mengambil gc_token awal via Selenium...")
        session_data, gc_token = refresh_gc_token_selenium()

    if not session_data:
        logging.critical("Gagal mendapatkan sesi otentikasi. Proses dihentikan.")
        return

    if not gc_token:
        logging.critical("Gagal mendapatkan gc_token awal. Proses dihentikan.")
        return

    csrf_token = session_data['csrf_token']
    session = requests.Session()
    for cookie in session_data['cookies']:
        session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'])

    post_headers = {
        'Accept': '*/*', 'Origin': 'https://matchapro.web.bps.go.id', 'Referer': DIR_URL,
        'User-Agent': CUSTOM_USER_AGENT, 'X-Requested-With': 'XMLHttpRequest'
    }

    all_files_stats = []

    # Proses setiap file
    for file_path in input_files:
        # --- Cek apakah file sedang dibuka ---
        lock_file_path = os.path.join(os.path.dirname(file_path), "~$" + os.path.basename(file_path))
        while os.path.exists(lock_file_path):
            print("\n" + "!" * 50)
            print(f"PERINGATAN: File '{os.path.basename(file_path)}' terdeteksi sedang dibuka.")
            print("Mohon TUTUP file Excel tersebut sebelum melanjutkan.")
            print("!" * 50 + "\n")
            print('\a') # Beep
            input("Tekan ENTER jika sudah menutup file untuk melanjutkan...")
            logging.info(f"Menunggu user menutup file: {file_path}")

        gc_token, stats = process_file(file_path, session, post_headers, gc_token, csrf_token, bbox_map)
        if stats:
            all_files_stats.append(stats)
            
            # --- AUTO-MOVE COMPLETED FILES ---
            # Pindahkan jika (success + skipped) sama dengan total data (artinya 100% selesai)
            if (stats['success'] + stats['skipped'] == stats['total']) and stats['total'] > 0:
                try:
                    if not os.path.exists(PROCESSED_DIR):
                        os.makedirs(PROCESSED_DIR)
                    dest_path = os.path.join(PROCESSED_DIR, os.path.basename(file_path))
                    shutil.move(file_path, dest_path)
                    logging.info(f"File '{os.path.basename(file_path)}' SELESAI 100% dan dipindahkan ke '{PROCESSED_DIR}'.")
                except Exception as e:
                    logging.error(f"Gagal memindahkan file selesai: {e}")

    generate_summary_report(all_files_stats)
    logging.info("Semua proses selesai.")


if __name__ == "__main__":
    main()
