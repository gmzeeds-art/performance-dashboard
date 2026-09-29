import os
import re
import io
import json
import time
import glob
import shutil
from datetime import datetime

import ddddocr
import git
from playwright.sync_api import sync_playwright

# =============================================================
# KONFIGURASI
# =============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAWDATA_DIR = os.path.join(BASE_DIR, "rawdata")
TEMP_DOWNLOAD_DIR = os.path.join(BASE_DIR, "temp_downloads")
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")  # JANGAN di-push ke GitHub!
os.makedirs(RAWDATA_DIR, exist_ok=True)
os.makedirs(TEMP_DOWNLOAD_DIR, exist_ok=True)

HEADLESS = False          # Ubah ke True kalau sudah stabil (browser tidak tampil)
MAX_LOGIN_ATTEMPT = 8     # Jumlah percobaan login (captcha kadang salah baca)
KEEPALIVE_ROUNDS = 6      # 6 x 150 detik = download ulang tiap ~15 menit
KEEPALIVE_SECONDS = 150

WEB_CONFIGS = [
    {
        "key": "SK",
        "name": "PinjamID (SK)",
        "url": "https://skconsole.pinjamid.com/collection/workdata/WorkStatistics",
        "file_prefix": "workStatistics_SK",
    },
    {
        "key": "LD",
        "name": "LumbungDana (LD)",
        "url": "https://collection.lumbungdana.co.id/collection/workdata/WorkStatistics",
        "file_prefix": "workStatistics_LD",
    },
]

ocr = ddddocr.DdddOcr(show_ad=False)


def load_credentials():
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(
            "File credentials.json tidak ditemukan. Buat file tersebut di folder project "
            "(lihat panduan)."
        )
    with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


# =============================================================
# LOGIN + CAPTCHA
# =============================================================
def is_login_page(page):
    try:
        return page.locator("input[type='password']:visible").count() > 0
    except Exception:
        return False


def find_captcha_image(page):
    """Cari elemen <img> captcha (ukuran kecil, sekitar 80-200 x 25-60 px)."""
    imgs = page.locator("img:visible")
    for i in range(imgs.count()):
        el = imgs.nth(i)
        box = el.bounding_box()
        if box and 60 <= box["width"] <= 250 and 20 <= box["height"] <= 80:
            return el
    return None


def solve_math_captcha(img_bytes):
    """
    Captcha berbentuk soal matematika, contoh: '3*2=?'
    OCR membaca gambar -> kita ambil angka & operator -> hitung hasilnya.
    """
    raw = ocr.classification(img_bytes)
    text = raw.lower().replace(" ", "")
    text = text.replace("x", "*").replace("×", "*").replace("÷", "/").replace("—", "-")
    m = re.search(r"(\d+)([+\-*/])(\d+)", text)
    if not m:
        return raw, None
    a, op, b = int(m.group(1)), m.group(2), int(m.group(3))
    try:
        if op == "+":
            result = a + b
        elif op == "-":
            result = a - b
        elif op == "*":
            result = a * b
        else:
            result = a // b if b != 0 else None
    except Exception:
        result = None
    return raw, result


def login(page, web_info, creds):
    for attempt in range(1, MAX_LOGIN_ATTEMPT + 1):
        print(f" -> Login {web_info['name']} (percobaan {attempt}/{MAX_LOGIN_ATTEMPT})...")
        page.goto(web_info["url"], wait_until="domcontentloaded", timeout=60000)
        time.sleep(3)

        if not is_login_page(page):
            print(" -> Sudah dalam kondisi login.")
            return True

        # Isi username & password
        inputs = page.locator("input:visible")
        user_input = inputs.nth(0)
        pass_input = page.locator("input[type='password']:visible").first

        user_input.fill("")
        user_input.fill(creds["username"])
        pass_input.fill("")
        pass_input.fill(creds["password"])

        # Captcha
        cap_input = page.get_by_placeholder("Kode Verifikasi")
        if cap_input.count() == 0:
            cap_input = inputs.nth(2)

        cap_img = find_captcha_image(page)
        if cap_img is None:
            print(" [!] Gambar captcha tidak ditemukan, coba lagi...")
            continue

        img_bytes = cap_img.screenshot()
        raw, answer = solve_math_captcha(img_bytes)
        print(f"    OCR membaca: '{raw}' -> jawaban: {answer}")

        if answer is None:
            print("    Gagal parsing captcha, refresh & ulangi...")
            continue

        cap_input.first.fill(str(answer))
        page.locator("button:has-text('Masuk')").first.click()
        time.sleep(4)

        if not is_login_page(page):
            print(" [OK] Login berhasil!")
            return True

        print("    Login gagal (captcha/password salah), mengulang...")

    raise RuntimeError(f"Login {web_info['name']} gagal setelah {MAX_LOGIN_ATTEMPT} percobaan.")


def ensure_logged_in(page, web_info, creds):
    """Buka halaman target; kalau ternyata diarahkan ke login, login otomatis."""
    page.goto(web_info["url"], wait_until="domcontentloaded", timeout=60000)
    time.sleep(3)
    if is_login_page(page):
        login(page, web_info, creds)
        page.goto(web_info["url"], wait_until="networkidle", timeout=60000)
        time.sleep(2)


# =============================================================
# DOWNLOAD
# =============================================================
def process_and_download(page, web_info, creds):
    print(f"\n[+] Memproses {web_info['name']}...")

    ensure_logged_in(page, web_info, creds)

    # Klik Cari
    print(" -> Mengklik tombol 'Cari'...")
    page.locator("button.el-button--primary:has-text('Cari')").first.click()
    page.wait_for_load_state("networkidle")
    time.sleep(3)

    # Pagination 1000/halaman
    print(" -> Mengubah pagination ke 1000/halaman...")
    try:
        selector = page.locator(".el-pagination .el-select").first
        if selector.count() > 0:
            selector.click()
            time.sleep(1)
            opt = page.locator(".el-select-dropdown__item:visible:has-text('1000')").first
            if opt.count() > 0:
                opt.click()
                page.wait_for_load_state("networkidle")
                time.sleep(5)
    except Exception as e:
        print(f" -> Catatan pergantian page size: {e}")

    # Ekspor
    print(" -> Mengklik tombol 'Ekspor' & mendownload Excel...")
    btn_ekspor = page.locator("button.el-button--warning:has-text('Ekspor')").first
    with page.expect_download(timeout=120000) as dl_info:
        btn_ekspor.click()

    download = dl_info.value
    ts = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    file_name = f"{web_info['file_prefix']}_{ts}.xlsx"
    saved_path = os.path.join(TEMP_DOWNLOAD_DIR, file_name)
    download.save_as(saved_path)
    print(f" [OK] File tersimpan: {file_name}")
    return saved_path


# =============================================================
# SYNC FOLDER RAWDATA (hapus daily lama, tambah daily baru)
# =============================================================
def sync_rawdata_folder(downloaded_files):
    print("\n[+] Mengupdate folder rawdata...")

    # Hanya hapus file daily (awalan workStatistics_). MTD & lainnya AMAN.
    for fp in glob.glob(os.path.join(RAWDATA_DIR, "workStatistics*")):
        try:
            os.remove(fp)
            print(f" -> Menghapus file daily lama: {os.path.basename(fp)}")
        except Exception as e:
            print(f" -> Gagal menghapus {fp}: {e}")

    for src in downloaded_files:
        if src and os.path.exists(src):
            dst = os.path.join(RAWDATA_DIR, os.path.basename(src))
            shutil.copy2(src, dst)
            print(f" -> Menambahkan file baru: {os.path.basename(dst)}")


# =============================================================
# PUSH KE GITHUB
# =============================================================
def push_to_github():
    print("\n[+] Menjalankan Git Push ke GitHub...")
    try:
        repo = git.Repo(BASE_DIR)
        # -A: ikut menstage file yang DIHAPUS (daily lama) dan file baru
        repo.git.add("-A", "rawdata")
        if repo.git.status("--porcelain", "rawdata").strip():
            msg = f"Auto-update daily rawdata ({datetime.now():%Y-%m-%d %H:%M:%S})"
            repo.index.commit(msg)
            repo.remote(name="origin").push()
            print(" [OK] Berhasil push ke GitHub! File daily lama terhapus di GitHub.")
        else:
            print(" -> Tidak ada perubahan di rawdata.")
    except Exception as e:
        print(f" [!] Gagal push ke GitHub: {e}")


# =============================================================
# MAIN
# =============================================================
def main():
    print("=====================================================")
    print("   AUTOMATION RUNNER PINJAMID & LUMBUNGDANA")
    print("=====================================================")

    all_creds = load_credentials()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context(accept_downloads=True)
        pages = {cfg["key"]: context.new_page() for cfg in WEB_CONFIGS}

        cycle = 0
        while True:
            cycle += 1
            print(f"\n============ SIKLUS PENGAMBILAN DATA #{cycle} ============")
            try:
                downloaded = []
                for cfg in WEB_CONFIGS:
                    downloaded.append(
                        process_and_download(pages[cfg["key"]], cfg, all_creds[cfg["key"]])
                    )
                sync_rawdata_folder(downloaded)
                push_to_github()
            except Exception as err:
                print(f"[!] Terjadi kendala: {err}")

            print(f"\nMode Keep-Alive (refresh tiap {KEEPALIVE_SECONDS} detik)...")
            for i in range(KEEPALIVE_ROUNDS):
                time.sleep(KEEPALIVE_SECONDS)
                print(f"[{datetime.now():%H:%M:%S}] Refresh halaman ({i+1}/{KEEPALIVE_ROUNDS})...")
                for cfg in WEB_CONFIGS:
                    try:
                        pages[cfg["key"]].reload(wait_until="networkidle")
                    except Exception as e:
                        print(f"Gagal refresh {cfg['name']}: {e}")


if __name__ == "__main__":
    main()