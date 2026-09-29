import os
import re
import io
import json
import time
import glob
import shutil
from collections import Counter
from datetime import datetime

import ddddocr
import git
from PIL import Image, ImageOps
from playwright.sync_api import sync_playwright

# =============================================================
# KONFIGURASI
# =============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAWDATA_DIR = os.path.join(BASE_DIR, "rawdata")
TEMP_DOWNLOAD_DIR = os.path.join(BASE_DIR, "temp_downloads")
PROFILE_DIR = os.path.join(BASE_DIR, "browser_profile")   # menyimpan sesi login
DEBUG_DIR = os.path.join(BASE_DIR, "captcha_debug")       # simpan gambar captcha utk analisa
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")  # JANGAN di-push ke GitHub!
for d in (RAWDATA_DIR, TEMP_DOWNLOAD_DIR, DEBUG_DIR):
    os.makedirs(d, exist_ok=True)

HEADLESS = False              # True = browser tidak tampil (jangan dulu sebelum stabil)
MAX_LOGIN_ATTEMPT = 10        # percobaan login OTOMATIS (OCR)
MANUAL_FALLBACK = True        # kalau OCR gagal semua, minta Anda ketik jawaban captcha
KEEPALIVE_ROUNDS = 6          # 6 x 150 detik = download ulang tiap ~15 menit
KEEPALIVE_SECONDS = 150
RETRY_AFTER_ERROR_SECONDS = 60

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


def load_credentials():
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError("credentials.json tidak ditemukan di folder project.")
    with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


# =============================================================
# OCR CAPTCHA
# =============================================================
def build_ocr_engines():
    engines = []
    try:
        engines.append(("beta", ddddocr.DdddOcr(show_ad=False, beta=True)))
    except Exception as e:
        print(f"[i] Model beta tidak tersedia ({e}), pakai model standar saja.")
    engines.append(("std", ddddocr.DdddOcr(show_ad=False)))
    return engines


OCR_ENGINES = build_ocr_engines()


def _to_png_bytes(img):
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def make_variants(img_bytes):
    """Buat beberapa versi gambar (asli, diperbesar, hitam-putih) supaya OCR lebih akurat."""
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    big = img.resize((img.width * 3, img.height * 3), Image.LANCZOS)

    gray = ImageOps.autocontrast(ImageOps.grayscale(big))
    red = ImageOps.autocontrast(big.split()[0])   # teks biru -> gelap di channel merah
    bw_gray = gray.point(lambda p: 255 if p > 120 else 0)
    bw_red = red.point(lambda p: 255 if p > 120 else 0)

    return [
        ("asli", img_bytes),
        ("gray", _to_png_bytes(gray)),
        ("red", _to_png_bytes(red)),
        ("bw_gray", _to_png_bytes(bw_gray)),
        ("bw_red", _to_png_bytes(bw_red)),
    ]


OP_PLUS = set("+十")
OP_MINUS = set("-−—_一")
OP_MUL = set("*xX×kK#米")
OP_DIV = set("/÷")
AMBIGUOUS_ORDER = ["*", "+"]   # dipakai bergantian jika simbol operator tidak dikenali


def parse_expression(text, attempt):
    """Cari pola: angka - simbol - angka (misal '3*2', '2t6'). Hasilkan (jawaban, ekspresi)."""
    t = text.replace(" ", "")
    m = re.search(r"(\d)([^\d])(\d)", t)
    if not m:
        return None, None
    a, c, b = int(m.group(1)), m.group(2), int(m.group(3))

    if c in OP_PLUS:
        op = "+"
    elif c in OP_MINUS:
        op = "-"
    elif c in OP_MUL:
        op = "*"
    elif c in OP_DIV:
        op = "/"
    else:
        op = AMBIGUOUS_ORDER[attempt % 2]

    if op == "+":
        ans = a + b
    elif op == "-":
        ans = a - b
    elif op == "*":
        ans = a * b
    else:
        ans = a // b if b != 0 and a % b == 0 else None
    return ans, f"{a}{op}{b}"


def solve_captcha(img_bytes, attempt):
    """Jalankan semua kombinasi (model x variasi gambar), lalu voting jawaban terbanyak."""
    votes = Counter()
    raws = []
    for vname, vbytes in make_variants(img_bytes):
        for ename, engine in OCR_ENGINES:
            try:
                raw = engine.classification(vbytes)
            except Exception:
                continue
            raws.append(raw)
            ans, expr = parse_expression(raw, attempt)
            if ans is not None:
                votes[(ans, expr)] += 1
    if not votes:
        return None, None, raws
    (ans, expr), _ = votes.most_common(1)[0]
    return ans, expr, raws


def save_debug_image(img_bytes, tag):
    try:
        safe = re.sub(r"[^0-9A-Za-z]", "_", str(tag))
        name = f"{datetime.now():%H%M%S}_{safe}.png"
        with open(os.path.join(DEBUG_DIR, name), "wb") as f:
            f.write(img_bytes)
        files = sorted(glob.glob(os.path.join(DEBUG_DIR, "*.png")))
        for old in files[:-60]:      # simpan 60 terbaru saja
            os.remove(old)
    except Exception:
        pass


# =============================================================
# LOGIN
# =============================================================
def is_login_page(page):
    try:
        return page.locator("input[type='password']:visible").count() > 0
    except Exception:
        return False


def find_captcha_image(page):
    imgs = page.locator("img:visible")
    for i in range(imgs.count()):
        el = imgs.nth(i)
        box = el.bounding_box()
        if box and 60 <= box["width"] <= 250 and 20 <= box["height"] <= 80:
            return el
    return None


def fill_credentials(page, creds):
    inputs = page.locator("input:visible")
    user_input = inputs.nth(0)
    pass_input = page.locator("input[type='password']:visible").first
    user_input.fill("")
    user_input.fill(creds["username"])
    pass_input.fill("")
    pass_input.fill(creds["password"])
    cap_input = page.get_by_placeholder("Kode Verifikasi")
    if cap_input.count() == 0:
        cap_input = inputs.nth(2)
    return cap_input.first


def submit_login(page, cap_input, answer):
    cap_input.fill(str(answer))
    page.locator("button:has-text('Masuk')").first.click()
    time.sleep(4)
    return not is_login_page(page)


def login(page, web_info, creds):
    # ---- Tahap 1: otomatis dengan OCR ----
    for attempt in range(1, MAX_LOGIN_ATTEMPT + 1):
        print(f" -> Login {web_info['name']} (otomatis {attempt}/{MAX_LOGIN_ATTEMPT})...")
        page.goto(web_info["url"], wait_until="domcontentloaded", timeout=60000)
        time.sleep(3)
        if not is_login_page(page):
            print(" -> Sudah dalam kondisi login.")
            return True

        cap_input = fill_credentials(page, creds)
        cap_img = find_captcha_image(page)
        if cap_img is None:
            print(" [!] Gambar captcha tidak ditemukan.")
            continue

        img_bytes = cap_img.screenshot()
        answer, expr, raws = solve_captcha(img_bytes, attempt)
        print(f"    OCR: {sorted(set(raws))[:6]} -> tebakan: {expr} = {answer}")
        save_debug_image(img_bytes, f"a{attempt}_{expr}_{answer}")

        if answer is None:
            continue
        if submit_login(page, cap_input, answer):
            print(" [OK] Login berhasil!")
            return True
        print("    Jawaban captcha salah, ulangi dengan captcha baru...")

    # ---- Tahap 2: manual (Anda ketik jawaban dari captcha di browser) ----
    if MANUAL_FALLBACK and not HEADLESS:
        print(f"\n [!] OCR belum berhasil. Beralih ke mode MANUAL untuk {web_info['name']}.")
        while True:
            page.goto(web_info["url"], wait_until="domcontentloaded", timeout=60000)
            time.sleep(3)
            if not is_login_page(page):
                return True
            page.bring_to_front()
            cap_input = fill_credentials(page, creds)
            ans = input("    Lihat captcha di browser, ketik HASIL hitungnya lalu ENTER "
                        "(kosong = ganti captcha, q = batal): ").strip()
            if ans.lower() == "q":
                break
            if not ans:
                continue
            if submit_login(page, cap_input, ans):
                print(" [OK] Login berhasil!")
                return True
            print("    Salah, coba lagi...")

    raise RuntimeError(f"Login {web_info['name']} gagal.")


def ensure_logged_in(page, web_info, creds):
    page.goto(web_info["url"], wait_until="domcontentloaded", timeout=60000)
    time.sleep(3)
    if is_login_page(page):
        login(page, web_info, creds)
        page.goto(web_info["url"], wait_until="networkidle", timeout=60000)
        time.sleep(2)
        if is_login_page(page):
            raise RuntimeError(f"Masih di halaman login setelah login {web_info['name']}.")


# =============================================================
# DOWNLOAD
# =============================================================
def process_and_download(page, web_info, creds):
    print(f"\n[+] Memproses {web_info['name']}...")
    ensure_logged_in(page, web_info, creds)

    print(" -> Mengklik tombol 'Cari'...")
    page.locator("button.el-button--primary:has-text('Cari')").first.click()
    page.wait_for_load_state("networkidle")
    time.sleep(3)

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
# SYNC RAWDATA
# =============================================================
def sync_rawdata_folder(downloaded_files):
    print("\n[+] Mengupdate folder rawdata...")
    # Hanya file daily (awalan workStatistics) yang dihapus. MTD & data lain AMAN.
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
# PUSH GITHUB
# =============================================================
def push_to_github():
    print("\n[+] Menjalankan Git Push ke GitHub...")
    try:
        repo = git.Repo(BASE_DIR)
        repo.git.add("-A", "rawdata")   # -A: file yang dihapus ikut tercatat
        if repo.git.status("--porcelain", "rawdata").strip():
            msg = f"Auto-update daily rawdata ({datetime.now():%Y-%m-%d %H:%M:%S})"
            repo.git.commit("-m", msg)
            repo.git.push("origin", "HEAD")
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
        # Profil browser disimpan -> login tersimpan, tidak perlu login ulang tiap dijalankan
        context = p.chromium.launch_persistent_context(
            PROFILE_DIR, headless=HEADLESS, accept_downloads=True
        )
        existing = context.pages
        pages = {}
        for i, cfg in enumerate(WEB_CONFIGS):
            pages[cfg["key"]] = existing[0] if (i == 0 and existing) else context.new_page()
            pages[cfg["key"]].set_default_timeout(30000)

        cycle = 0
        while True:
            cycle += 1
            print(f"\n============ SIKLUS PENGAMBILAN DATA #{cycle} ============")
            ok = False
            try:
                downloaded = [
                    process_and_download(pages[c["key"]], c, all_creds[c["key"]])
                    for c in WEB_CONFIGS
                ]
                sync_rawdata_folder(downloaded)
                push_to_github()
                ok = True
            except Exception as err:
                print(f"[!] Terjadi kendala: {err}")

            if not ok:
                print(f"Mencoba lagi dalam {RETRY_AFTER_ERROR_SECONDS} detik...")
                time.sleep(RETRY_AFTER_ERROR_SECONDS)
                continue

            print(f"\nMode Keep-Alive (refresh tiap {KEEPALIVE_SECONDS} detik)...")
            for i in range(KEEPALIVE_ROUNDS):
                time.sleep(KEEPALIVE_SECONDS)
                print(f"[{datetime.now():%H:%M:%S}] Refresh halaman ({i+1}/{KEEPALIVE_ROUNDS})...")
                for c in WEB_CONFIGS:
                    try:
                        pages[c["key"]].reload(wait_until="networkidle")
                    except Exception as e:
                        print(f"Gagal refresh {c['name']}: {e}")


if __name__ == "__main__":
    main()
