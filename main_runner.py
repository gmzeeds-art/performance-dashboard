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
import requests
from PIL import Image, ImageOps
from playwright.sync_api import sync_playwright

# =============================================================
# KONFIGURASI GLOBAL
# =============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAWDATA_DIR = os.path.join(BASE_DIR, "rawdata")
TEMP_DOWNLOAD_DIR = os.path.join(BASE_DIR, "temp_downloads")
PROFILE_DIR = os.path.join(BASE_DIR, "browser_profile")
DEBUG_DIR = os.path.join(BASE_DIR, "captcha_debug")
SCREENSHOT_DIR = os.path.join(BASE_DIR, "screenshots")
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")

for d in (RAWDATA_DIR, TEMP_DOWNLOAD_DIR, DEBUG_DIR, SCREENSHOT_DIR):
    os.makedirs(d, exist_ok=True)

LARK_WEBHOOK_URL = "https://open.larksuite.com/open-apis/bot/v2/hook/2be1db3d-00f9-477e-9cbf-d251ea23fbe3"
STREAMLIT_URL = "https://performa-collection-mpt05.streamlit.app/"
GITHUB_RAW_BASE = "https://raw.githubusercontent.com/gmzeeds-art/performance-dashboard/main/screenshots"

INTERVAL_MINUTES = 40
HEADLESS = False  # Set False agar captcha dan proses unduh terpantau jelas

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

# =============================================================
# MODUL OCR CAPTCHA & LOGIN
# =============================================================
def load_credentials():
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError("credentials.json tidak ditemukan di folder project.")
    with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def build_ocr_engines():
    engines = []
    try:
        engines.append(("beta", ddddocr.DdddOcr(show_ad=False, beta=True)))
    except Exception:
        pass
    engines.append(("std", ddddocr.DdddOcr(show_ad=False)))
    return engines

OCR_ENGINES = build_ocr_engines()

def _to_png_bytes(img):
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def make_variants(img_bytes):
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    big = img.resize((img.width * 3, img.height * 3), Image.LANCZOS)
    gray = ImageOps.autocontrast(ImageOps.grayscale(big))
    red = ImageOps.autocontrast(big.split()[0])
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
AMBIGUOUS_ORDER = ["*", "+"]

def parse_expression(text, attempt):
    t = text.replace(" ", "")
    m = re.search(r"(\d)([^\d])(\d)", t)
    if not m:
        return None, None
    a, c, b = int(m.group(1)), m.group(2), int(m.group(3))
    if c in OP_PLUS: op = "+"
    elif c in OP_MINUS: op = "-"
    elif c in OP_MUL: op = "*"
    elif c in OP_DIV: op = "/"
    else: op = AMBIGUOUS_ORDER[attempt % 2]

    if op == "+": ans = a + b
    elif op == "-": ans = a - b
    elif op == "*": ans = a * b
    else: ans = a // b if b != 0 and a % b == 0 else None
    return ans, f"{a}{op}{b}"

def solve_captcha(img_bytes, attempt):
    votes = Counter()
    raws = []
    for _, vbytes in make_variants(img_bytes):
        for _, engine in OCR_ENGINES:
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
    for attempt in range(1, 11):
        print(f" -> Login {web_info['name']} (percobaan {attempt}/10)...")
        page.goto(web_info["url"], wait_until="domcontentloaded", timeout=60000)
        time.sleep(3)
        if not is_login_page(page):
            return True

        cap_input = fill_credentials(page, creds)
        cap_img = find_captcha_image(page)
        if cap_img is None:
            continue

        img_bytes = cap_img.screenshot()
        answer, expr, _ = solve_captcha(img_bytes, attempt)
        if answer is None:
            continue
        if submit_login(page, cap_input, answer):
            print(" [OK] Login berhasil!")
            return True

    if not HEADLESS:
        print(f" [!] Masukkan jawaban captcha manual untuk {web_info['name']}:")
        cap_input = fill_credentials(page, creds)
        ans = input(" -> Masukkan hasil hitungan captcha lalu ENTER: ").strip()
        if submit_login(page, cap_input, ans):
            return True

    raise RuntimeError(f"Gagal login ke {web_info['name']}.")

def ensure_logged_in(page, web_info, creds):
    page.goto(web_info["url"], wait_until="domcontentloaded", timeout=60000)
    time.sleep(3)
    if is_login_page(page):
        login(page, web_info, creds)
        time.sleep(2)
        print(f" -> Mengarahkan ke halaman data: {web_info['url']}...")
        page.goto(web_info["url"], wait_until="networkidle", timeout=60000)
        time.sleep(3)

    if "indexPage" in page.url or "WorkStatistics" not in page.url:
        print(f" -> Navigasi ulang ke modul kerja: {web_info['url']}...")
        page.goto(web_info["url"], wait_until="networkidle", timeout=60000)
        time.sleep(3)

# =============================================================
# MODUL DOWNLOAD DATA EXCEL
# =============================================================
def process_and_download(page, web_info, creds):
    print(f"\n[+] Memproses download: {web_info['name']}...")
    ensure_logged_in(page, web_info, creds)

    if "WorkStatistics" not in page.url:
        page.goto(web_info["url"], wait_until="networkidle", timeout=60000)
        time.sleep(3)

    print(" -> Menunggu form dan menekan tombol 'Cari'...")
    btn_cari = page.locator("button:has-text('Cari'), button.el-button--primary:has-text('Cari')").first
    btn_cari.wait_for(state="visible", timeout=30000)
    btn_cari.click()
    page.wait_for_load_state("networkidle")
    time.sleep(3)

    try:
        page.wait_for_selector(".el-table__body-wrapper tbody tr", timeout=30000)
        print(" -> Data tabel awal berhasil tampil di layar.")
    except Exception:
        print(" [!] Menunggu respon render data awal...")
        time.sleep(4)

    print(" -> Mengubah pagination ke 1000 baris/halaman...")
    for attempt in range(1, 4):
        try:
            select_trigger = page.locator(".el-pagination .el-select, .el-pagination__sizes").first
            select_trigger.click(force=True)
            time.sleep(1.5)

            opt_1000 = page.locator(".el-select-dropdown:visible .el-select-dropdown__item:has-text('1000'), li.el-select-dropdown__item:has-text('1000'):visible").first
            if opt_1000.is_visible(timeout=3000):
                opt_1000.click(force=True)
                print(f"    [OK] Opsi 1000 baris terklik (percobaan {attempt}).")
                page.wait_for_load_state("networkidle")
                time.sleep(6)
                break
            else:
                page.evaluate("""() => {
                    const items = Array.from(document.querySelectorAll('.el-select-dropdown__item, li, span'));
                    for (let it of items) {
                        if (it.textContent.trim().includes('1000')) {
                            it.click();
                            return true;
                        }
                    }
                    return false;
                }""")
                time.sleep(5)
                break
        except Exception as e:
            print(f"    [!] Percobaan pagination {attempt} terkendala: {e}")
            time.sleep(2)

    row_count = 0
    try:
        row_count = page.locator(".el-table__body-wrapper tbody tr").count()
        print(f" -> Terdeteksi {row_count} baris data pada tabel sebelum ekspor.")
    except Exception:
        pass

    if row_count <= 10:
        print(" [!] Data terdeteksi <= 10 baris, memicu 'Cari' ulang...")
        btn_cari.click()
        page.wait_for_load_state("networkidle")
        time.sleep(5)
        row_count = page.locator(".el-table__body-wrapper tbody tr").count()
        print(f" -> Baris data setelah 'Cari' ulang: {row_count} baris.")

    # Tutup kemungkinan overlay dropdown sebelum trigger ekspor
    try:
        page.keyboard.press("Escape")
        time.sleep(1)
    except Exception:
        pass

    print(" -> Memicu unduhan ekspor Excel langsung via JavaScript DOM...")
    with page.expect_download(timeout=120000) as dl_info:
        page.evaluate("""() => {
            const buttons = Array.from(document.querySelectorAll('button'));
            const exportBtn = buttons.find(b => b.textContent.trim().includes('Ekspor'));
            if (exportBtn) {
                exportBtn.click();
            } else {
                const warnBtn = document.querySelector('button.el-button--warning');
                if (warnBtn) warnBtn.click();
            }
        }""")

    download = dl_info.value
    ts = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    file_name = f"{web_info['file_prefix']}_{ts}.xlsx"
    saved_path = os.path.join(TEMP_DOWNLOAD_DIR, file_name)
    download.save_as(saved_path)

    file_size_kb = os.path.getsize(saved_path) / 1024
    print(f" [OK] Berhasil disimpan: {file_name} (Ukuran: {file_size_kb:.1f} KB)")

    if file_size_kb < 50:
        print(" [!] File masih terlalu kecil (< 50 KB). Mengulangi klik Ekspor...")
        time.sleep(3)
        with page.expect_download(timeout=120000) as dl_retry:
            page.evaluate("""() => {
                const buttons = Array.from(document.querySelectorAll('button'));
                const exportBtn = buttons.find(b => b.textContent.trim().includes('Ekspor'));
                if (exportBtn) exportBtn.click();
            }""")
        download_retry = dl_retry.value
        download_retry.save_as(saved_path)
        file_size_kb = os.path.getsize(saved_path) / 1024
        print(f" [OK] Hasil download ulang: {file_size_kb:.1f} KB")

    return saved_path

def sync_rawdata_folder(downloaded_files):
    print("\n[+] Memperbarui folder rawdata...")
    for fp in glob.glob(os.path.join(RAWDATA_DIR, "workStatistics*")):
        try:
            os.remove(fp)
        except Exception:
            pass

    for src in downloaded_files:
        if src and os.path.exists(src):
            dst = os.path.join(RAWDATA_DIR, os.path.basename(src))
            shutil.copy2(src, dst)
            print(f" -> File siap: {os.path.basename(dst)}")

# =============================================================
# MODUL PUSH GITHUB
# =============================================================
def git_sync_and_push(target_folder, commit_msg):
    print(f"\n[+] Sinkronisasi & Push Git ({target_folder})...")
    try:
        repo = git.Repo(BASE_DIR)
        
        # 1. Bersihkan sisa direktori rebase jika ada
        for bad_dir_name in ("rebase-merge", "rebase-apply"):
            bad_dir = os.path.join(BASE_DIR, ".git", bad_dir_name)
            if os.path.exists(bad_dir):
                try:
                    shutil.rmtree(bad_dir)
                    print(f" -> Membersihkan folder macet: {bad_dir_name}")
                except Exception:
                    pass

        # 2. Stage file target (misal: rawdata atau screenshots)
        repo.git.add("-A", target_folder)
        
        # 3. Buat commit lokal jika ada perubahan file
        if repo.git.status("--porcelain", target_folder).strip():
            repo.git.commit("-m", commit_msg)
            print(f" -> Commit lokal {target_folder} dibuat.")
        else:
            print(f" -> Tidak ada perubahan baru di folder {target_folder}.")

        # 4. Push dengan parameter --force agar tidak pernah tertahan non-fast-forward
        repo.git.push("origin", "HEAD:main", "--force")
        print(f" [OK] {target_folder} berhasil di-push ke GitHub!")
        time.sleep(3)
        return True

    except Exception as e:
        print(f" [!] Gagal push Git ({target_folder}): {e}")
        return False

# =============================================================
# MODUL PENGIRIMAN GAMBAR LANGSUNG KE LARK
# =============================================================
def upload_image_to_lark(image_path):
    """Mengunggah biner file gambar ke OpenAPI Lark untuk mendapatkan image_key."""
    upload_url = "https://open.larksuite.com/open-apis/im/v1/images"
    try:
        with open(image_path, "rb") as f:
            files = {
                "image": (os.path.basename(image_path), f, "image/png"),
                "image_type": (None, "message")
            }
            resp = requests.post(upload_url, files=files, timeout=30)
            data = resp.json()
            if data.get("code") == 0:
                return data.get("data", {}).get("image_key")
    except Exception:
        pass
    return None

def send_lark_direct_image(title, file_name, file_path):
    """Mengirim gambar fisik langsung ke obrolan grup Lark."""
    img_key = upload_image_to_lark(file_path)
    if img_key:
        payload = {
            "msg_type": "image",
            "content": {"image_key": img_key}
        }
        r = requests.post(LARK_WEBHOOK_URL, json=payload, timeout=30)
        print(f" [Lark] Gambar fisik terkirim: {title} ({r.status_code})")
        return

    raw_img_url = f"{GITHUB_RAW_BASE}/{file_name}?raw=true&v={int(time.time())}"
    card_payload = {
        "msg_type": "interactive",
        "card": {
            "header": {
                "title": {"tag": "plain_text", "content": f"📊 {title}"},
                "template": "blue"
            },
            "elements": [
                {
                    "tag": "div",
                    "text": {
                        "tag": "lark_md",
                        "content": f"**Laporan:** {title}\n\n![]({raw_img_url})"
                    }
                },
                {
                    "tag": "action",
                    "actions": [
                        {
                            "tag": "button",
                            "text": {"tag": "plain_text", "content": "🔍 Buka Gambar Full-HD"},
                            "type": "primary",
                            "url": raw_img_url
                        }
                    ]
                }
            ]
        }
    }
    r = requests.post(LARK_WEBHOOK_URL, json=card_payload, timeout=30)
    print(f" [Lark] Kartu laporan terkirim: {title} ({r.status_code})")

# =============================================================
# MODUL CAPTURE REPORT (MENGGUNAKAN EXISTING CONTEXT)
# =============================================================
def get_streamlit_frame(page):
    start = time.time()
    while time.time() - start < 45:
        for f in page.frames:
            try:
                wake = f.locator("button:has-text('Yes, get this app back up!')").first
                if wake.is_visible(timeout=1000):
                    print(" [i] Membangunkan Streamlit Cloud...")
                    wake.click()
                    time.sleep(15)
            except Exception:
                pass

        for f in page.frames:
            try:
                if f.locator("[data-testid='stAppViewContainer'], [role='tablist'], .stApp").count() > 0:
                    return f
            except Exception:
                pass
        time.sleep(1.5)
    return page

def wait_streamlit_idle(frame, timeout_sec=60):
    start = time.time()
    while time.time() - start < timeout_sec:
        try:
            status = frame.locator("[data-testid='stStatusWidget'], button:has-text('Stop')").first
            if not status.is_visible():
                break
        except Exception:
            break
        time.sleep(1)
    time.sleep(3)

def click_tab_safely(frame, tab_name):
    print(f" -> Memilih tab: '{tab_name}'...")
    try:
        tab_loc = frame.locator(f"[role='tab']:has-text('{tab_name}'):visible, button:has-text('{tab_name}'):visible").last
        if tab_loc.is_visible(timeout=5000):
            tab_loc.click(force=True)
            wait_streamlit_idle(frame)
            time.sleep(2)
            return True
    except Exception:
        pass

    frame.evaluate("""(txt) => {
        const tabs = Array.from(document.querySelectorAll('[role="tab"], button, label'));
        const visibleTabs = tabs.filter(t => t.textContent.trim() === txt && t.offsetParent !== null);
        if (visibleTabs.length > 0) {
            visibleTabs[visibleTabs.length - 1].click();
            return true;
        }
        return false;
    }""", tab_name)

    wait_streamlit_idle(frame)
    time.sleep(2)
    return True

def expand_and_size_canvas(frame, page):
    """Membongkar container scrollbar dan memanjangkan viewport browser."""
    try:
        frame.evaluate("""() => {
            document.querySelectorAll('div, section, main').forEach(el => {
                const s = window.getComputedStyle(el);
                if (s.overflowY === 'auto' || s.overflowY === 'scroll' || el.style.maxHeight) {
                    el.style.setProperty('overflow', 'visible', 'important');
                    el.style.setProperty('overflow-y', 'visible', 'important');
                    el.style.setProperty('max-height', 'none', 'important');
                    el.style.setProperty('height', 'auto', 'important');
                }
            });
            const h = document.querySelector('header');
            if (h) h.style.display = 'none';
            const f = document.querySelector('footer');
            if (f) f.style.display = 'none';
        }""")
        time.sleep(2)

        doc_height = page.evaluate("() => Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)")
        page.set_viewport_size({"width": 2400, "height": max(int(doc_height) + 400, 2200)})
        time.sleep(1)
    except Exception:
        pass

def capture_and_report_all(browser_context):
    print("\n==================================================")
    print("        PENGAMBILAN 6 SCREENSHOT LAPORAN          ")
    print("==================================================")

    results = []
    page = browser_context.new_page()
    page.set_viewport_size({"width": 2400, "height": 1800})

    try:
        print(f"[+] Membuka dashboard Streamlit: {STREAMLIT_URL}")
        page.goto(STREAMLIT_URL, wait_until="domcontentloaded", timeout=120000)

        st_frame = get_streamlit_frame(page)
        wait_streamlit_idle(st_frame)

        # 1. Performance Team -> Semua Bucket
        print("\n[+] 1/6 Capture: Performance Team -> Semua Bucket...")
        click_tab_safely(st_frame, "Performance Team")
        click_tab_safely(st_frame, "Semua Bucket")
        expand_and_size_canvas(st_frame, page)

        cap1 = "01_Performance_Team_Semua_Bucket.png"
        p1 = os.path.join(SCREENSHOT_DIR, cap1)
        page.screenshot(path=p1, full_page=True)
        results.append(("Performance Team Semua Bucket", cap1, p1))
        print(f" [OK] Tersimpan: {cap1}")

        # 2. Performance Perusahaan -> Semua Bucket
        print("\n[+] 2/6 Capture: Performance Perusahaan -> Semua Bucket...")
        click_tab_safely(st_frame, "Performance Perusahaan")
        time.sleep(1.5)
        click_tab_safely(st_frame, "Semua Bucket")
        expand_and_size_canvas(st_frame, page)

        cap2 = "02_Performance_Perusahaan_Semua_Bucket.png"
        p2 = os.path.join(SCREENSHOT_DIR, cap2)
        page.screenshot(path=p2, full_page=True)
        results.append(("Performance Perusahaan Semua Bucket", cap2, p2))
        print(f" [OK] Tersimpan: {cap2}")

        # 3 - 6. Agent Performance (S1 - LD, S1 - SK, S0 - LD, D0 - SK)
        print("\n[+] Beralih ke menu: Agent Performance...")
        click_tab_safely(st_frame, "Agent Performance")
        time.sleep(2)

        agent_subtabs = [
            ("S1 LD", "S1 - LD"),
            ("S1 SK", "S1 - SK"),
            ("S0 LD", "S0 - LD"),
            ("D0 SK", "D0 - SK")
        ]

        for idx, (label_clean, subtab_name) in enumerate(agent_subtabs, start=3):
            print(f"\n[+] {idx}/6 Capture: Agent Performance -> {subtab_name}...")
            click_tab_safely(st_frame, subtab_name)
            expand_and_size_canvas(st_frame, page)

            cap_agent = f"{idx:02d}_Performance_Agent_{label_clean.replace(' ', '_')}.png"
            p_agent = os.path.join(SCREENSHOT_DIR, cap_agent)
            page.screenshot(path=p_agent, full_page=True)
            results.append((f"Performance Agent {label_clean}", cap_agent, p_agent))
            print(f" [OK] Tersimpan: {cap_agent}")

    finally:
        page.close()

    git_sync_and_push("screenshots", "Update performance report screenshots")

    print("\n[+] Mengirim 6 gambar laporan ke grup Lark...")
    for title, fname, fpath in results:
        send_lark_direct_image(title, fname, fpath)
        time.sleep(2)
    print(" [OK] Seluruh laporan berhasil dikirim ke grup Lark.")

# =============================================================
# RUNNER UTAMA DENGAN INTERVAL 40 MENIT
# =============================================================
def main():
    print("=====================================================")
    print("   AUTOMATION RUNNER TERPADU (DOWNLOAD + REPORT)     ")
    print(f"   Frekuensi Eksekusi: Setiap {INTERVAL_MINUTES} Menit")
    print("=====================================================")

    all_creds = load_credentials()

    with sync_playwright() as p:
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
            print(f"\n=======================================================")
            print(f"       MEMULAI SIKLUS OTOMASI #{cycle} ({datetime.now():%Y-%m-%d %H:%M:%S})")
            print(f"=======================================================")

            download_success = False
            try:
                downloaded = [
                    process_and_download(pages[c["key"]], c, all_creds[c["key"]])
                    for c in WEB_CONFIGS
                ]
                sync_rawdata_folder(downloaded)
                git_sync_and_push("rawdata", f"Auto-update daily rawdata ({datetime.now():%Y-%m-%d %H:%M:%S})")
                download_success = True
            except Exception as err:
                print(f"[!] Terjadi kendala saat download: {err}")

            if not download_success:
                print(" [!] Unduhan gagal. Mencoba ulang dalam 30 detik...")
                time.sleep(30)
                continue

            print("\n[+] Menunggu 45 detik agar Streamlit Cloud memuat rawdata baru...")
            time.sleep(45)

            try:
                capture_and_report_all(context)
            except Exception as cap_err:
                print(f"[!] Terjadi kendala saat capture dashboard: {cap_err}")

            sleep_seconds = INTERVAL_MINUTES * 60
            next_run = datetime.fromtimestamp(time.time() + sleep_seconds)
            print(f"\n[Zzz] Siklus #{cycle} selesai sempurna. Menunggu {INTERVAL_MINUTES} menit (Siklus berikutnya: {next_run:%H:%M:%S})...")
            
            time.sleep(sleep_seconds)
            
            for c in WEB_CONFIGS:
                try:
                    pages[c["key"]].reload(wait_until="domcontentloaded")
                except Exception:
                    pass


if __name__ == "__main__":
    main()