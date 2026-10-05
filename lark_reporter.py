import os
import time
import requests
import git
from playwright.sync_api import sync_playwright

LARK_WEBHOOK_URL = "https://open.larksuite.com/open-apis/bot/v2/hook/2be1db3d-00f9-477e-9cbf-d251ea23fbe3"
STREAMLIT_URL = "https://performa-collection-mpt05.streamlit.app/"
GITHUB_RAW_BASE = "https://raw.githubusercontent.com/gmzeeds-art/performance-dashboard/main/screenshots"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCREENSHOT_DIR = os.path.join(BASE_DIR, "screenshots")
os.makedirs(SCREENSHOT_DIR, exist_ok=True)


def push_screenshots_to_github():
    """Mengunggah screenshot ke GitHub secara aman dengan resolusi konflik rebase otomatis."""
    print("\n[+] Mengunggah screenshot ke GitHub...")
    try:
        repo = git.Repo(BASE_DIR)
        
        # Batalkan rebase yang macet di masa lalu jika ada
        try:
            repo.git.rebase("--abort")
        except Exception:
            pass

        # Tarik update branch utama
        try:
            repo.git.pull("origin", "main", "--rebase", "--autostash")
        except Exception as pull_e:
            print(f" -> Info pull: {pull_e}")

        repo.git.add("-A", "screenshots")
        if repo.git.status("--porcelain", "screenshots").strip():
            repo.git.commit("-m", "Update auto performance screenshots")
            repo.git.push("origin", "HEAD:main")
            print(" [OK] Screenshot berhasil diunggah ke GitHub!")
            time.sleep(3)
            return True
        else:
            print(" -> File screenshot sudah yang paling baru di GitHub.")
            return True
    except Exception as e:
        print(f" [!] Gagal push gambar ke GitHub: {e}")
        return False


def send_lark_image_card(title, file_name):
    """Mengirim gambar ke grup Lark via format Markdown Image (Pasti muncul)."""
    img_url = f"{GITHUB_RAW_BASE}/{file_name}?raw=true&v={int(time.time())}"
    payload = {
        "msg_type": "interactive",
        "card": {
            "header": {
                "title": {"tag": "plain_text", "content": title},
                "template": "blue"
            },
            "elements": [
                {
                    "tag": "div",
                    "text": {
                        "tag": "lark_md",
                        "content": f"**Laporan Terbaru:** {title}\n\n![]({img_url})"
                    }
                }
            ]
        }
    }
    try:
        resp = requests.post(LARK_WEBHOOK_URL, json=payload, timeout=30)
        print(f" [Lark] Kirim {title}: status {resp.status_code}, respon: {resp.text}")
    except Exception as e:
        print(f" [Lark] Error kirim {title}: {e}")


def wait_streamlit_idle(frame, timeout_sec=60):
    """Menunggu animasi kalkulasi data Streamlit selesai."""
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


def get_streamlit_frame(page):
    """Mendeteksi frame aktif tempat konten dashboard Streamlit berjalan."""
    start = time.time()
    while time.time() - start < 45:
        # Cek tombol wake up jika aplikasi tertidur
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


def click_tab_safely(frame, tab_name):
    """Mengeklik tab atau sub-tab Streamlit yang sedang terlihat aktif di layar."""
    print(f" -> Berpindah ke: '{tab_name}'...")
    try:
        # Targetkan tab yang terlihat di layer terdepan
        tab_loc = frame.locator(f"[role='tab']:has-text('{tab_name}'):visible, button:has-text('{tab_name}'):visible").last
        if tab_loc.is_visible(timeout=5000):
            tab_loc.click(force=True)
            wait_streamlit_idle(frame)
            time.sleep(2)
            return True
    except Exception:
        pass

    # Fallback klik lewat JS pada elemen yang visible
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


def expand_full_tables(frame, page):
    """Membongkar container scroll dan memanjangkan viewport browser agar seluruh baris tertangkap."""
    try:
        # Buka paksa inline-style pembungkus tabel
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

        # Panjangkan viewport Playwright agar muat dari baris 1 sampai baris paling bawah
        doc_height = page.evaluate("() => Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)")
        page.set_viewport_size({"width": 2400, "height": max(int(doc_height) + 400, 2200)})
        time.sleep(1)
    except Exception as e:
        print(f" [!] Catatan expand table: {e}")


def capture_all_reports():
    print("==================================================")
    print("     MEMULAI CAPTURE DASHBOARD SECARA LENGKAP     ")
    print("==================================================")

    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            viewport={"width": 2400, "height": 1800},
            device_scale_factor=2
        )
        page = context.new_page()

        print(f"[+] Membuka dashboard: {STREAMLIT_URL}")
        page.goto(STREAMLIT_URL, wait_until="domcontentloaded", timeout=120000)

        st_frame = get_streamlit_frame(page)
        wait_streamlit_idle(st_frame)

        # -------------------------------------------------------------
        # 1. Performance Team -> Semua Bucket
        # -------------------------------------------------------------
        print("\n[+] 1/6 Capture: Performance Team -> Semua Bucket...")
        click_tab_safely(st_frame, "Performance Team")
        click_tab_safely(st_frame, "Semua Bucket")
        expand_full_tables(st_frame, page)

        cap1 = "01_Performance_Team_Semua_Bucket.png"
        p1 = os.path.join(SCREENSHOT_DIR, cap1)
        page.screenshot(path=p1, full_page=True)
        results.append(("Performance Team Semua Bucket", cap1))
        print(f" [OK] Tersimpan: {cap1}")

        # -------------------------------------------------------------
        # 2. Performance Perusahaan -> Semua Bucket
        # -------------------------------------------------------------
        print("\n[+] 2/6 Capture: Performance Perusahaan -> Semua Bucket...")
        click_tab_safely(st_frame, "Performance Perusahaan")
        time.sleep(1.5)
        click_tab_safely(st_frame, "Semua Bucket")
        expand_full_tables(st_frame, page)

        cap2 = "02_Performance_Perusahaan_Semua_Bucket.png"
        p2 = os.path.join(SCREENSHOT_DIR, cap2)
        page.screenshot(path=p2, full_page=True)
        results.append(("Performance Perusahaan Semua Bucket", cap2))
        print(f" [OK] Tersimpan: {cap2}")

        # -------------------------------------------------------------
        # 3 - 6. Agent Performance (S1 - LD, S1 - SK, S0 - LD, D0 - SK)
        # -------------------------------------------------------------
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
            expand_full_tables(st_frame, page)

            cap_agent = f"{idx:02d}_Performance_Agent_{label_clean.replace(' ', '_')}.png"
            p_agent = os.path.join(SCREENSHOT_DIR, cap_agent)
            page.screenshot(path=p_agent, full_page=True)
            results.append((f"Performance Agent {label_clean}", cap_agent))
            print(f" [OK] Tersimpan: {cap_agent}")

        browser.close()

    # Unggah ke GitHub & Kirim ke Lark
    pushed = push_screenshots_to_github()
    if pushed:
        print("\n[+] Mengirim 6 kartu gambar laporan ke Lark...")
        for title, fname in results:
            send_lark_image_card(title, fname)
            time.sleep(2)
        print("\n[SELESAI] Seluruh 6 laporan gambar berhasil terkirim ke Lark.")
    else:
        print("\n[!] Pengiriman Lark ditunda karena push GitHub gagal.")


if __name__ == "__main__":
    capture_all_reports()