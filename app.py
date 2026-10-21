import os
import glob
import re
import subprocess
import io
import json
from datetime import datetime
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# KONFIGURASI HALAMAN
# ---------------------------------------------------------
st.set_page_config(
    page_title="Performance Collection Automation",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------
# STYLING TINGKAT EKSEKUTIF (CORPORATE ENTERPRISE THEME)
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    [data-testid="stSidebar"] { display: none; }
    .main .block-container {
        padding-top: 0.8rem;
        padding-bottom: 2rem;
        max-width: 98%;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        margin-bottom: 0.8rem;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 2px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 15px;
        font-weight: 600;
        font-size: 0.83rem;
        color: #64748b;
        border-radius: 6px 6px 0 0;
        transition: all 0.2s ease;
    }
    .stTabs [aria-selected="true"] {
        color: #0f172a !important;
        background-color: #f1f5f9;
        border-bottom: 3px solid #2563eb !important;
    }
    
    .clean-group-title {
        font-size: 1.05rem;
        font-weight: 800;
        color: #0f172a;
        padding: 10px 0 6px 0;
        border-bottom: 2px solid #cbd5e1;
        margin-bottom: 12px;
        letter-spacing: -0.3px;
    }
    
    .exec-card-container {
        display: flex;
        gap: 12px;
        margin-bottom: 14px;
        flex-wrap: wrap;
    }
    .exec-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 10px 16px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
        min-width: 170px;
    }
    .exec-card-title {
        font-size: 0.73rem;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .exec-card-value {
        font-size: 1.25rem;
        font-weight: 800;
        color: #0f172a;
        margin-top: 2px;
    }

    .section-banner-title {
        font-size: 1.12rem;
        font-weight: 800;
        letter-spacing: 0.5px;
        color: #0f172a;
        display: flex;
        align-items: center;
        gap: 8px;
        margin-top: 6px;
        margin-bottom: 10px;
    }
    .section-banner-title::before {
        content: "";
        display: inline-block;
        width: 5px;
        height: 18px;
        background: #2563eb;
        border-radius: 3px;
    }
    .section-banner-title.tl::before {
        background: #059669;
    }
    .section-banner-title.daily::before {
        background: #d97706;
    }
    
    .month-header-clean {
        background: #0f172a;
        border: 1px solid #0f172a;
        padding: 6px 10px;
        border-radius: 6px 6px 0 0;
        text-align: center;
        font-weight: 700;
        font-size: 0.84rem;
        color: #f8fafc;
        letter-spacing: 0.5px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# KONFIGURASI GLOBAL & DEFINISI BUCKET
# ---------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAWDATA_DIR = os.path.join(BASE_DIR, "rawdata")
MASTER_FILE = os.path.join(BASE_DIR, "data_karyawan.xlsx")
TRAINING_FILE = os.path.join(BASE_DIR, "training_schedule.json")
os.makedirs(RAWDATA_DIR, exist_ok=True)

TARGET_BUCKETS = [
    ("S1 - SK", "S1", "SK"),
    ("D0 - SK", "D0", "SK"),
    ("S1 - LD", "S1", "LD"),
    ("S1 - LD-BA", "S1", "LD-BA"),
    ("S0 - LD", "S0", "LD"),
    ("C - LD-C", "C", "LD-C"),
    ("Semua Bucket", None, None),
]

RANK_ORDER = {
    "Konsisten TOP 🌟": 1,
    "Terus Meningkat 📈": 2,
    "Fluktuatif 🔄": 3,
    "Terus Menurun 📉": 4,
    "Konsisten BOTTOM ⚠️": 5,
    "Data Baru 🆕": 6,
}

METRIC_SUM_COLS = ["Total Penambahan Hari Ini", "Total Penerimaan Hari Ini"]
METRIC_AVG_COLS = [
    "Jumlah total data hari ini", "Jumlah Panggilan", "Jumlah Panggilan Berhasil",
    "Waktu Kerja Rata-rata", "Jumlah SMS", "Jumlah Pengiriman Template WABA"
]
ALL_METRICS = METRIC_SUM_COLS + METRIC_AVG_COLS

# ---------------------------------------------------------
# TRANSLASI JUDUL KOLOM KE MANDARIN
# ---------------------------------------------------------
HEADER_TRANSLATIONS = {
    "ID Agent": "ID Agent<br>(坐席工号)",
    "Nama Agent": "Nama Agent<br>(坐席姓名)",
    "ID Pengguna": "ID Pengguna<br>(用户ID)",
    "Nama Pengguna": "Nama Pengguna<br>(用户名)",
    "Perusahaan": "Perusahaan<br>(公司)",
    "Status Kerja": "Status Kerja<br>(工作状态)",
    "Hari Kerja": "Hari Kerja<br>(出勤天数)",
    "Bucket": "Bucket<br>(逾期阶段)",
    "Hiredate": "Hiredate<br>(入职日期)",
    "Project": "Project<br>(项目)",
    "Apakah Ikut Berkompetisi": "Ikut Kompetisi<br>(是否参赛)",
    "Rank": "Rank<br>(排名)",
    "Ranking": "Ranking<br>(排名)",
    "Ranking (Avg)": "Ranking Avg<br>(平均排名)",
    "Ranking Interval": "Ranking Interval<br>(排名区间)",
    "Nama TL": "Nama TL<br>(组长姓名)",
    "Nama Team Leader": "Nama Team Leader<br>(组长姓名)",
    "Total Agent": "Total Agent Aktif<br>(在职总坐席)",
    "Total Penambahan (Monthly)": "Total Penambahan (M)<br>(月累计新增本金)",
    "Total Penerimaan (Monthly)": "Total Penerimaan (M)<br>(月累计回款总额)",
    "Recovery Rate (Monthly)": "Recovery Rate (M)<br>(月催回率)",
    "Recovery Rate (%)": "Recovery Rate<br>(催回率)",
    "Recovery Rate": "Recovery Rate<br>(催回率)",
    "Gap": "Gap<br>(差距率)",
    "Gap Amount": "Gap Amount<br>(差距金额)",
    "Gap (+1 Rank)": "Gap (+1 Rank)<br>(升1名差距金额)",
    "Gap (Rank #1)": "Gap (Rank #1)<br>(登顶差距金额)",
    "Penambahan (Hari Ini)": "Penambahan (Today)<br>(今日新增本金)",
    "Penerimaan (Hari Ini)": "Penerimaan (Today)<br>(今日回款总额)",
    "Recovery Rate (Hari Ini)": "Recovery Rate (Today)<br>(今日催回率)",
    "Status WL": "Status WL<br>(警告信状态)",
    "Peringkat di Bucket": "Rank / Total<br>(组内排名/总数)",
    "Persentil Rank": "Persentil<br>(排名百分比)",
    "Periode Evaluasi": "Periode<br>(考核周期)",
    "Jumlah total data hari ini": "Total Data<br>(今日分配量)",
    "Jumlah Panggilan": "Jumlah Panggilan<br>(呼叫次数)",
    "Connected Rate": "Connected Rate<br>(接通率)",
    "Waktu Kerja Rata-rata": "Waktu Kerja<br>(工作时长)",
    "Jumlah SMS": "Jumlah SMS<br>(短信发送量)",
    "Jumlah Pengiriman Template WABA": "Template WABA<br>(WABA模板)",
    "Penambahan": "Penambahan<br>(今日新增本金)",
    "Penerimaan": "Penerimaan<br>(今日回款总额)",
    "Status TL": "Status TL<br>(组长状态)",
    "Rata-rata Rate (All-Time)": "Avg Rate (All-Time)<br>(历史平均催回率)",
    "Analisis Kinerja": "Analisis Kinerja<br>(绩效分析)",
    "Kategori Agent": "Kategori Agent<br>(员工属性)",
    "% Agent Baru": "% Agent Baru<br>(新坐席占比)",
    "% Agent Lama": "% Agent Lama<br>(老坐席占比)",
    "Baru: Top 30%": "Baru: Top 30%<br>(新-前30%)",
    "Baru: Middle 30-70%": "Baru: Mid 30-70%<br>(新-中40%)",
    "Baru: Bottom 30%": "Baru: Bot 30%<br>(新-后30%)",
    "Lama: Top 30%": "Lama: Top 30%<br>(老-前30%)",
    "Lama: Middle 30-70%": "Lama: Mid 30-70%<br>(老-中40%)",
    "Lama: Bottom 30%": "Lama: Bot 30%<br>(老-后30%)",
    "jadwal_training": "Jadwal Training (Hari & Jam)<br>(培训时间安排)",
    "topic_pembahasan": "Topic Pembahasan<br>(培训主题与要点)",
    "link_meeting": "Link Gmeet / Zoom<br>(线上会议链接)",
}

def translate_columns(df):
    new_cols = {}
    for c in df.columns:
        if c in HEADER_TRANSLATIONS:
            new_cols[c] = HEADER_TRANSLATIONS[c]
        else:
            if "Bln" in str(c):
                new_cols[c] = f"{c}<br>(月份数据)"
            elif " (" in str(c):
                new_cols[c] = str(c).replace(" (", "<br>(")
            else:
                new_cols[c] = c
    return df.rename(columns=new_cols)

# ---------------------------------------------------------
# RENDER TABEL EKSEKUTIF
# ---------------------------------------------------------
def display_full_table(df_input, fit_content=False):
    df = df_input.data if hasattr(df_input, "data") else df_input

    container_css = "display: inline-block; width: auto; max-width: 100%;" if fit_content else "width: 100%;"
    tbl_css = "width: auto;" if fit_content else "width: 100%;"

    html = [
        f'<div style="{container_css} overflow-x: auto; margin-bottom: 0.9rem; '
        f'border: 1.5px solid #94a3b8; border-radius: 6px; box-shadow: 0 4px 10px rgba(15, 23, 42, 0.08); background-color: #ffffff;">'
    ]
    html.append(
        f'<table style="{tbl_css} border-collapse: separate; border-spacing: 0; '
        'font-size: 0.81rem; text-align: center; background-color: #ffffff;">'
    )
    
    html.append('<thead><tr>')
    for col in df.columns:
        col_header = str(col).replace("\n", "<br>")
        html.append(
            f'<th style="position: sticky; top: 0; z-index: 5; background: #0f172a; '
            f'color: #ffffff; padding: 9px 12px; font-weight: 700; white-space: nowrap; line-height: 1.25; vertical-align: middle; '
            f'border-bottom: 2.5px solid #020617; border-right: 1px solid #475569; letter-spacing: 0.3px;">{col_header}</th>'
        )
    html.append('</tr></thead><tbody>')
    
    for idx, row in df.iterrows():
        row_str = " ".join([str(v) for v in row.values]).upper()
        is_mpt05 = "MPT05" in row_str
        is_resign = "RESIGN" in row_str
        is_wl = "WL" in row_str or "WARNING" in row_str

        if is_wl:
            row_style = "background-color: #fee2e2; font-weight: 700; color: #991b1b;"
        elif is_resign:
            row_style = "background-color: #f1f5f9; color: #94a3b8;"
        elif is_mpt05:
            row_style = "background-color: #fef3c7; font-weight: 700; color: #92400e;"
        else:
            bg_zebra = "#ffffff" if idx % 2 == 0 else "#f8fafc"
            row_style = f"background-color: {bg_zebra}; color: #0f172a; font-weight: 500;"
            
        html.append(f'<tr style="{row_style}">')
        for col_name, val in zip(df.columns, row.values):
            val_str = str(val)
            
            if col_name in ["Rank", "Rank<br>(排名)", "Ranking", "Ranking<br>(排名)"]:
                if val_str in ["#1", "1"]:
                    val_render = '<span style="background: #fde047; color: #713f12; padding: 2px 7px; border-radius: 9999px; font-weight: 800; font-size: 0.76rem; border: 1px solid #ca8a04;">#1</span>'
                elif val_str in ["#2", "2"]:
                    val_render = '<span style="background: #e2e8f0; color: #1e293b; padding: 2px 7px; border-radius: 9999px; font-weight: 800; font-size: 0.76rem; border: 1px solid #94a3b8;">#2</span>'
                elif val_str in ["#3", "3"]:
                    val_render = '<span style="background: #fed7aa; color: #9a3412; padding: 2px 7px; border-radius: 9999px; font-weight: 800; font-size: 0.76rem; border: 1px solid #ea580c;">#3</span>'
                else:
                    val_render = f'<span style="color: #475569; font-weight: 700;">{val_str}</span>'
            elif "GAP" in str(col_name).upper() and val_str.startswith("-") and val_str != "-":
                val_render = f'<span style="background: #fee2e2; color: #b91c1c; padding: 2px 6px; border-radius: 4px; font-weight: 700; font-family: monospace; border: 1px solid #fca5a5;">{val_str}</span>'
            elif "TOP 10%" in val_str.upper() or val_str.upper() == "TOP10%":
                val_render = '<span style="background: #dcfce7; color: #166534; padding: 2px 7px; border-radius: 4px; font-weight: 800;">TOP 10%</span>'
            elif "BOTTOM 20%" in val_str.upper() or val_str.upper() == "BOTTOM20%":
                val_render = '<span style="background: #fee2e2; color: #991b1b; padding: 2px 7px; border-radius: 4px; font-weight: 800;">BOTTOM 20%</span>'
            elif "HTTP://" in val_str.upper() or "HTTPS://" in val_str.upper():
                val_render = f'<a href="{val_str}" target="_blank" style="color: #2563eb; font-weight: 700; text-decoration: underline;">🔗 Buka Link</a>'
            else:
                val_render = val_str

            html.append(
                f'<td style="padding: 8px 12px; border-bottom: 1px solid #cbd5e1; border-right: 1px solid #cbd5e1; '
                f'vertical-align: middle; white-space: nowrap; font-variant-numeric: tabular-nums;">{val_render}</td>'
            )
        html.append('</tr>')
        
    html.append('</tbody></table></div>')
    st.markdown("".join(html), unsafe_allow_html=True)

# ---------------------------------------------------------
# HELPER PARSER & FORMATTER
# ---------------------------------------------------------
def clean_numeric(val):
    if pd.isna(val): return 0.0
    val_str = str(val).strip().replace(",", "")
    val_str = re.sub(r"[^\d.-]", "", val_str)
    try:
        return float(val_str) if val_str != "" else 0.0
    except ValueError:
        return 0.0

def format_num(val, is_currency=False):
    if pd.isna(val) or val == 0: return "0"
    if is_currency:
        return f"{int(round(val)):,}".replace(",", ".")
    return f"{val:,.2f}".rstrip("0").rstrip(".") if isinstance(val, float) else str(val)

def normalize_clean(val):
    if pd.isna(val): return ""
    cleaned = str(val).replace("\xa0", " ").strip().upper()
    return " ".join(cleaned.split())

def map_status_kerja(st_raw):
    s = str(st_raw).strip().lower()
    if "mengundurkan" in s or "resign" in s or "tidak aktif" in s or "non aktif" in s:
        return "Resign"
    elif "aktif" in s or "istirahat" in s:
        return "Aktif"
    return "Resign" if s != "" and s != "nan" else "-"

# 1. Ranking Interval Asli untuk Menu Agent Performance (Top 10%, 10-20%, 20-30%, dst.)
def get_ranking_interval(rank, total):
    if total <= 0: return "-"
    pct = rank / total
    if pct <= 0.10:
        return "TOP 10%"
    elif pct <= 0.20:
        return "10–20%"
    elif pct <= 0.30:
        return "20–30%"
    elif pct <= 0.50:
        return "30–50%"
    elif pct <= 0.70:
        return "50–70%"
    elif pct <= 0.80:
        return "70–80%"
    else:
        return "Bottom 20%"

# 2. Ranking Interval Khusus Menu Komposisi Agent (3 Kelompok)
def get_interval_3cat(rank, total):
    if total <= 0: return "-"
    pct = rank / total
    if pct <= 0.30:
        return "Top 30%"
    elif pct <= 0.70:
        return "Middle 30-70%"
    else:
        return "Bottom 30%"

# ---------------------------------------------------------
# 1. LOAD DATABASE KARYAWAN & TAB MASTER
# ---------------------------------------------------------
@st.cache_data(ttl=60, show_spinner=False)
def load_database_and_master(file_path):
    if not os.path.exists(file_path):
        return None, None
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names

        df_karyawan = pd.read_excel(file_path, sheet_name=sheet_names[0])
        df_master = pd.read_excel(file_path, sheet_name=sheet_names[1]) if len(sheet_names) > 1 else pd.DataFrame()

        df_karyawan.columns = [str(c).strip() for c in df_karyawan.columns]
        if not df_master.empty:
            df_master.columns = [str(c).strip() for c in df_master.columns]

        return df_karyawan, df_master
    except Exception as e:
        st.error(f"Error membaca database: {e}")
        return None, None

# ---------------------------------------------------------
# 2. PENGHITUNGAN HARI KERJA AGEN (TERPISAH KETAT LD & SK)
# ---------------------------------------------------------
@st.cache_data(ttl=60, show_spinner=False)
def calculate_agent_working_days(rawdata_folder, current_daily_files):
    agent_days_attended = {}
    absen_base_dir = os.path.join(rawdata_folder, "DailyRawdata_absen")
    cur_month_num = str(datetime.now().month)

    if os.path.exists(absen_base_dir):
        for item in os.listdir(absen_base_dir):
            p = os.path.join(absen_base_dir, item)
            if not os.path.isdir(p):
                continue

            item_upper = item.upper()
            if not (item.startswith(cur_month_num) or "OKT" in item_upper or "OCT" in item_upper):
                continue

            if "SK" in item_upper:
                proj_cat = "SK"
            elif "LD" in item_upper:
                proj_cat = "LD"
            else:
                continue

            daily_excel_files = glob.glob(os.path.join(p, "*.xlsx")) + glob.glob(os.path.join(p, "*.xls"))
            for d_file in daily_excel_files:
                day_tag = os.path.splitext(os.path.basename(d_file))[0]
                try:
                    df_day = pd.read_excel(d_file)
                    df_day.columns = [str(c).strip() for c in df_day.columns]

                    for _, r in df_day.iterrows():
                        id_raw = normalize_clean(r.get("Nama Pengguna", ""))
                        nm_raw = normalize_clean(r.get("Nama", ""))

                        if id_raw:
                            k_id = (id_raw, proj_cat)
                            agent_days_attended.setdefault(k_id, set()).add(f"{proj_cat}_{day_tag}")
                        if nm_raw and nm_raw != id_raw:
                            k_nm = (nm_raw, proj_cat)
                            agent_days_attended.setdefault(k_nm, set()).add(f"{proj_cat}_{day_tag}")
                except Exception:
                    pass

    for d_path in current_daily_files:
        fname = os.path.basename(d_path).upper()
        proj_cat = "SK" if "SK" in fname else "LD"

        try:
            df_curr = pd.read_excel(d_path)
            df_curr.columns = [str(c).strip() for c in df_curr.columns]

            for _, r in df_curr.iterrows():
                pen_today = clean_numeric(r.get("Total Penambahan Hari Ini", 0))
                if pen_today > 0:
                    id_raw = normalize_clean(r.get("Nama Pengguna", ""))
                    nm_raw = normalize_clean(r.get("Nama", ""))

                    if id_raw:
                        k_id = (id_raw, proj_cat)
                        agent_days_attended.setdefault(k_id, set()).add(f"{proj_cat}_today")
                    if nm_raw and nm_raw != id_raw:
                        k_nm = (nm_raw, proj_cat)
                        agent_days_attended.setdefault(k_nm, set()).add(f"{proj_cat}_today")
        except Exception:
            pass

    working_days_map = {}
    for k, days_set in agent_days_attended.items():
        working_days_map[k] = len(days_set)

    return working_days_map

# ---------------------------------------------------------
# 3. SCAN SELURUH FOLDER RAWDATA
# ---------------------------------------------------------
def natural_sort_month_key(file_path):
    fname = os.path.basename(file_path)
    m = re.match(r"^(\d+)", fname)
    if m:
        return (int(m.group(1)), fname)
    return (0, fname)

@st.cache_data(ttl=60, show_spinner=False)
def read_excel_standard(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = [str(c).strip() for c in df.columns]
        fname = os.path.basename(file_path).upper()
        detected_proj = "SK" if "_SK" in fname or "-SK" in fname or " SK" in fname else "LD"
        df["_source_project"] = detected_proj

        for col in ALL_METRICS:
            if col in df.columns:
                df[col] = df[col].apply(clean_numeric)
            else:
                df[col] = 0.0
        return df
    except Exception:
        return pd.DataFrame()

def scan_dashboard_folders(base_rawdata_dir):
    bulanan_dir = os.path.join(base_rawdata_dir, "Bulanan")
    if not os.path.exists(bulanan_dir):
        bulanan_dir = os.path.join(base_rawdata_dir, "rawdata-Bulanan")

    mingguan_dir = os.path.join(base_rawdata_dir, "Mingguan")
    if not os.path.exists(mingguan_dir):
        mingguan_dir = os.path.join(base_rawdata_dir, "1. Weekly & Monthly")

    all_bulanan_files = {}
    monthly_files = []
    if os.path.exists(bulanan_dir):
        b_excels = glob.glob(os.path.join(bulanan_dir, "*.xlsx")) + glob.glob(os.path.join(bulanan_dir, "*.xls"))
        for bf in b_excels:
            m = re.match(r"^(\d+)", os.path.basename(bf))
            p_key = m.group(1).zfill(2) if m else "00"
            all_bulanan_files.setdefault(p_key, []).append(bf)

        sk_b = sorted([f for f in b_excels if "SK" in os.path.basename(f).upper()], key=natural_sort_month_key)
        ld_b = sorted([f for f in b_excels if "LD" in os.path.basename(f).upper()], key=natural_sort_month_key)
        if sk_b: monthly_files.append(sk_b[-1])
        if ld_b: monthly_files.append(ld_b[-1])

    weekly_dict = {}
    if monthly_files:
        weekly_dict["Monthly"] = monthly_files

    if os.path.exists(mingguan_dir):
        w_excels = glob.glob(os.path.join(mingguan_dir, "*.xlsx")) + glob.glob(os.path.join(mingguan_dir, "*.xls"))
        for wf in w_excels:
            bw = os.path.basename(wf).upper()
            if "W1" in bw or "1-7" in bw or "1-8" in bw:
                weekly_dict.setdefault("W1", []).append(wf)
            elif "W2" in bw or "8-14" in bw or "9-14" in bw:
                weekly_dict.setdefault("W2", []).append(wf)
            elif "W3" in bw or "15-21" in bw:
                weekly_dict.setdefault("W3", []).append(wf)
            elif "W4" in bw or "22-28" in bw:
                weekly_dict.setdefault("W4", []).append(wf)

    mtd_agent_files = []
    all_rawdata_excels = glob.glob(os.path.join(base_rawdata_dir, "*.xlsx")) + glob.glob(os.path.join(base_rawdata_dir, "*.xls"))
    sk_mtd_ag = sorted([f for f in all_rawdata_excels if "MTDAGENT" in os.path.basename(f).upper() and "SK" in os.path.basename(f).upper()])
    ld_mtd_ag = sorted([f for f in all_rawdata_excels if "MTDAGENT" in os.path.basename(f).upper() and "LD" in os.path.basename(f).upper()])
    if sk_mtd_ag: mtd_agent_files.append(sk_mtd_ag[-1])
    if ld_mtd_ag: mtd_agent_files.append(ld_mtd_ag[-1])

    daily_cands = glob.glob(os.path.join(base_rawdata_dir, "workStatistics*.xlsx")) + glob.glob(os.path.join(base_rawdata_dir, "workStatistics*.xls"))
    daily_files = []
    sk_dailies = sorted([f for f in daily_cands if "SK" in os.path.basename(f).upper()])
    ld_dailies = sorted([f for f in daily_cands if "LD" in os.path.basename(f).upper()])
    if sk_dailies: daily_files.append(sk_dailies[-1])
    if ld_dailies: daily_files.append(ld_dailies[-1])

    wl_dict = {}
    wl_dir = os.path.join(base_rawdata_dir, "WarningLetter")
    if os.path.exists(wl_dir):
        wl_excels = glob.glob(os.path.join(wl_dir, "*.xlsx")) + glob.glob(os.path.join(wl_dir, "*.xls"))
        for f in wl_excels:
            clean_name = os.path.basename(f).upper()
            m_wl = re.match(r"^(\d+-\d+)[-_ ]?.*", clean_name)
            if m_wl:
                wl_dict.setdefault(m_wl.group(1), []).append(f)

    track_agent_dict = {}
    track_dir = os.path.join(base_rawdata_dir, "TrackRecordAgent")
    if os.path.exists(track_dir):
        t_excels = glob.glob(os.path.join(track_dir, "*.xlsx")) + glob.glob(os.path.join(track_dir, "*.xls"))
        for f in t_excels:
            m_code = re.match(r"^(\d{6})[-_].*", os.path.basename(f))
            if m_code:
                track_agent_dict.setdefault(m_code.group(1), []).append(f)

    return weekly_dict, mtd_agent_files, daily_files, all_bulanan_files, wl_dict, track_agent_dict

# ---------------------------------------------------------
# PROSES AWAL DATA & PEMETAAN KARYAWAN
# ---------------------------------------------------------
df_karyawan, df_master = load_database_and_master(MASTER_FILE)
weekly_dict, mtd_agent_files, daily_files, all_bulanan_dict, wl_dict, track_agent_dict = scan_dashboard_folders(RAWDATA_DIR)
agent_workdays_map = calculate_agent_working_days(RAWDATA_DIR, daily_files)

col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown("### **Performance Collection Automation (催收绩效自动化系统)**")
with col_h2:
    if st.button("Refresh Data (刷新数据)", use_container_width=True):
        try:
            subprocess.run(["git", "pull"], cwd=BASE_DIR, capture_output=True, text=True, timeout=15)
        except Exception:
            pass
        st.cache_data.clear()
        st.rerun()

if df_karyawan is None:
    st.error("File `data_karyawan.xlsx` tidak ditemukan.")
    st.stop()

tl_col = "Nama TL" if "Nama TL" in df_karyawan.columns else "Perusahaan"
bucket_col = "Bucket" if "Bucket" in df_karyawan.columns else "Bucket"
project_col = "Project" if "Project" in df_karyawan.columns else "Project"

karyawan_dict = {}
for _, row in df_karyawan.iterrows():
    id_akun = normalize_clean(row.get("ID Pengguna", ""))
    nama_orang = normalize_clean(row.get("Nama Pengguna", ""))
    proj_exact = str(row.get(project_col, "")).strip().upper()
    proj_cat = "SK" if proj_exact == "SK" else "LD"
    raw_status = row.get("Status Kerja", "")

    info = {
        "Nama TL": str(row.get(tl_col, "")).strip(),
        "Bucket": str(row.get(bucket_col, "")).strip(),
        "Project": proj_exact,
        "Status Kerja": map_status_kerja(raw_status),
        "Hiredate": row.get("Hiredate", None),
        "ID_Akun_Resmi": str(row.get("ID Pengguna", "")).strip(),
        "Nama_Orang_Resmi": str(row.get("Nama Pengguna", "")).strip(),
    }

    if id_akun:
        karyawan_dict[(id_akun, proj_cat)] = info
        karyawan_dict[(id_akun, proj_exact)] = info
    if nama_orang:
        karyawan_dict[(nama_orang, proj_cat)] = info
        karyawan_dict[(nama_orang, proj_exact)] = info

master_rules = []
if df_master is not None and not df_master.empty:
    for _, r in df_master.iterrows():
        master_rules.append({
            "tl": normalize_clean(r.get("Nama TL", "")),
            "bucket": normalize_clean(r.get("Bucket", "")),
            "project": normalize_clean(r.get("Project", "")),
            "perusahaan_master": str(r.get("Perusahaan", "")).strip(),
        })

def get_master_status(tl, bucket, project):
    t_clean = normalize_clean(tl)
    b_clean = normalize_clean(bucket)
    p_clean = normalize_clean(project)

    for rule in master_rules:
        m_proj = (rule["project"] == "") or (rule["project"] == p_clean)
        m_buck = (rule["bucket"] == "") or (rule["bucket"] == b_clean)
        m_tl = (rule["tl"] == "") or (rule["tl"] == t_clean)
        if m_proj and m_buck and m_tl:
            return "Ya", rule["perusahaan_master"]

    return "Tidak", "Non-Kompetisi"

# ---------------------------------------------------------
# SINKRONISASI DATA DAILY KE NAMA TL
# ---------------------------------------------------------
daily_comp_map = {}
daily_tl_map = {}

if daily_files:
    for df_path in daily_files:
        try:
            df_d = pd.read_excel(df_path)
            df_d.columns = [str(c).strip() for c in df_d.columns]
            fname = os.path.basename(df_path).upper()
            proj_hint = "SK" if "SK" in fname else "LD"
            p_cat = "SK" if "SK" in proj_hint else "LD"

            for _, r in df_d.iterrows():
                id_a = normalize_clean(r.get("Nama Pengguna", ""))
                nm_a = normalize_clean(r.get("Nama", ""))
                acc_k = id_a if id_a else nm_a
                if not acc_k: continue

                k_info = karyawan_dict.get((acc_k, p_cat)) or karyawan_dict.get((nm_a, p_cat)) or {}
                tl_val = k_info.get("Nama TL", "")
                buck_val = k_info.get("Bucket", "")
                proj_val = k_info.get("Project", proj_hint)

                if not tl_val or tl_val == "Unknown":
                    if "MPT05-S1-W4" in acc_k:
                        tl_val = "W-MPT05-2" if p_cat == "LD" else "MPT05-TL-004"
                        buck_val = "S1"
                    elif "MPT05-S1-W3" in acc_k:
                        tl_val = "W-MPT05-3" if p_cat == "LD" else "MPT05-TL-003"
                        buck_val = "S1"
                    elif "MPT05-S1-BA" in acc_k:
                        tl_val = "W-MPT05-BA"
                        buck_val = "S1"
                        proj_val = "LD-BA"
                    elif "MPT05-S1-W0" in acc_k or "MPT05-S1-W2" in acc_k:
                        tl_val = "W-MPT05" if p_cat == "LD" else "MPT05-TL-001"
                        buck_val = "S1"
                    elif "MPT05-S0" in acc_k:
                        tl_val = "W-MPT05-S0-01"
                        buck_val = "S0"
                    elif "W-MPT05-C" in acc_k:
                        tl_val = "W-MPT05-C"
                        buck_val = "C"
                        proj_val = "LD-C"
                    elif "DC01" in acc_k: tl_val, buck_val = "DC01", "S1"
                    elif "DC06" in acc_k: tl_val, buck_val = "DC06", "S1"
                    elif "DC07" in acc_k: tl_val, buck_val = "DC07", "S1"
                    elif "INTERNAL" in acc_k: tl_val, buck_val = "Internal Collection", "S1"

                is_comp, perush = get_master_status(tl_val, buck_val, proj_val)
                if is_comp == "Ya":
                    pen_d = clean_numeric(r.get("Total Penambahan Hari Ini", 0))
                    rec_d = clean_numeric(r.get("Total Penerimaan Hari Ini", 0))

                    k_comp = (proj_val, buck_val, perush)
                    daily_comp_map.setdefault(k_comp, {"pen": 0.0, "rec": 0.0})
                    daily_comp_map[k_comp]["pen"] += pen_d
                    daily_comp_map[k_comp]["rec"] += rec_d

                    k_tl = (proj_val, buck_val, tl_val)
                    daily_tl_map.setdefault(k_tl, {"pen": 0.0, "rec": 0.0})
                    daily_tl_map[k_tl]["pen"] += pen_d
                    daily_tl_map[k_tl]["rec"] += rec_d
        except Exception:
            pass

# ---------------------------------------------------------
# FUNGSI MEMBACA FILE DEPARTEMEN BULANAN & MINGGUAN
# ---------------------------------------------------------
def process_department_weekly_data(file_list):
    records = []
    for f in file_list:
        try:
            df = pd.read_excel(f)
            df.columns = [str(c).strip() for c in df.columns]
            fname = os.path.basename(f).upper()
            proj_hint = "SK" if "SK" in fname else "LD"

            if "Departemen" in df.columns:
                for _, r in df.iterrows():
                    tl_raw = str(r.get("Departemen", "")).strip()
                    bucket_raw = str(r.get("Kelompok", "")).strip()
                    if not tl_raw or not bucket_raw or tl_raw.upper() == "TOTAL":
                        continue

                    pen = clean_numeric(r.get("Jumlah yang dibagikan", 0))
                    rec = clean_numeric(r.get("Jumlah pembayaran", 0))

                    is_comp = "Tidak"
                    matched_comp = "Non-Kompetisi"
                    matched_proj = proj_hint

                    for rule in master_rules:
                        r_proj = rule["project"]
                        r_buck = rule["bucket"]
                        r_tl = rule["tl"]

                        m_proj = (r_proj == "") or (r_proj == proj_hint) or (proj_hint == "LD" and "LD" in r_proj)
                        m_buck = (r_buck == normalize_clean(bucket_raw))
                        m_tl = (r_tl == normalize_clean(tl_raw))

                        if m_buck and m_tl and m_proj:
                            is_comp = "Ya"
                            matched_comp = rule["perusahaan_master"]
                            matched_proj = r_proj if r_proj else proj_hint
                            break

                    if is_comp == "Ya":
                        records.append({
                            "Project": matched_proj,
                            "Bucket": bucket_raw,
                            "Nama TL": tl_raw,
                            "Perusahaan": matched_comp,
                            "Total Penambahan": pen,
                            "Total Penerimaan": rec,
                        })
        except Exception:
            pass
    return pd.DataFrame(records)

# ---------------------------------------------------------
# FUNGSI MEMPROSES PERFORMA AGEN (MTD + DAILY + HARI KERJA + HIREDATE)
# ---------------------------------------------------------
def process_agent_performance_data(mtd_agent_files, daily_files):
    if not mtd_agent_files and not daily_files:
        return pd.DataFrame()

    mtd_dfs = [read_excel_standard(f) for f in mtd_agent_files]
    df_mtd_all = pd.concat([d for d in mtd_dfs if not d.empty], ignore_index=True) if mtd_dfs else pd.DataFrame()

    daily_dfs = [read_excel_standard(f) for f in daily_files]
    df_daily_all = pd.concat([d for d in daily_dfs if not d.empty], ignore_index=True) if daily_dfs else pd.DataFrame()

    daily_map = {}
    if not df_daily_all.empty:
        for _, r in df_daily_all.iterrows():
            id_d = normalize_clean(r.get("Nama Pengguna", ""))
            nm_d = normalize_clean(r.get("Nama", ""))
            p_src = normalize_clean(r.get("_source_project", ""))
            p_cat = "SK" if "SK" in p_src else "LD"

            acc_key = id_d if id_d else nm_d
            if not acc_key: continue

            pen_today = clean_numeric(r.get("Total Penambahan Hari Ini", 0))
            rec_today = clean_numeric(r.get("Total Penerimaan Hari Ini", 0))

            vals = {m: clean_numeric(r.get(m, 0)) for m in ALL_METRICS}
            vals["_raw_id"] = str(r.get("Nama Pengguna", "")).strip()
            vals["_raw_nama"] = str(r.get("Nama", "")).strip()
            vals["_proj"] = p_src
            vals["_proj_cat"] = p_cat
            vals["_is_active_today"] = (pen_today > 0)

            k_tuple = (acc_key, p_cat)
            if k_tuple not in daily_map:
                daily_map[k_tuple] = vals
            else:
                daily_map[k_tuple]["Total Penambahan Hari Ini"] += pen_today
                daily_map[k_tuple]["Total Penerimaan Hari Ini"] += rec_today
                if pen_today > 0:
                    daily_map[k_tuple]["_is_active_today"] = True

    mtd_map = {}
    if not df_mtd_all.empty:
        for _, r in df_mtd_all.iterrows():
            id_m = normalize_clean(r.get("Nama Pengguna", ""))
            nm_m = normalize_clean(r.get("Nama", ""))
            p_src = normalize_clean(r.get("_source_project", ""))
            p_cat = "SK" if "SK" in p_src else "LD"

            acc_key = id_m if id_m else nm_m
            if not acc_key: continue

            pen_val = clean_numeric(r.get("Total Penambahan Hari Ini", 0))
            rec_val = clean_numeric(r.get("Total Penerimaan Hari Ini", 0))

            k_tuple = (acc_key, p_cat)
            if k_tuple not in mtd_map:
                mtd_map[k_tuple] = {
                    "Total Penambahan Hari Ini": pen_val,
                    "Total Penerimaan Hari Ini": rec_val,
                    "_raw_id": str(r.get("Nama Pengguna", "")).strip(),
                    "_raw_nama": str(r.get("Nama", "")).strip(),
                    "_proj": p_src,
                    "_proj_cat": p_cat,
                    "_count": 1
                }
                for m in METRIC_AVG_COLS: mtd_map[k_tuple][m] = clean_numeric(r.get(m, 0))
            else:
                mtd_map[k_tuple]["Total Penambahan Hari Ini"] += pen_val
                mtd_map[k_tuple]["Total Penerimaan Hari Ini"] += rec_val
                mtd_map[k_tuple]["_count"] += 1
                for m in METRIC_AVG_COLS: mtd_map[k_tuple][m] += clean_numeric(r.get(m, 0))

        for k_tuple in mtd_map:
            cnt = mtd_map[k_tuple]["_count"]
            if cnt > 1:
                for m in METRIC_AVG_COLS: mtd_map[k_tuple][m] /= cnt

    all_keys = set(mtd_map.keys()).union(set(daily_map.keys()))
    records = []
    now_dt = datetime.now()
    cur_year = now_dt.year
    cur_month = now_dt.month

    for (acc_key, proj_cat) in all_keys:
        m_vals = mtd_map.get((acc_key, proj_cat), {})
        d_vals = daily_map.get((acc_key, proj_cat), {})

        raw_id_str = d_vals.get("_raw_id", "") or m_vals.get("_raw_id", "") or acc_key
        raw_nm_str = d_vals.get("_raw_nama", "") or m_vals.get("_raw_nama", "") or acc_key
        p_src = d_vals.get("_proj", "") or m_vals.get("_proj", "") or proj_cat

        k_info = karyawan_dict.get((acc_key, proj_cat)) or {}
        if not k_info:
            alt_id = normalize_clean(raw_nm_str)
            k_info = karyawan_dict.get((alt_id, proj_cat)) or karyawan_dict.get((acc_key, p_src)) or {}

        tl_val = k_info.get("Nama TL", "Unknown")
        bucket_val = k_info.get("Bucket", "Unknown")
        project_val = k_info.get("Project", p_src or "Unknown")
        status_val = k_info.get("Status Kerja", "Aktif")
        raw_hiredate = k_info.get("Hiredate", None)

        kat_agent = "Agent Lama"
        hire_str = "-"
        if raw_hiredate is not None and not pd.isna(raw_hiredate):
            try:
                h_dt = pd.to_datetime(raw_hiredate)
                hire_str = h_dt.strftime("%Y-%m-%d")
                if h_dt.year == cur_year and h_dt.month == cur_month:
                    kat_agent = "Agent Baru"
                else:
                    kat_agent = "Agent Lama"
            except Exception:
                pass

        is_comp, perush_master = get_master_status(tl_val, bucket_val, project_val)

        pen_mtd = m_vals.get("Total Penambahan Hari Ini", 0.0)
        rec_mtd = m_vals.get("Total Penerimaan Hari Ini", 0.0)
        pen_daily = d_vals.get("Total Penambahan Hari Ini", 0.0)
        rec_daily = d_vals.get("Total Penerimaan Hari Ini", 0.0)

        clean_id_u = normalize_clean(raw_id_str)
        proj_bucket_cat = "SK" if proj_cat == "SK" or "SK" in str(project_val).upper() else "LD"

        w_days = agent_workdays_map.get((clean_id_u, proj_bucket_cat), 0)
        if w_days == 0:
            alt_nm_u = normalize_clean(raw_nm_str)
            w_days = agent_workdays_map.get((alt_nm_u, proj_bucket_cat), 0)

        is_active_today = d_vals.get("_is_active_today", False)
        tot_pen_agent = pen_mtd + pen_daily
        tot_rec_agent = rec_mtd + rec_daily
        rec_rate_num = (tot_rec_agent / tot_pen_agent * 100.0) if tot_pen_agent > 0 else 0.0

        row_item = {
            "ID Agent": raw_id_str,
            "Nama Pengguna": raw_nm_str,
            "Status Kerja": status_val,
            "Hiredate": hire_str,
            "Kategori Agent": kat_agent,
            "Hari Kerja Angka": w_days,
            "Hari Kerja": f"{w_days} Hari" if w_days > 0 else "-",
            "Nama TL": tl_val,
            "Bucket": bucket_val,
            "Project": project_val,
            "Perusahaan": perush_master,
            "Apakah Ikut Berkompetisi": is_comp,
            "Total Penambahan Hari Ini_Total": tot_pen_agent,
            "Total Penerimaan Hari Ini_Total": tot_rec_agent,
            "Recovery Rate Num": rec_rate_num,
            "Total Penambahan Hari Ini_Daily": pen_daily,
            "Total Penerimaan Hari Ini_Daily": rec_daily,
            "Is Active Today": is_active_today,
            "Jumlah total data hari ini_Daily": d_vals.get("Jumlah total data hari ini", 0.0),
            "Jumlah Panggilan_Daily": d_vals.get("Jumlah Panggilan", 0.0),
            "Jumlah Panggilan Berhasil_Daily": d_vals.get("Jumlah Panggilan Berhasil", 0.0),
            "Waktu Kerja Rata-rata_Daily": d_vals.get("Waktu Kerja Rata-rata", 0.0),
            "Jumlah SMS_Daily": d_vals.get("Jumlah SMS", 0.0),
            "Jumlah Pengiriman Template WABA_Daily": d_vals.get("Jumlah Pengiriman Template WABA", 0.0),
        }
        records.append(row_item)

    df_res_agents = pd.DataFrame(records)

    # Hitung Ranking dan Dua Versi Interval (Kompetisi & Agent Performance)
    if not df_res_agents.empty:
        rank_map = {}
        intv_map = {}
        intv_3cat_map = {}

        df_comp_only = df_res_agents[df_res_agents["Apakah Ikut Berkompetisi"] == "Ya"].copy()
        for (p, b), grp in df_comp_only.groupby(["Project", "Bucket"]):
            grp_sorted = grp.sort_values(by="Recovery Rate Num", ascending=False)
            tot_b = len(grp_sorted)
            for idx_b, (_, r_b) in enumerate(grp_sorted.iterrows()):
                rk = idx_b + 1
                rank_map[(r_b["ID Agent"], r_b["Project"])] = f"#{rk}"
                # Interval Asli untuk Agent Performance
                intv_map[(r_b["ID Agent"], r_b["Project"])] = get_ranking_interval(rk, tot_b)
                # Interval 3 Kategori untuk Komposisi Agent
                intv_3cat_map[(r_b["ID Agent"], r_b["Project"])] = get_interval_3cat(rk, tot_b)

        df_res_agents["Rank MTD"] = df_res_agents.apply(lambda r: rank_map.get((r["ID Agent"], r["Project"]), "-"), axis=1)
        df_res_agents["Ranking Interval"] = df_res_agents.apply(lambda r: intv_map.get((r["ID Agent"], r["Project"]), "-"), axis=1)
        df_res_agents["Ranking Interval 3Cat"] = df_res_agents.apply(lambda r: intv_3cat_map.get((r["ID Agent"], r["Project"]), "-"), axis=1)

    return df_res_agents

# ---------------------------------------------------------
# PROSES DATA AGENT PERFORMA (DEKLARASI GLOBAL PENTING)
# ---------------------------------------------------------
df_agents_processed = process_agent_performance_data(mtd_agent_files, daily_files)

# ---------------------------------------------------------
# FUNGSI PEMBANTU (WARNING LETTER DI GLOBAL SCOPE)
# ---------------------------------------------------------
def get_tl_wl_cutoff(total_teams):
    if total_teams <= 1: return 0
    elif total_teams in [2, 3, 4]: return 1
    elif total_teams == 5: return 2
    elif total_teams == 6: return 3
    else: return total_teams // 2

def calculate_warning_letter_data(df_period_all, periode_name=""):
    if df_period_all.empty: return pd.DataFrame(), pd.DataFrame()
    df_comp = df_period_all[df_period_all["Apakah Ikut Berkompetisi"] == "Ya"].copy()
    if df_comp.empty: return pd.DataFrame(), pd.DataFrame()

    wl_agents, wl_tls = [], []
    for (p, b), grp in df_comp.groupby(["Project", "Bucket"]):
        grp_ag = grp[grp["Hari Kerja Angka"] > 7].copy()
        if not grp_ag.empty:
            grp_ag["_tot_pen"] = grp_ag["Total Penambahan Hari Ini_Total"]
            grp_ag["_tot_rec"] = grp_ag["Total Penerimaan Hari Ini_Total"]
            grp_ag["_rate"] = (grp_ag["_tot_pen"] > 0).astype(int) * (
                grp_ag["_tot_rec"] / grp_ag["_tot_pen"].replace(0, 1) * 100
            )
            grp_ag = grp_ag.sort_values(by="_rate", ascending=False).reset_index(drop=True)
            total_ag = len(grp_ag)

            for idx, row_ag in grp_ag.iterrows():
                rank_ag = idx + 1
                pct_ag = rank_ag / total_ag if total_ag > 0 else 1.0

                if pct_ag > 0.70:
                    perush = str(row_ag.get("Perusahaan", "")).upper()
                    tl_nm = str(row_ag.get("Nama TL", "")).upper()
                    id_ag = str(row_ag.get("ID Agent", "")).upper()

                    if "MPT05" in perush or "MPT05" in tl_nm or "MPT05" in id_ag:
                        wl_agents.append({
                            "Periode Evaluasi": periode_name,
                            "Project": p,
                            "Bucket": b,
                            "Rank": f"#{rank_ag}",
                            "Peringkat di Bucket": f"Rank #{rank_ag} dari {total_ag} Agen (>7 Hari)",
                            "Persentil Rank": f"Bottom {int(round((1 - pct_ag + (1/total_ag))*100))}%",
                            "Nama TL": row_ag.get("Nama TL", ""),
                            "ID Agent": row_ag.get("ID Agent", ""),
                            "Nama Pengguna": row_ag.get("Nama Pengguna", ""),
                            "Status Kerja": row_ag.get("Status Kerja", ""),
                            "Hari Kerja": row_ag.get("Hari Kerja", "-"),
                            "Total Penambahan (Periode)": format_num(row_ag["_tot_pen"], is_currency=True),
                            "Total Penerimaan (Periode)": format_num(row_ag["_tot_rec"], is_currency=True),
                            "Recovery Rate (%)": f"{row_ag['_rate']:.2f}%",
                            "Status WL": "⚠️ WL AGENT",
                        })

        tl_list = []
        for tl_name, tl_group in grp.groupby("Nama TL"):
            t_pen = tl_group["Total Penambahan Hari Ini_Total"].sum()
            t_rec = tl_group["Total Penerimaan Hari Ini_Total"].sum()
            t_rate = (t_rec / t_pen * 100) if t_pen > 0 else 0.0
            tl_list.append({
                "Project": p,
                "Bucket": b,
                "Nama TL": tl_name,
                "Perusahaan": tl_group.iloc[0].get("Perusahaan", ""),
                "Total Agent": len(tl_group),
                "_pen": t_pen,
                "_rec": t_rec,
                "_rate": t_rate,
            })

        df_tl_rank = pd.DataFrame(tl_list)
        if not df_tl_rank.empty:
            df_tl_rank = df_tl_rank.sort_values(by="_rate", ascending=False).reset_index(drop=True)
            total_teams = len(df_tl_rank)
            cutoff = get_tl_wl_cutoff(total_teams)

            if cutoff > 0:
                bottom_tls = df_tl_rank.tail(cutoff)
                for _, row_tl in bottom_tls.iterrows():
                    tl_nm = str(row_tl.get("Nama TL", "")).upper()
                    perush = str(row_tl.get("Perusahaan", "")).upper()

                    if "MPT05" in tl_nm or "MPT05" in perush:
                        actual_rank = df_tl_rank[df_tl_rank["Nama TL"] == row_tl["Nama TL"]].index[0] + 1
                        wl_tls.append({
                            "Periode Evaluasi": periode_name,
                            "Project": p,
                            "Bucket": b,
                            "Rank": f"#{actual_rank}",
                            "Peringkat di Bucket": f"Rank #{actual_rank} dari {total_teams} Tim",
                            "Nama TL": row_tl["Nama TL"],
                            "Perusahaan": row_tl["Perusahaan"],
                            "Total Agent": row_tl["Total Agent"],
                            "Total Penambahan (Periode)": format_num(row_tl["_pen"], is_currency=True),
                            "Total Penerimaan (Periode)": format_num(row_tl["_rec"], is_currency=True),
                            "Recovery Rate (%)": f"{row_tl['_rate']:.2f}%",
                            "Status WL": "🚨 WL TL",
                        })

    return pd.DataFrame(wl_agents), pd.DataFrame(wl_tls)

# ---------------------------------------------------------
# RENDER DUAL TABLES UNTUK MENU AGENT PERFORMANCE
# (INTERVAL ASLI: Top 10%, 10-20%, 20-30%, 30-50%, 50-70%, 70-80%, Bottom 20%)
# ---------------------------------------------------------
def render_agent_split_tables(target_agents_df, title_text=None):
    if title_text:
        st.markdown(f'<div class="clean-group-title">{title_text}</div>', unsafe_allow_html=True)

    rows_m = []
    rows_d = []
    for _, r in target_agents_df.iterrows():
        st_kerja = r.get("Status Kerja", "Aktif")
        h_kerja = r.get("Hari Kerja", "-")

        pen_m = r["Total Penambahan Hari Ini_Total"]
        rec_m = r["Total Penerimaan Hari Ini_Total"]
        rate_m = (rec_m / pen_m * 100) if pen_m > 0 else 0.0

        rows_m.append({
            "Project": r.get("Project", ""),
            "Bucket": r.get("Bucket", ""),
            "Nama TL": r.get("Nama TL", ""),
            "ID Agent": r.get("ID Agent", ""),
            "Nama Pengguna": r.get("Nama Pengguna", ""),
            "Status Kerja": st_kerja,
            "Hari Kerja": h_kerja,
            "_tot_pen": pen_m,
            "Total Penambahan (Monthly)": format_num(pen_m, is_currency=True),
            "Total Penerimaan (Monthly)": format_num(rec_m, is_currency=True),
            "_rate_num": rate_m,
            "Recovery Rate (%)": f"{rate_m:.2f}%",
        })

        if r.get("Is Active Today", False):
            pen_d = r["Total Penambahan Hari Ini_Daily"]
            rec_d = r["Total Penerimaan Hari Ini_Daily"]
            rate_d = (rec_d / pen_d * 100) if pen_d > 0 else 0.0

            rows_d.append({
                "ID Agent": r.get("ID Agent", ""),
                "Nama Pengguna": r.get("Nama Pengguna", ""),
                "Status Kerja": st_kerja,
                "Hari Kerja": h_kerja,
                "_daily_pen": pen_d,
                "Penambahan (Hari Ini)": format_num(pen_d, is_currency=True),
                "Penerimaan (Hari Ini)": format_num(rec_d, is_currency=True),
                "_rate_num": rate_d,
                "Recovery Rate (Hari Ini)": f"{rate_d:.2f}%",
                "Jumlah Panggilan": f"{r['Jumlah Panggilan_Daily']:.2f}",
                "Waktu Kerja Rata-rata": f"{r['Waktu Kerja Rata-rata_Daily']:.2f}",
                "Jumlah SMS": f"{r['Jumlah SMS_Daily']:.2f}",
                "Jumlah Pengiriman Template WABA": f"{r['Jumlah Pengiriman Template WABA_Daily']:.2f}",
            })

    # Tabel Bulanan (Interval 7 Kategori Asli)
    df_am = pd.DataFrame(rows_m)
    if not df_am.empty:
        df_am = df_am.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
        total_n = len(df_am)
        gaps, gaps_amt, intervals = ["-"], ["-"], [get_ranking_interval(1, total_n)]
        for i in range(1, total_n):
            diff = df_am.loc[i - 1, "_rate_num"] - df_am.loc[i, "_rate_num"]
            pen_c = df_am.loc[i, "_tot_pen"]
            if diff > 0 and pen_c > 0:
                gaps.append(f"+{diff:.2f}%")
                gaps_amt.append(f"-{format_num((diff / 100.0) * pen_c, is_currency=True)}")
            else:
                gaps.append("0.00%")
                gaps_amt.append("-")
            intervals.append(get_ranking_interval(i + 1, total_n))

        df_am["Gap"] = gaps
        df_am["Gap Amount"] = gaps_amt
        df_am["Ranking Interval"] = intervals
        df_am["Rank"] = [f"#{i+1}" for i in range(total_n)]
        cols_am = ["Rank", "Ranking Interval", "Project", "Bucket", "Nama TL", "ID Agent", "Nama Pengguna", "Status Kerja", "Hari Kerja", "Total Penambahan (Monthly)", "Total Penerimaan (Monthly)", "Recovery Rate (%)", "Gap", "Gap Amount"]
        df_am = df_am[cols_am]

    # Tabel Harian (Interval 7 Kategori Asli)
    df_ad = pd.DataFrame(rows_d)
    if not df_ad.empty:
        df_ad = df_ad.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
        total_nd = len(df_ad)
        gaps_d, gaps_amt_d, intervals_d = ["-"], ["-"], [get_ranking_interval(1, total_nd)]
        for i in range(1, total_nd):
            diff_d = df_ad.loc[i - 1, "_rate_num"] - df_ad.loc[i, "_rate_num"]
            pen_dc = df_ad.loc[i, "_daily_pen"]
            if diff_d > 0 and pen_dc > 0:
                gaps_d.append(f"+{diff_d:.2f}%")
                gaps_amt_d.append(f"-{format_num((diff_d / 100.0) * pen_dc, is_currency=True)}")
            else:
                gaps_d.append("0.00%")
                gaps_amt_d.append("-")
            intervals_d.append(get_ranking_interval(i + 1, total_nd))

        df_ad["Gap"] = gaps_d
        df_ad["Gap Amount"] = gaps_amt_d
        df_ad["Ranking Interval"] = intervals_d
        df_ad["Rank"] = [f"#{i+1}" for i in range(total_nd)]
        cols_ad = ["Rank", "Ranking Interval", "ID Agent", "Nama Pengguna", "Status Kerja", "Hari Kerja", "Penambahan (Hari Ini)", "Penerimaan (Hari Ini)", "Recovery Rate (Hari Ini)", "Gap", "Gap Amount", "Jumlah Panggilan", "Waktu Kerja Rata-rata", "Jumlah SMS", "Jumlah Pengiriman Template WABA"]
        df_ad = df_ad[cols_ad]

    col_left, col_right = st.columns([1, 1], gap="small")
    with col_left:
        st.markdown('<div class="section-banner-title">Agent Bulanan (月度坐席绩效)</div>', unsafe_allow_html=True)
        if not df_am.empty:
            display_full_table(translate_columns(df_am))
        else:
            st.info("Tidak ada data bulanan.")
    with col_right:
        st.markdown('<div class="section-banner-title tl">Agent Harian (今日实时绩效 - Khusus Hadir)</div>', unsafe_allow_html=True)
        if not df_ad.empty:
            display_full_table(translate_columns(df_ad))
        else:
            st.caption("Tidak ada agen yang terdistribusi pada file daily hari ini.")

def render_agent_bucket_tabs(df_agent_source):
    if df_agent_source.empty:
        st.info("Tidak ada data agen.")
        return

    sub_tabs = st.tabs([label for label, _, _ in TARGET_BUCKETS])
    for tab, (tab_label, b_filter, p_filter) in zip(sub_tabs, TARGET_BUCKETS):
        with tab:
            if b_filter and p_filter:
                filtered_df = df_agent_source[(df_agent_source["Bucket"] == b_filter) & (df_agent_source["Project"] == p_filter)]
                if filtered_df.empty:
                    st.info("Tidak ada data agen pada bucket ini.")
                    continue
                render_agent_split_tables(filtered_df)
            else:
                projects = sorted(df_agent_source["Project"].dropna().unique().tolist())
                for p in projects:
                    p_df = df_agent_source[df_agent_source["Project"] == p]
                    p_buckets = sorted(p_df["Bucket"].dropna().unique().tolist())
                    for b in p_buckets:
                        b_df = p_df[p_df["Bucket"] == b]
                        if b_df.empty: continue
                        render_agent_split_tables(b_df, f"Project: {p} &nbsp;|&nbsp; Bucket: {b} ({len(b_df)} Total Agent)")

# =========================================================
# TAB UTAMA: NAVIGASI MENU (LENGKAP)
# =========================================================
tab_summary_weekly, tab_team_perf, tab_agent_senior, tab_agent_junior, tab_agent_comp, tab_training, tab_productivity, tab_wl, tab_track_agent, tab_track_tl, tab_track_comp = st.tabs([
    "Summary Weekly & Monthly",
    "Team Performance",
    "Performance Agent (>7 Hari Kerja)",
    "Performance Agent (<=7 Hari Kerja)",
    "Komposisi Agent (Baru vs Lama)",
    "Schedule Training Team Leader",
    "Productivity Agents",
    "Warning Letter",
    "Track Record Performance Agents",
    "Track Record Performance TL",
    "Track Record Performance Gabungan",
])

# =========================================================
# MENU 1: SUMMARY WEEKLY & MONTHLY (LAYOUT 3 ROW)
# =========================================================
with tab_summary_weekly:
    if not weekly_dict:
        st.info("Belum ada file di dalam folder `rawdata/Bulanan/` atau `rawdata/Mingguan/`.")
    else:
        summary_order = ["Monthly", "W1", "W2", "W3", "W4"]
        available_periods = [p for p in summary_order if p in weekly_dict]

        weekly_processed = {}
        for p_label in available_periods:
            df_w_proc = process_department_weekly_data(weekly_dict[p_label])
            if not df_w_proc.empty:
                weekly_processed[p_label] = df_w_proc

        if not weekly_processed:
            st.warning("Tidak ditemukan data kompetisi pada folder `Bulanan` maupun `Mingguan`.")
        else:
            sum_sub_tabs = st.tabs([label for label, _, _ in TARGET_BUCKETS])

            def render_summary_three_rows_block(b_filt, p_filt, title_header=None):
                if title_header:
                    st.markdown(f'<div class="clean-group-title">{title_header}</div>', unsafe_allow_html=True)

                def filter_data(df_in):
                    if df_in.empty: return df_in
                    return df_in[(df_in["Bucket"] == b_filt) & (df_in["Project"] == p_filt)]

                has_any_data = any(
                    not filter_data(weekly_processed.get(p, pd.DataFrame())).empty
                    for p in available_periods
                )
                if not has_any_data:
                    st.info(f"Tidak ada data kompetisi untuk divisi {p_filt} - {b_filt}.")
                    return

                df_monthly_base = filter_data(weekly_processed.get("Monthly", pd.DataFrame()))

                # ROW 1 DATA: PERFORMA GABUNGAN PERUSAHAAN
                comp_summary_df = pd.DataFrame()
                top_mpt_rank, top_mpt_rate, gap_to_top = "-", "-", "-"

                if not df_monthly_base.empty:
                    comp_rows = []
                    for comp_name, grp in df_monthly_base.groupby("Perusahaan"):
                        pen_file = grp["Total Penambahan"].sum()
                        rec_file = grp["Total Penerimaan"].sum()

                        d_info = daily_comp_map.get((p_filt, b_filt, comp_name), {"pen": 0.0, "rec": 0.0})
                        pen_total = pen_file + d_info["pen"]
                        rec_total = rec_file + d_info["rec"]
                        rate_total = (rec_total / pen_total * 100) if pen_total > 0 else 0.0

                        comp_rows.append({
                            "Perusahaan": comp_name,
                            "_pen_num": pen_total,
                            "Total Penambahan (Monthly)": format_num(pen_total, is_currency=True),
                            "Total Penerimaan (Monthly)": format_num(rec_total, is_currency=True),
                            "_rate_num": rate_total,
                            "Recovery Rate (Monthly)": f"{rate_total:.2f}%",
                        })

                    df_comp_temp = pd.DataFrame(comp_rows)
                    if not df_comp_temp.empty:
                        df_comp_temp = df_comp_temp.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
                        top_rate = df_comp_temp.loc[0, "_rate_num"]
                        gaps_plus_one, gaps_to_first = ["-"], ["-"]

                        for i in range(1, len(df_comp_temp)):
                            pen_cur = df_comp_temp.loc[i, "_pen_num"]
                            rate_cur = df_comp_temp.loc[i, "_rate_num"]

                            diff_1 = df_comp_temp.loc[i - 1, "_rate_num"] - rate_cur
                            gaps_plus_one.append(f"-{format_num((diff_1 / 100.0) * pen_cur, is_currency=True)}" if diff_1 > 0 and pen_cur > 0 else "-")

                            diff_top = top_rate - rate_cur
                            gaps_to_first.append(f"-{format_num((diff_top / 100.0) * pen_cur, is_currency=True)}" if diff_top > 0 and pen_cur > 0 else "-")

                        df_comp_temp["Rank"] = [f"#{i+1}" for i in range(len(df_comp_temp))]
                        df_comp_temp["Gap (+1 Rank)"] = gaps_plus_one
                        df_comp_temp["Gap (Rank #1)"] = gaps_to_first
                        comp_summary_df = df_comp_temp

                        mpt_match = df_comp_temp[df_comp_temp["Perusahaan"].str.contains("MPT05")]
                        if not mpt_match.empty:
                            top_mpt_rank = mpt_match.iloc[0]["Rank"]
                            top_mpt_rate = mpt_match.iloc[0]["Recovery Rate (Monthly)"]
                            gap_to_top = mpt_match.iloc[0]["Gap (Rank #1)"]

                card_html = f"""
                <div class="exec-card-container">
                    <div class="exec-card" style="border-left: 4px solid #2563eb;">
                        <div class="exec-card-title">Ranking MPT05</div>
                        <div class="exec-card-value" style="color: #2563eb;">{top_mpt_rank}</div>
                    </div>
                    <div class="exec-card" style="border-left: 4px solid #10b981;">
                        <div class="exec-card-title">Recovery Rate MPT05</div>
                        <div class="exec-card-value" style="color: #059669;">{top_mpt_rate}</div>
                    </div>
                    <div class="exec-card" style="border-left: 4px solid #ef4444;">
                        <div class="exec-card-title">Gap Amount Menuju Juara #1</div>
                        <div class="exec-card-value" style="color: #dc2626; font-size: 1.15rem;">{gap_to_top}</div>
                    </div>
                </div>
                """
                st.markdown(card_html, unsafe_allow_html=True)

                # ROW 1 TAMPILAN: PERFORMA GABUNGAN PERUSAHAAN (FULL LEBAR)
                st.markdown('<div class="section-banner-title">机构月度累计绩效 (Performa Gabungan Perusahaan)</div>', unsafe_allow_html=True)
                if not comp_summary_df.empty:
                    cols_c = ["Rank", "Perusahaan", "Total Penambahan (Monthly)", "Total Penerimaan (Monthly)", "Recovery Rate (Monthly)", "Gap (+1 Rank)", "Gap (Rank #1)"]
                    display_full_table(translate_columns(comp_summary_df[cols_c]))
                else:
                    st.caption("Tidak ada data.")

                st.write("")

                # ROW 2 DATA: PERFORMA TL MONTHLY & DAILY
                df_tl_monthly = pd.DataFrame()
                if not df_monthly_base.empty:
                    tl_m_rows = []
                    for tl_name, grp in df_monthly_base.groupby("Nama TL"):
                        c_name = grp.iloc[0].get("Perusahaan", "")
                        pen_file = grp["Total Penambahan"].sum()
                        rec_file = grp["Total Penerimaan"].sum()

                        d_tl = daily_tl_map.get((p_filt, b_filt, tl_name), {"pen": 0.0, "rec": 0.0})
                        pen_total = pen_file + d_tl["pen"]
                        rec_total = rec_file + d_tl["rec"]
                        rate_total = (rec_total / pen_total * 100) if pen_total > 0 else 0.0

                        tl_m_rows.append({
                            "Nama TL": tl_name,
                            "Perusahaan": c_name,
                            "_pen_num": pen_total,
                            "Total Penambahan (Monthly)": format_num(pen_total, is_currency=True),
                            "Total Penerimaan (Monthly)": format_num(rec_total, is_currency=True),
                            "_rate_num": rate_total,
                            "Recovery Rate (Monthly)": f"{rate_total:.2f}%",
                        })

                    df_tl_m_temp = pd.DataFrame(tl_m_rows)
                    if not df_tl_m_temp.empty:
                        df_tl_m_temp = df_tl_m_temp.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
                        gaps_pct_m, gaps_amt_m = ["-"], ["-"]
                        for i in range(1, len(df_tl_m_temp)):
                            diff = df_tl_m_temp.loc[i - 1, "_rate_num"] - df_tl_m_temp.loc[i, "_rate_num"]
                            pen_c = df_tl_m_temp.loc[i, "_pen_num"]
                            if diff > 0 and pen_c > 0:
                                gaps_pct_m.append(f"+{diff:.2f}%")
                                gaps_amt_m.append(f"-{format_num((diff / 100.0) * pen_c, is_currency=True)}")
                            else:
                                gaps_pct_m.append("0.00%")
                                gaps_amt_m.append("-")

                        df_tl_m_temp["Rank"] = [f"#{i+1}" for i in range(len(df_tl_m_temp))]
                        df_tl_m_temp["Gap"] = gaps_pct_m
                        df_tl_m_temp["Gap Amount"] = gaps_amt_m
                        cols_tl_m = ["Rank", "Nama TL", "Perusahaan", "Total Penambahan (Monthly)", "Total Penerimaan (Monthly)", "Recovery Rate (Monthly)", "Gap", "Gap Amount"]
                        df_tl_monthly = df_tl_m_temp[cols_tl_m]

                df_tl_daily = pd.DataFrame()
                if not df_monthly_base.empty:
                    tl_d_rows = []
                    for tl_name, grp in df_monthly_base.groupby("Nama TL"):
                        d_tl = daily_tl_map.get((p_filt, b_filt, tl_name), {"pen": 0.0, "rec": 0.0})
                        d_pen = d_tl["pen"]
                        d_rec = d_tl["rec"]
                        d_rate = (d_rec / d_pen * 100) if d_pen > 0 else 0.0

                        if d_pen > 0:
                            tl_d_rows.append({
                                "Nama TL": tl_name,
                                "_daily_pen": d_pen,
                                "Penambahan (Hari Ini)": format_num(d_pen, is_currency=True),
                                "Penerimaan (Hari Ini)": format_num(d_rec, is_currency=True),
                                "_rate_num": d_rate,
                                "Recovery Rate (Hari Ini)": f"{d_rate:.2f}%",
                            })

                    df_tl_d_temp = pd.DataFrame(tl_d_rows)
                    if not df_tl_d_temp.empty:
                        df_tl_d_temp = df_tl_d_temp.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
                        gaps_pct_d, gaps_amt_d = ["-"], ["-"]
                        for i in range(1, len(df_tl_d_temp)):
                            diff_d = df_tl_d_temp.loc[i - 1, "_rate_num"] - df_tl_d_temp.loc[i, "_rate_num"]
                            pen_dc = df_tl_d_temp.loc[i, "_daily_pen"]
                            if diff_d > 0 and pen_dc > 0:
                                gaps_pct_d.append(f"+{diff_d:.2f}%")
                                gaps_amt_d.append(f"-{format_num((diff_d / 100.0) * pen_dc, is_currency=True)}")
                            else:
                                gaps_pct_d.append("0.00%")
                                gaps_amt_d.append("-")

                        df_tl_d_temp["Rank"] = [f"#{i+1}" for i in range(len(df_tl_d_temp))]
                        df_tl_d_temp["Gap"] = gaps_pct_d
                        df_tl_d_temp["Gap Amount"] = gaps_amt_d
                        cols_tl_d = ["Rank", "Nama TL", "Penambahan (Hari Ini)", "Penerimaan (Hari Ini)", "Recovery Rate (Hari Ini)", "Gap", "Gap Amount"]
                        df_tl_daily = df_tl_d_temp[cols_tl_d]

                # ROW 2 TAMPILAN: PERFORMA TL MONTHLY | PERFORMA TL DAILY (2 KOLOM SEJAJAR)
                col_r2_left, col_r2_right = st.columns([1, 1], gap="small")

                with col_r2_left:
                    st.markdown('<div class="section-banner-title tl">组长月度累计绩效 (Performa TL Monthly)</div>', unsafe_allow_html=True)
                    if not df_tl_monthly.empty:
                        display_full_table(translate_columns(df_tl_monthly))
                    else:
                        st.caption("Tidak ada data.")

                with col_r2_right:
                    st.markdown('<div class="section-banner-title daily">组长今日实时绩效 (Performa TL Daily)</div>', unsafe_allow_html=True)
                    if not df_tl_daily.empty:
                        display_full_table(translate_columns(df_tl_daily))
                    else:
                        st.caption("Belum ada data setoran harian hari ini.")

                st.write("")

                # ROW 3 TAMPILAN: KHUSUS MINGGUAN (WEEKLY 1, 2, 3, 4)
                weekly_only_periods = [p for p in available_periods if p != "Monthly"]
                if weekly_only_periods:
                    st.markdown('<div class="section-banner-title">组长周度绩效对比 (Performa Weekly 1, 2, 3, 4)</div>', unsafe_allow_html=True)
                    cols_w_layout = st.columns(len(weekly_only_periods))

                    active_running_w = weekly_only_periods[-1]

                    for idx_w, w_label in enumerate(weekly_only_periods):
                        with cols_w_layout[idx_w]:
                            st.markdown(f'<div class="month-header-clean">{w_label}</div>', unsafe_allow_html=True)
                            df_w_data = filter_data(weekly_processed.get(w_label, pd.DataFrame()))

                            if df_w_data.empty:
                                st.caption("Tidak ada data (暂无数据)")
                            else:
                                is_running = (w_label == active_running_w)
                                tl_w_rows = []
                                for tl_name, grp in df_w_data.groupby("Nama TL"):
                                    c_name = grp.iloc[0].get("Perusahaan", "")
                                    pen_val = grp["Total Penambahan"].sum()
                                    rec_val = grp["Total Penerimaan"].sum()

                                    if is_running:
                                        d_tl = daily_tl_map.get((p_filt, b_filt, tl_name), {"pen": 0.0, "rec": 0.0})
                                        pen_val += d_tl["pen"]
                                        rec_val += d_tl["rec"]

                                    rate_val = (rec_val / pen_val * 100) if pen_val > 0 else 0.0

                                    tl_w_rows.append({
                                        "Nama TL": tl_name,
                                        "Perusahaan": c_name,
                                        "_rate_num": rate_val,
                                        "Recovery Rate": f"{rate_val:.2f}%",
                                    })

                                df_tl_w_card = pd.DataFrame(tl_w_rows)
                                if not df_tl_w_card.empty:
                                    df_tl_w_card = df_tl_w_card.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
                                    df_tl_w_card["Rank"] = [f"#{i+1}" for i in range(len(df_tl_w_card))]
                                    cols_w_view = ["Rank", "Nama TL", "Perusahaan", "Recovery Rate"]
                                    display_full_table(translate_columns(df_tl_w_card[cols_w_view]))

                st.write("")

            for s_tab, (s_label, b_filt, p_filt) in zip(sum_sub_tabs, TARGET_BUCKETS):
                with s_tab:
                    if b_filt and p_filt:
                        render_summary_three_rows_block(b_filt, p_filt)
                    else:
                        actual_divisions = [
                            ("S1 - SK", "S1", "SK"),
                            ("D0 - SK", "D0", "SK"),
                            ("S1 - LD", "S1", "LD"),
                            ("S1 - LD-BA", "S1", "LD-BA"),
                            ("S0 - LD", "S0", "LD"),
                            ("C - LD-C", "C", "LD-C"),
                        ]
                        for div_label, b_div, p_div in actual_divisions:
                            render_summary_three_rows_block(
                                b_div, 
                                p_div, 
                                title_header=f"Project (项目): {p_div} &nbsp;|&nbsp; Bucket (阶段): {b_div}"
                            )

# =========================================================
# MENU 2: TEAM PERFORMANCE (SEMUA BUCKET TANPA SUB-TAB)
# =========================================================
with tab_team_perf:
    st.caption("Team Performance (团队绩效排行榜) — Performa Bulanan & Harian Seluruh Bucket Tanpa Pemisahan Tab")

    df_monthly_team_all = weekly_processed.get("Monthly", pd.DataFrame()) if "weekly_processed" in locals() and weekly_processed else pd.DataFrame()
    if df_monthly_team_all.empty and "Monthly" in weekly_dict:
        df_monthly_team_all = process_department_weekly_data(weekly_dict["Monthly"])

    if df_monthly_team_all.empty:
        st.info("Belum ada data bulanan di folder `rawdata/Bulanan/`.")
    else:
        actual_divisions = [
            ("S1 - SK", "S1", "SK"),
            ("D0 - SK", "D0", "SK"),
            ("S1 - LD", "S1", "LD"),
            ("S1 - LD-BA", "S1", "LD-BA"),
            ("S0 - LD", "S0", "LD"),
            ("C - LD-C", "C", "LD-C"),
        ]

        for div_label, b_code, p_code in actual_divisions:
            df_div_base = df_monthly_team_all[
                (df_monthly_team_all["Bucket"] == b_code) & 
                (df_monthly_team_all["Project"] == p_code)
            ]
            if df_div_base.empty:
                continue

            st.markdown(f'<div class="clean-group-title">Project: {p_code} &nbsp;|&nbsp; Bucket: {b_code}</div>', unsafe_allow_html=True)

            # 1. Tabel Bulanan Team
            tl_m_rows = []
            for tl_name, grp in df_div_base.groupby("Nama TL"):
                c_name = grp.iloc[0].get("Perusahaan", "")
                pen_file = grp["Total Penambahan"].sum()
                rec_file = grp["Total Penerimaan"].sum()

                d_tl = daily_tl_map.get((p_code, b_code, tl_name), {"pen": 0.0, "rec": 0.0})
                pen_total = pen_file + d_tl["pen"]
                rec_total = rec_file + d_tl["rec"]
                rate_total = (rec_total / pen_total * 100) if pen_total > 0 else 0.0

                tl_m_rows.append({
                    "Project": p_code,
                    "Bucket": b_code,
                    "Nama TL": tl_name,
                    "Perusahaan": c_name,
                    "_pen_num": pen_total,
                    "Total Penambahan (Monthly)": format_num(pen_total, is_currency=True),
                    "Total Penerimaan (Monthly)": format_num(rec_total, is_currency=True),
                    "_rate_num": rate_total,
                    "Recovery Rate (Monthly)": f"{rate_total:.2f}%",
                })

            df_tl_m = pd.DataFrame(tl_m_rows)
            if not df_tl_m.empty:
                df_tl_m = df_tl_m.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
                gaps_pct_m, gaps_amt_m = ["-"], ["-"]
                for i in range(1, len(df_tl_m)):
                    diff = df_tl_m.loc[i - 1, "_rate_num"] - df_tl_m.loc[i, "_rate_num"]
                    pen_c = df_tl_m.loc[i, "_pen_num"]
                    if diff > 0 and pen_c > 0:
                        gaps_pct_m.append(f"+{diff:.2f}%")
                        gaps_amt_m.append(f"-{format_num((diff / 100.0) * pen_c, is_currency=True)}")
                    else:
                        gaps_pct_m.append("0.00%")
                        gaps_amt_m.append("-")

                df_tl_m["Rank"] = [f"#{i+1}" for i in range(len(df_tl_m))]
                df_tl_m["Gap"] = gaps_pct_m
                df_tl_m["Gap Amount"] = gaps_amt_m
                cols_tl_m = ["Rank", "Project", "Bucket", "Nama TL", "Perusahaan", "Total Penambahan (Monthly)", "Total Penerimaan (Monthly)", "Recovery Rate (Monthly)", "Gap", "Gap Amount"]
                df_tl_m = df_tl_m[cols_tl_m]

            # 2. Tabel Harian Team
            tl_d_rows = []
            for tl_name, grp in df_div_base.groupby("Nama TL"):
                d_tl = daily_tl_map.get((p_code, b_code, tl_name), {"pen": 0.0, "rec": 0.0})
                d_pen = d_tl["pen"]
                d_rec = d_tl["rec"]
                d_rate = (d_rec / d_pen * 100) if d_pen > 0 else 0.0

                if d_pen > 0:
                    tl_d_rows.append({
                        "Nama TL": tl_name,
                        "_daily_pen": d_pen,
                        "Penambahan (Hari Ini)": format_num(d_pen, is_currency=True),
                        "Penerimaan (Hari Ini)": format_num(d_rec, is_currency=True),
                        "_rate_num": d_rate,
                        "Recovery Rate (Hari Ini)": f"{d_rate:.2f}%",
                    })

            df_tl_d = pd.DataFrame(tl_d_rows)
            if not df_tl_d.empty:
                df_tl_d = df_tl_d.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
                gaps_pct_d, gaps_amt_d = ["-"], ["-"]
                for i in range(1, len(df_tl_d)):
                    diff_d = df_tl_d.loc[i - 1, "_rate_num"] - df_tl_d.loc[i, "_rate_num"]
                    pen_dc = df_tl_d.loc[i, "_daily_pen"]
                    if diff_d > 0 and pen_dc > 0:
                        gaps_pct_d.append(f"+{diff_d:.2f}%")
                        gaps_amt_d.append(f"-{format_num((diff_d / 100.0) * pen_dc, is_currency=True)}")
                    else:
                        gaps_pct_d.append("0.00%")
                        gaps_amt_d.append("-")

                df_tl_d["Rank"] = [f"#{i+1}" for i in range(len(df_tl_d))]
                df_tl_d["Gap"] = gaps_pct_d
                df_tl_d["Gap Amount"] = gaps_amt_d
                cols_tl_d = ["Rank", "Nama TL", "Penambahan (Hari Ini)", "Penerimaan (Hari Ini)", "Recovery Rate (Hari Ini)", "Gap", "Gap Amount"]
                df_tl_d = df_tl_d[cols_tl_d]

            # Render Bersebelahan
            c_left, c_right = st.columns([1.15, 1], gap="small")
            with c_left:
                st.markdown('<div class="section-banner-title">Performa Bulanan (月度累计绩效)</div>', unsafe_allow_html=True)
                if not df_tl_m.empty:
                    display_full_table(translate_columns(df_tl_m))
                else:
                    st.caption("Tidak ada data.")

            with c_right:
                st.markdown('<div class="section-banner-title tl">Performa Harian (今日实时绩效)</div>', unsafe_allow_html=True)
                if not df_tl_d.empty:
                    display_full_table(translate_columns(df_tl_d))
                else:
                    st.caption("Belum ada data setoran harian hari ini.")

            st.write("")

# =========================================================
# MENU 3: PERFORMANCE AGENT (>7 HARI KERJA)
# (Ranking Interval: TOP 10%, 10-20%, 20-30%, 30-50%, 50-70%, 70-80%, BOTTOM 20%)
# =========================================================
with tab_agent_senior:
    st.caption("Performance Agent (>7 Hari Kerja) — Akumulasi MTD + Daily, Berkompetisi Penuh (出勤天数 > 7天坐席排行榜)")
    if df_agents_processed.empty:
        st.info("Belum ada data agen.")
    else:
        df_senior = df_agents_processed[
            (df_agents_processed["Apakah Ikut Berkompetisi"] == "Ya") & 
            (df_agents_processed["Hari Kerja Angka"] > 7)
        ].copy()
        render_agent_bucket_tabs(df_senior)

# =========================================================
# MENU 4: PERFORMANCE AGENT (<=7 HARI KERJA)
# (Ranking Interval: TOP 10%, 10-20%, 20-30%, 30-50%, 50-70%, 70-80%, BOTTOM 20%)
# =========================================================
with tab_agent_junior:
    st.caption("Performance Agent (<=7 Hari Kerja) — Masa Percobaan / Agen Baru (出勤天数 <= 7天坐席排行榜)")
    if df_agents_processed.empty:
        st.info("Belum ada data agen.")
    else:
        df_junior = df_agents_processed[
            (df_agents_processed["Hari Kerja Angka"] <= 7)
        ].copy()
        render_agent_bucket_tabs(df_junior)

# =========================================================
# MENU 5: KOMPOSISI AGENT (1 ROW TERPADU: TOP 30%, MID, BOT 30%)
# =========================================================
with tab_agent_comp:
    st.markdown("### 👥 **Komposisi Agent per Team (Khusus Tim Berkompetisi & Agent Aktif)**")
    st.caption("Hanya menghitung agen AKTIF. Menampilkan persentase Baru vs Lama dan sebaran ranking interval perorangan (Top 30%, Middle 30-70%, Bottom 30%).")

    df_comp_active_agents = df_agents_processed[
        (df_agents_processed["Apakah Ikut Berkompetisi"] == "Ya") &
        (df_agents_processed["Status Kerja"] == "Aktif")
    ].copy()

    if df_comp_active_agents.empty:
        st.info("Tidak ada data agen aktif yang terdaftar di Master Kompetisi.")
    else:
        comp_sub_tabs = st.tabs([label for label, _, _ in TARGET_BUCKETS])

        def render_consolidated_composition_block(df_target_comp, title_block=None):
            if title_block:
                st.markdown(f'<div class="clean-group-title">{title_block}</div>', unsafe_allow_html=True)

            if df_target_comp.empty:
                st.info("Tidak ada data tim kompetisi aktif pada divisi ini.")
                return

            rows_consolidated = []
            for (p_c, b_c, tl_c), grp in df_target_comp.groupby(["Project", "Bucket", "Nama TL"]):
                tot_ag = len(grp)
                grp_baru = grp[grp["Kategori Agent"] == "Agent Baru"]
                grp_lama = grp[grp["Kategori Agent"] == "Agent Lama"]

                tot_baru = len(grp_baru)
                tot_lama = len(grp_lama)

                pct_baru = (tot_baru / tot_ag * 100.0) if tot_ag > 0 else 0.0
                pct_lama = (tot_lama / tot_ag * 100.0) if tot_ag > 0 else 0.0
                c_name = grp.iloc[0].get("Perusahaan", "")

                # Sebaran 3 Interval Khusus Komposisi (Ranking Interval 3Cat)
                baru_top30 = len(grp_baru[grp_baru["Ranking Interval 3Cat"] == "Top 30%"])
                baru_mid = len(grp_baru[grp_baru["Ranking Interval 3Cat"] == "Middle 30-70%"])
                baru_bot30 = len(grp_baru[grp_baru["Ranking Interval 3Cat"] == "Bottom 30%"])

                lama_top30 = len(grp_lama[grp_lama["Ranking Interval 3Cat"] == "Top 30%"])
                lama_mid = len(grp_lama[grp_lama["Ranking Interval 3Cat"] == "Middle 30-70%"])
                lama_bot30 = len(grp_lama[grp_lama["Ranking Interval 3Cat"] == "Bottom 30%"])

                rows_consolidated.append({
                    "Project": p_c,
                    "Bucket": b_c,
                    "Nama TL": tl_c,
                    "Perusahaan": c_name,
                    "Total Agent": tot_ag,
                    "% Agent Baru": f"{pct_baru:.1f}%",
                    "% Agent Lama": f"{pct_lama:.1f}%",
                    "Baru: Top 30%": baru_top30,
                    "Baru: Middle 30-70%": baru_mid,
                    "Baru: Bottom 30%": baru_bot30,
                    "Lama: Top 30%": lama_top30,
                    "Lama: Middle 30-70%": lama_mid,
                    "Lama: Bottom 30%": lama_bot30,
                })

            df_consolidated = pd.DataFrame(rows_consolidated)
            if not df_consolidated.empty:
                df_consolidated = df_consolidated.sort_values(by=["Project", "Bucket", "Nama TL"]).reset_index(drop=True)
                df_consolidated.insert(0, "No", range(1, len(df_consolidated) + 1))
                display_full_table(translate_columns(df_consolidated))

        for tab_c, (tab_label, b_filter, p_filter) in zip(comp_sub_tabs, TARGET_BUCKETS):
            with tab_c:
                if b_filter and p_filter:
                    view_c = df_comp_active_agents[
                        (df_comp_active_agents["Bucket"] == b_filter) & 
                        (df_comp_active_agents["Project"] == p_filter)
                    ]
                    render_consolidated_composition_block(view_c)
                else:
                    actual_divisions = [
                        ("S1 - SK", "S1", "SK"),
                        ("D0 - SK", "D0", "SK"),
                        ("S1 - LD", "S1", "LD"),
                        ("S1 - LD-BA", "S1", "LD-BA"),
                        ("S0 - LD", "S0", "LD"),
                        ("C - LD-C", "C", "LD-C"),
                    ]
                    for div_lbl, b_c, p_c in actual_divisions:
                        sub_view = df_comp_active_agents[
                            (df_comp_active_agents["Bucket"] == b_c) & 
                            (df_comp_active_agents["Project"] == p_c)
                        ]
                        if not sub_view.empty:
                            render_consolidated_composition_block(sub_view, f"Project: {p_c} &nbsp;|&nbsp; Bucket: {b_c} ({len(sub_view)} Agent Aktif)")

        with st.expander("🔍 **Detail Perorangan Agent Aktif (Hiredate, Kategori & Interval Ranking)**", expanded=False):
            cols_ag_detail = [
                "Rank MTD", "Ranking Interval", "Project", "Bucket", "Nama TL",
                "ID Agent", "Nama Pengguna", "Hiredate", "Kategori Agent",
                "Total Penambahan Hari Ini_Total", "Total Penerimaan Hari Ini_Total", "Recovery Rate Num"
            ]
            df_d_view = df_comp_active_agents[cols_ag_detail].copy()
            df_d_view = df_d_view.rename(columns={
                "Rank MTD": "Rank",
                "Total Penambahan Hari Ini_Total": "Total Penambahan (Monthly)",
                "Total Penerimaan Hari Ini_Total": "Total Penerimaan (Monthly)",
                "Recovery Rate Num": "Recovery Rate (%)"
            })
            df_d_view["Total Penambahan (Monthly)"] = df_d_view["Total Penambahan (Monthly)"].apply(lambda v: format_num(v, is_currency=True))
            df_d_view["Total Penerimaan (Monthly)"] = df_d_view["Total Penerimaan (Monthly)"].apply(lambda v: format_num(v, is_currency=True))
            df_d_view["Recovery Rate (%)"] = df_d_view["Recovery Rate (%)"].apply(lambda v: f"{v:.2f}%")
            display_full_table(translate_columns(df_d_view))

# =========================================================
# MENU 6: SCHEDULE TRAINING TEAM LEADER (KHUSUS MPT05)
# =========================================================
with tab_training:
    st.markdown("### 📅 **Schedule Training Team Leader ke Agent (Khusus MPT05)**")
    st.caption("Jadwal dan topik pembahasan training internal MPT05. Dilengkapi proteksi password untuk edit/tambah jadwal.")

    def load_training_data():
        if os.path.exists(TRAINING_FILE):
            try:
                with open(TRAINING_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return []
        return []

    def save_training_data(data_list):
        try:
            with open(TRAINING_FILE, "w", encoding="utf-8") as f:
                json.dump(data_list, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

    training_records = load_training_data()

    mpt05_tls_list = [
        "MPT05-TL-001 (S1 - SK)",
        "MPT05-TL-002 (S1 - SK)",
        "MPT05-TL-003 (S1 - SK)",
        "MPT05-TL-004 (S1 - SK)",
        "MPT05-D0-TL-001 (D0 - SK)",
        "MPT05-D0-TL-002 (D0 - SK)",
        "W-MPT05 (S1 - LD)",
        "W-MPT05-2 (S1 - LD)",
        "W-MPT05-3 (S1 - LD)",
        "W-MPT05-S0-01 (S0 - LD)",
        "W-MPT05-BA (S1 - LD-BA)",
        "W-MPT05-C (C - LD-C)",
    ]

    TRAINING_ADMIN_PASSWORD = "mpt05training"

    if "training_edit_unlocked" not in st.session_state:
        st.session_state["training_edit_unlocked"] = False

    with st.expander("🔒 **Form Input / Edit Jadwal Training (Perlu Password)**", expanded=False):
        c_pwd, c_btn, _ = st.columns([2, 1.2, 3])
        with c_pwd:
            input_tr_pwd = st.text_input("Password TL / Admin:", type="password", key="input_tr_pwd")
        with c_btn:
            st.write("")
            if not st.session_state["training_edit_unlocked"]:
                if st.button("🔓 Buka Akses Edit", use_container_width=True, key="btn_unlock_tr"):
                    if input_tr_pwd in [TRAINING_ADMIN_PASSWORD, "mpt05admin"]:
                        st.session_state["training_edit_unlocked"] = True
                        st.success("Akses edit jadwal terbuka!")
                        st.rerun()
                    else:
                        st.error("Password salah!")
            else:
                if st.button("🔒 Kunci Akses", use_container_width=True, key="btn_lock_tr"):
                    st.session_state["training_edit_unlocked"] = False
                    st.rerun()

        if st.session_state["training_edit_unlocked"]:
            st.markdown("---")
            st.markdown("##### ✏️ **Tambah / Update Jadwal Training:**")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                sel_tl = st.selectbox("Pilih Team Leader:", mpt05_tls_list, key="sel_tr_tl")
                in_jadwal = st.text_input("Jadwal Training (Hari & Jam):", placeholder="Contoh: Setiap Rabu, 09:00 - 10:30 WIB", key="in_tr_jadwal")
            with col_f2:
                in_link = st.text_input("Link Gmeet / Zoom Meeting:", placeholder="https://meet.google.com/xxx-xxxx-xxx", key="in_tr_link")
                in_topik = st.text_area("Topic Pembahasan:", placeholder="Contoh: Teknik Negosiasi Nasabah Janji Bayar & Closing Penagihan", key="in_tr_topik", height=70)

            col_sub1, col_sub2 = st.columns([1.5, 4])
            with col_sub1:
                if st.button("💾 Simpan Jadwal Training", use_container_width=True, key="btn_save_tr"):
                    if not in_jadwal or not in_topik:
                        st.warning("Mohon lengkapi Jadwal dan Topic Pembahasan!")
                    else:
                        tl_name_clean = sel_tl.split(" (")[0]
                        existing_idx = next((i for i, r in enumerate(training_records) if r.get("Nama Team Leader") == tl_name_clean), -1)

                        new_entry = {
                            "Nama Team Leader": tl_name_clean,
                            "Divisi": sel_tl.split(" (")[1].replace(")", "") if " (" in sel_tl else "-",
                            "jadwal_training": in_jadwal,
                            "topic_pembahasan": in_topik,
                            "link_meeting": in_link if in_link else "-",
                            "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M")
                        }

                        if existing_idx >= 0:
                            training_records[existing_idx] = new_entry
                        else:
                            training_records.append(new_entry)

                        save_training_data(training_records)
                        st.success(f"Jadwal training untuk {tl_name_clean} berhasil disimpan!")
                        st.rerun()

    # Tampilkan Tabel Jadwal Training
    records_map = {r.get("Nama Team Leader"): r for r in training_records}
    final_tr_rows = []
    for tl_str in mpt05_tls_list:
        tl_n = tl_str.split(" (")[0]
        if tl_n in records_map:
            r_item = records_map[tl_n]
            final_tr_rows.append({
                "Nama Team Leader": tl_n,
                "jadwal_training": r_item.get("jadwal_training", "-"),
                "topic_pembahasan": r_item.get("topic_pembahasan", "-"),
                "link_meeting": r_item.get("link_meeting", "-"),
            })
        else:
            final_tr_rows.append({
                "Nama Team Leader": tl_n,
                "jadwal_training": "Belum Dijadwalkan",
                "topic_pembahasan": "Belum Diisi oleh TL",
                "link_meeting": "-",
            })

    df_tr_view = pd.DataFrame(final_tr_rows)
    df_tr_view.insert(0, "No", range(1, len(df_tr_view) + 1))
    display_full_table(translate_columns(df_tr_view))

# =========================================================
# MENU 7: PRODUCTIVITY AGENTS (MPT05 DAILY TARGET MONITOR)
# =========================================================
with tab_productivity:
    st.markdown("### 🎯 **Produktivitas Harian Agen MPT05 (Agent Daily Productivity)**")
    st.caption("Monitoring performa operasional dan pencapaian target harian agen MPT05 dari data daily terbaru.")

    TARGET_KEYS = ["D0 - SK", "S1 - SK", "S1 - LD", "S1 - LD-BA", "S0 - LD", "C - LD-C"]
    ADMIN_TARGET_PASSWORD = "mpt05admin"

    if "target_unlocked" not in st.session_state:
        st.session_state["target_unlocked"] = False

    with st.expander("🔒 **Pengaturan Target per Divisi / Bucket (Perlu Password untuk Mengubah)**", expanded=False):
        col_pwd, col_btn_lock, _ = st.columns([2, 1.2, 3])
        with col_pwd:
            input_pwd = st.text_input("Masukkan Password Admin:", type="password", key="input_target_pwd_prod")
        with col_btn_lock:
            st.write("")
            if not st.session_state["target_unlocked"]:
                if st.button("🔓 Buka Kunci Edit", use_container_width=True, key="btn_unlock_prod"):
                    if input_pwd == ADMIN_TARGET_PASSWORD:
                        st.session_state["target_unlocked"] = True
                        st.success("Akses edit terbuka!")
                        st.rerun()
                    else:
                        st.error("Password salah!")
            else:
                if st.button("🔒 Kunci Kembali", use_container_width=True, key="btn_lock_prod"):
                    st.session_state["target_unlocked"] = False
                    st.info("Pengaturan dikunci kembali.")
                    st.rerun()

        is_editable = st.session_state["target_unlocked"]
        target_configs = {}
        for b_name in TARGET_KEYS:
            st.markdown(f"**Divisi: `{b_name}`**")
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                t_call = st.number_input("Target Call", min_value=0, value=300, step=10, key=f"t_call_{b_name}", disabled=not is_editable)
            with c2:
                t_work = st.number_input("Waktu Kerja (Jam)", min_value=0.0, value=8.0, step=0.5, key=f"t_work_{b_name}", disabled=not is_editable)
            with c3:
                t_sms = st.number_input("Target SMS", min_value=0, value=75, step=5, key=f"t_sms_{b_name}", disabled=not is_editable)
            with c4:
                t_waba = st.number_input("Target WABA", min_value=0, value=0, step=5, key=f"t_waba_{b_name}", disabled=not is_editable)
            with c5:
                t_rec = st.number_input("Target Rec Rate (%)", min_value=0.0, max_value=100.0, value=15.0, step=1.0, key=f"t_rec_{b_name}", disabled=not is_editable)

            target_configs[b_name] = {
                "call": t_call,
                "work": t_work,
                "sms": t_sms,
                "waba": t_waba,
                "rec_rate": t_rec,
            }
            st.write("")

    if not daily_files:
        st.warning("Belum ada file data harian (daily) di folder rawdata.")
    else:
        daily_dfs_prod = []
        for f in daily_files:
            try:
                df_temp = pd.read_excel(f)
                df_temp.columns = [str(c).strip() for c in df_temp.columns]
                fname = os.path.basename(f).upper()
                df_temp["_source_project"] = "SK" if "SK" in fname else "LD"
                daily_dfs_prod.append(df_temp)
            except Exception:
                pass

        if not daily_dfs_prod:
            st.info("Tidak dapat memuat berkas daily.")
        else:
            df_daily_raw = pd.concat(daily_dfs_prod, ignore_index=True)
            agent_productivity_list = []
            for _, r in df_daily_raw.iterrows():
                id_raw = str(r.get("Nama Pengguna", "")).strip()
                nm_raw = str(r.get("Nama", "")).strip()
                proj_src = r.get("_source_project", "")
                p_cat = "SK" if "SK" in proj_src else "LD"

                id_clean = normalize_clean(id_raw)
                nm_clean = normalize_clean(nm_raw)

                k_info = karyawan_dict.get((id_clean, p_cat)) or karyawan_dict.get((nm_clean, p_cat)) or karyawan_dict.get((id_clean, proj_src)) or {}
                tl_name = k_info.get("Nama TL", "Unknown")
                bucket_name = k_info.get("Bucket", "Unknown")
                proj_name = k_info.get("Project", proj_src)
                perusahaan_name = k_info.get("Perusahaan", "")

                is_mpt05 = ("MPT05" in tl_name.upper()) or ("MPT05" in id_clean) or ("MPT05" in perusahaan_name.upper())
                if not is_mpt05:
                    continue

                pen_hari_ini = clean_numeric(r.get("Total Penambahan Hari Ini", 0))
                rec_hari_ini = clean_numeric(r.get("Total Penerimaan Hari Ini", 0))
                rec_rate = (rec_hari_ini / pen_hari_ini * 100.0) if pen_hari_ini > 0 else 0.0

                tot_data_today = clean_numeric(r.get("Jumlah total data hari ini", 0))
                jml_calls = clean_numeric(r.get("Jumlah Panggilan", 0))
                jml_calls_success = clean_numeric(r.get("Jumlah Panggilan Berhasil", 0))
                connected_rate = (jml_calls_success / jml_calls * 100.0) if jml_calls > 0 else 0.0

                work_time = clean_numeric(r.get("Waktu Kerja Rata-rata", 0))
                jml_sms = clean_numeric(r.get("Jumlah SMS", 0))
                waba_sent = clean_numeric(r.get("Jumlah Pengiriman Template WABA", 0))

                agent_productivity_list.append({
                    "Project": proj_name,
                    "Bucket": bucket_name,
                    "Nama TL": tl_name,
                    "ID Agent": k_info.get("ID_Akun_Resmi", id_raw),
                    "Nama Agent": k_info.get("Nama_Orang_Resmi", nm_raw),
                    "Total_Data": tot_data_today,
                    "Calls": jml_calls,
                    "Calls_Success": jml_calls_success,
                    "Connected_Rate": connected_rate,
                    "Work_Time": work_time,
                    "SMS": jml_sms,
                    "WABA": waba_sent,
                    "Penambahan": pen_hari_ini,
                    "Penerimaan": rec_hari_ini,
                    "Rec_Rate": rec_rate,
                })

            df_prod_all = pd.DataFrame(agent_productivity_list)
            if df_prod_all.empty:
                st.info("Tidak ada agen MPT05 yang ditemukan pada data harian hari ini.")
            else:
                for tl_team, group in df_prod_all.groupby("Nama TL"):
                    first_item = group.iloc[0]
                    p_code = first_item["Project"]
                    b_code = first_item["Bucket"]

                    div_key = f"{b_code} - {p_code}"
                    if div_key not in target_configs:
                        div_key = next((k for k in target_configs if b_code in k and p_code in k), TARGET_KEYS[0])

                    tg = target_configs.get(div_key, {"call": 300, "work": 8.0, "sms": 75, "waba": 0, "rec_rate": 15.0})

                    team_header = (
                        f'<div style="background-color: #064e3b; color: #ecfdf5; padding: 7px 14px; '
                        f'border-radius: 6px; font-weight: 700; font-size: 0.95rem; margin-top: 15px; margin-bottom: 8px;">'
                        f'👥 Team: {tl_team} &nbsp;|&nbsp; Divisi: {div_key} &nbsp;'
                        f'(Target Call: {tg["call"]} | Waktu: {tg["work"]} Jam | SMS: {tg["sms"]})'
                        f'</div>'
                    )
                    st.markdown(team_header, unsafe_allow_html=True)

                    html_tbl = [
                        '<div style="width: 100%; overflow-x: auto; margin-bottom: 1.2rem; border: 1.5px solid #94a3b8; border-radius: 4px; background-color: #ffffff;">'
                        '<table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: -apple-system, sans-serif; font-size: 0.82rem; text-align: center; background-color: #ffffff;">'
                        '<thead><tr>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">No</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">ID Agent</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">Nama Agent</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">Jumlah Total Data</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">Jumlah Panggilan</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">Connected Rate</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">Waktu Kerja (Jam)</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">Jumlah SMS</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">Template WABA</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">Penambahan</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px; border-right: 1px solid #047857;">Penerimaan</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 8px 4px;">Recovery Rate</th>'
                        '</tr></thead><tbody>'
                    ]

                    group_sorted = group.sort_values(by="Rec_Rate", ascending=False).reset_index(drop=True)

                    for idx, row in group_sorted.iterrows():
                        c_call = "color: #dc2626; font-weight: 700;" if row["Calls"] < tg["call"] else "color: #0f172a;"
                        c_work = "color: #dc2626; font-weight: 700;" if row["Work_Time"] < tg["work"] else "color: #0f172a;"
                        c_sms = "color: #dc2626; font-weight: 700;" if row["SMS"] < tg["sms"] else "color: #0f172a;"
                        c_waba = "color: #dc2626; font-weight: 700;" if row["WABA"] < tg["waba"] else "color: #0f172a;"
                        c_rec = "color: #dc2626; font-weight: 700;" if row["Rec_Rate"] < tg["rec_rate"] else "color: #0f172a;"

                        row_html = (
                            f'<tr style="border-bottom: 1px solid #cbd5e1;">'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1;">{idx+1}</td>'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1; font-weight: 600;">{row["ID Agent"]}</td>'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1;">{row["Nama Agent"]}</td>'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1;">{int(row["Total_Data"]):,}</td>'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1; {c_call}">{int(row["Calls"]):,}</td>'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1;">{row["Connected_Rate"]:.2f}%</td>'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1; {c_work}">{row["Work_Time"]:.2f}</td>'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1; {c_sms}">{int(row["SMS"]):,}</td>'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1; {c_waba}">{int(row["WABA"]):,}</td>'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1;">{format_num(row["Penambahan"], is_currency=True)}</td>'
                            f'<td style="padding: 6px 4px; border-right: 1px solid #cbd5e1;">{format_num(row["Penerimaan"], is_currency=True)}</td>'
                            f'<td style="padding: 6px 4px; {c_rec}">{row["Rec_Rate"]:.2f}%</td>'
                            f'</tr>'
                        )
                        html_tbl.append(row_html)

                    tot_pen = group_sorted["Penambahan"].sum()
                    tot_rec = group_sorted["Penerimaan"].sum()
                    avg_rate = (tot_rec / tot_pen * 100.0) if tot_pen > 0 else 0.0
                    tot_data = group_sorted["Total_Data"].sum()
                    avg_calls = group_sorted["Calls"].mean()
                    tot_calls_succ = group_sorted["Calls_Success"].sum()
                    tot_calls_all = group_sorted["Calls"].sum()
                    avg_conn = (tot_calls_succ / tot_calls_all * 100.0) if tot_calls_all > 0 else 0.0
                    avg_work = group_sorted["Work_Time"].mean()
                    tot_sms = group_sorted["SMS"].sum()
                    tot_waba = group_sorted["WABA"].sum()

                    total_html = (
                        f'<tr style="background-color: #064e3b; color: #ecfdf5; font-weight: bold; border-top: 2px solid #059669;">'
                        f'<td colspan="3" style="padding: 7px 4px; text-align: center; border-right: 1px solid #047857;">TOTAL / RATA-RATA TIM ({len(group_sorted)} AGENT)</td>'
                        f'<td style="padding: 7px 4px; border-right: 1px solid #047857;">{int(tot_data):,}</td>'
                        f'<td style="padding: 7px 4px; border-right: 1px solid #047857;">{int(avg_calls):,}</td>'
                        f'<td style="padding: 7px 4px; border-right: 1px solid #047857;">{avg_conn:.2f}%</td>'
                        f'<td style="padding: 7px 4px; border-right: 1px solid #047857;">{avg_work:.2f}</td>'
                        f'<td style="padding: 7px 4px; border-right: 1px solid #047857;">{int(tot_sms):,}</td>'
                        f'<td style="padding: 7px 4px; border-right: 1px solid #047857;">{int(tot_waba):,}</td>'
                        f'<td style="padding: 7px 4px; border-right: 1px solid #047857;">{format_num(tot_pen, is_currency=True)}</td>'
                        f'<td style="padding: 7px 4px; border-right: 1px solid #047857;">{format_num(tot_rec, is_currency=True)}</td>'
                        f'<td style="padding: 7px 4px;">{avg_rate:.2f}%</td>'
                        f'</tr>'
                    )
                    html_tbl.append(total_html)
                    html_tbl.append('</tbody></table></div>')

                    st.markdown("".join(html_tbl), unsafe_allow_html=True)

# =========================================================
# MENU 8: WARNING LETTER (KHUSUS >7 HARI KERJA)
# =========================================================
with tab_wl:
    st.caption("Daftar Warning Letter untuk Agen dan Team Leader (警告信名单 — Khusus Agen > 7 Hari Kerja)")
    if not wl_dict:
        st.info("Belum ada file mingguan di folder `rawdata/WarningLetter/` (contoh: `1-7-SK.xlsx`, `8-14-LD.xlsx`, `15-21-SK.xlsx`).")
    else:
        wl_period_keys = sorted(wl_dict.keys(), key=lambda x: [int(n) for n in re.findall(r"\d+", x)] or [0])
        col_w1, col_w2, col_w3 = st.columns([2, 2.5, 1.5])
        with col_w1:
            selected_wl_period = st.selectbox("Pilih Periode Mingguan (Agent WL):", wl_period_keys, key="sel_wl_period")
        
        files_for_period = wl_dict.get(selected_wl_period, [])
        df_wl_raw = process_agent_performance_data(files_for_period, [])
        df_wl_ag, _ = calculate_warning_letter_data(df_wl_raw, f"Periode {selected_wl_period}")

        df_wl_tl = pd.DataFrame()
        label_tl_period = selected_wl_period
        if len(wl_period_keys) >= 2:
            p_idx = wl_period_keys.index(selected_wl_period)
            pair_start = (p_idx // 2) * 2
            pair_keys = wl_period_keys[pair_start : pair_start + 2]
            combined_tl_files = []
            for pk in pair_keys: combined_tl_files.extend(wl_dict.get(pk, []))
            label_tl_period = " & ".join(pair_keys)
            with col_w2: st.info(f"Evaluasi TL 2 Mingguan: Periode **{label_tl_period}**")
            df_wl_tl_raw = process_agent_performance_data(combined_tl_files, [])
            _, df_wl_tl = calculate_warning_letter_data(df_wl_tl_raw, f"2-Minggu ({label_tl_period})")

        cols_display_ag = ["Project", "Bucket", "Rank", "ID Agent", "Nama Pengguna", "Nama TL", "Peringkat di Bucket", "Hari Kerja", "Total Penambahan (Periode)", "Total Penerimaan (Periode)", "Recovery Rate (%)", "Status WL"]
        cols_display_tl = ["Project", "Bucket", "Rank", "Nama TL", "Perusahaan", "Peringkat di Bucket", "Total Agent", "Total Penambahan (Periode)", "Total Penerimaan (Periode)", "Recovery Rate (%)", "Status WL"]

        df_export_ag = pd.DataFrame()
        if df_wl_ag is not None and not df_wl_ag.empty:
            df_export_ag = df_wl_ag[[c for c in cols_display_ag if c in df_wl_ag.columns]].rename(columns={"Peringkat di Bucket": "Rank / Total"})

        df_export_tl = pd.DataFrame()
        if df_wl_tl is not None and not df_wl_tl.empty:
            df_export_tl = df_wl_tl[[c for c in cols_display_tl if c in df_wl_tl.columns]].rename(columns={"Peringkat di Bucket": "Rank / Total"})

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            if not df_export_ag.empty:
                df_export_ag.to_excel(writer, sheet_name="WL_Agent_MPT05", index=False)
            else:
                pd.DataFrame({"Keterangan": ["N/A"]}).to_excel(writer, sheet_name="WL_Agent_MPT05", index=False)
            
            if not df_export_tl.empty:
                df_export_tl.to_excel(writer, sheet_name="WL_TL_MPT05", index=False)
            else:
                pd.DataFrame({"Keterangan": ["N/A"]}).to_excel(writer, sheet_name="WL_TL_MPT05", index=False)

        with col_w3:
            st.write("")
            st.download_button(
                label="Unduh Data WL (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=f"Warning_Letter_MPT05_Periode_{selected_wl_period}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        wl_sub_tabs = st.tabs([label for label, _, _ in TARGET_BUCKETS])
        def render_wl_tables(df_target_ag, df_target_tl, title_text=None):
            if title_text: st.markdown(f'<div class="clean-group-title">{title_text}</div>', unsafe_allow_html=True)
            col_l, col_r = st.columns([1, 1], gap="small")
            with col_l:
                st.markdown('<div class="section-banner-title">Warning Letter Agent (Bottom 30% — >7 Hari)</div>', unsafe_allow_html=True)
                if df_target_ag is None or df_target_ag.empty: st.success("Tidak ada agen MPT05 yang terkena Warning Letter di bucket ini.")
                else: display_full_table(translate_columns(df_target_ag[[c for c in cols_display_ag if c in df_target_ag.columns]]))
            with col_r:
                st.markdown('<div class="section-banner-title tl">Warning Letter TL (Evaluasi 2 Minggu)</div>', unsafe_allow_html=True)
                if df_target_tl is None or df_target_tl.empty: st.success("Tidak ada Team Leader MPT05 yang terkena Warning Letter di bucket ini.")
                else: display_full_table(translate_columns(df_target_tl[[c for c in cols_display_tl if c in df_target_tl.columns]]))

        for tab_w, (w_label, b_filt, p_filt) in zip(wl_sub_tabs, TARGET_BUCKETS):
            with tab_w:
                if b_filt and p_filt:
                    ag_b = df_wl_ag[(df_wl_ag["Bucket"] == b_filt) & (df_wl_ag["Project"] == p_filt)] if df_wl_ag is not None and not df_wl_ag.empty else pd.DataFrame()
                    tl_b = df_wl_tl[(df_wl_tl["Bucket"] == b_filt) & (df_wl_tl["Project"] == p_filt)] if df_wl_tl is not None and not df_wl_tl.empty else pd.DataFrame()
                    render_wl_tables(ag_b, tl_b)
                else:
                    projects_ag = df_wl_ag["Project"].tolist() if df_wl_ag is not None and not df_wl_ag.empty else []
                    projects_tl = df_wl_tl["Project"].tolist() if df_wl_tl is not None and not df_wl_tl.empty else []
                    all_p = sorted(list(set(projects_ag + projects_tl)))
                    for p_item in all_p:
                        ag_p = df_wl_ag[df_wl_ag["Project"] == p_item] if df_wl_ag is not None and not df_wl_ag.empty else pd.DataFrame()
                        tl_p = df_wl_tl[df_wl_tl["Project"] == p_item] if df_wl_tl is not None and not df_wl_tl.empty else pd.DataFrame()
                        b_list = sorted(list(set(ag_p["Bucket"].tolist() + tl_p["Bucket"].tolist())))
                        for b_item in b_list:
                            render_wl_tables(ag_p[ag_p["Bucket"] == b_item] if not ag_p.empty else pd.DataFrame(), tl_p[tl_p["Bucket"] == b_item] if not tl_p.empty else pd.DataFrame(), f"Project: {p_item} | Bucket: {b_item}")

# =========================================================
# HELPER HISTORIS AGENT & TL
# =========================================================
all_period_records = {}
for p_code in sorted(track_agent_dict.keys()):
    df_p_proc = process_agent_performance_data(track_agent_dict[p_code], [])
    if not df_p_proc.empty:
        all_period_records[f"Bln {p_code[:2]}"] = df_p_proc[df_p_proc["Apakah Ikut Berkompetisi"] == "Ya"].copy()

if not df_agents_processed.empty:
    cur_m_str = datetime.now().strftime("%m")
    all_period_records[f"Bln {cur_m_str} (MTD)"] = df_agents_processed[df_agents_processed["Apakah Ikut Berkompetisi"] == "Ya"].copy()

# =========================================================
# MENU 9: TRACK RECORD PERFORMANCE AGENTS (RESIGN DI PALING BAWAH)
# =========================================================
with tab_track_agent:
    st.caption("Memantau rekam jejak historis agen MPT05 lintas bulan dari folder `rawdata/TrackRecordAgent/` + MTD Berjalan.")
    if not all_period_records:
        st.info("Belum ada data historis di folder `rawdata/TrackRecordAgent/`.")
    else:
        mpt_agents_dict = {}
        for p_label, df_p in all_period_records.items():
            for (p, b), grp in df_p.groupby(["Project", "Bucket"]):
                grp_sorted = grp.copy()
                grp_sorted["_rate"] = (grp_sorted["Total Penambahan Hari Ini_Total"] > 0).astype(int) * (
                    grp_sorted["Total Penerimaan Hari Ini_Total"] / grp_sorted["Total Penambahan Hari Ini_Total"].replace(0, 1) * 100
                )
                grp_sorted = grp_sorted.sort_values(by="_rate", ascending=False).reset_index(drop=True)
                total_in_bucket = len(grp_sorted)

                for idx_rank, ag in grp_sorted.iterrows():
                    perush_name = str(ag.get("Perusahaan", "")).upper()
                    tl_name = str(ag.get("Nama TL", "")).upper()

                    if "MPT05" in perush_name or "MPT05" in tl_name:
                        id_ag = str(ag.get("ID Agent", "")).strip()
                        if id_ag not in mpt_agents_dict:
                            mpt_agents_dict[id_ag] = {
                                "ID Agent": id_ag, "Nama Pengguna": str(ag.get("Nama Pengguna", "")).strip(),
                                "Status Kerja": ag.get("Status Kerja", "Aktif"), "rates": [], "is_top_30_list": [], "is_bottom_30_list": [], "history": {},
                            }

                        rec_rate_val = ag["_rate"]
                        actual_rank = idx_rank + 1
                        pct_in_bucket = actual_rank / total_in_bucket if total_in_bucket > 0 else 1.0

                        mpt_agents_dict[id_ag]["rates"].append(rec_rate_val)
                        mpt_agents_dict[id_ag]["is_top_30_list"].append(pct_in_bucket <= 0.30)
                        mpt_agents_dict[id_ag]["is_bottom_30_list"].append(pct_in_bucket > 0.70)
                        mpt_agents_dict[id_ag]["history"][p_label] = f"{rec_rate_val:.2f}% (Rank #{actual_rank}/{total_in_bucket}) | {p}-{b}"

        history_rows = []
        p_columns = list(all_period_records.keys())
        for id_ag, data in mpt_agents_dict.items():
            avg_rate = sum(data["rates"]) / len(data["rates"]) if data["rates"] else 0.0
            tot_m = len(data["rates"])

            if tot_m <= 1: trend_label = "Data Baru 🆕"
            elif all(data["is_top_30_list"]): trend_label = "Konsisten TOP 🌟"
            elif all(data["is_bottom_30_list"]): trend_label = "Konsisten BOTTOM ⚠️"
            elif all(data["rates"][i] < data["rates"][i+1] for i in range(tot_m - 1)): trend_label = "Terus Meningkat 📈"
            elif all(data["rates"][i] > data["rates"][i+1] for i in range(tot_m - 1)): trend_label = "Terus Menurun 📉"
            else: trend_label = "Fluktuatif 🔄"

            is_resign_val = 1 if str(data["Status Kerja"]).strip().lower() == "resign" else 0

            row_item = {
                "ID Agent": data["ID Agent"], "Nama Pengguna": data["Nama Pengguna"], "Status Kerja": data["Status Kerja"],
                "Rata-rata Rate (All-Time)": f"{avg_rate:.2f}%", "Analisis Kinerja": trend_label,
                "_is_resign": is_resign_val,
                "_order_val": RANK_ORDER.get(trend_label, 9), "_avg_sort": avg_rate,
            }
            for p_col in p_columns: row_item[p_col] = data["history"].get(p_col, "-")
            history_rows.append(row_item)

        df_hist_ag = pd.DataFrame(history_rows)
        if df_hist_ag.empty: st.info("Tidak ada data agen MPT05.")
        else:
            # PENGURUTAN: AKTIF DI ATAS, RESIGN DI PALING BAWAH
            df_hist_ag = df_hist_ag.sort_values(
                by=["_is_resign", "_order_val", "_avg_sort"],
                ascending=[True, True, False]
            ).reset_index(drop=True)
            df_hist_ag.insert(0, "Ranking (Avg)", [f"#{i+1}" for i in range(len(df_hist_ag))])
            df_hist_ag.drop(columns=["_is_resign", "_order_val", "_avg_sort"], inplace=True)
            display_full_table(translate_columns(df_hist_ag))

# =========================================================
# MENU 10: TRACK RECORD PERFORMANCE TL
# =========================================================
with tab_track_tl:
    st.caption("Memantau rekam jejak kinerja seluruh Team Leader MPT05 lintas bulan dari folder `rawdata/Bulanan/`.")
    if not all_bulanan_dict:
        st.info("Belum ada data bulanan di folder `rawdata/Bulanan/`.")
    else:
        tl_period_records = {}
        for b_code in sorted(all_bulanan_dict.keys()):
            df_b_proc = process_department_weekly_data(all_bulanan_dict[b_code])
            if not df_b_proc.empty:
                tl_period_records[f"Bulan {b_code}"] = df_b_proc

        mpt_tl_dict = {}
        for p_label, df_p in tl_period_records.items():
            for (p, b), grp in df_p.groupby(["Project", "Bucket"]):
                tl_summary = []
                for tl_name, tl_grp in grp.groupby("Nama TL"):
                    pen = tl_grp["Total Penambahan"].sum()
                    rec = tl_grp["Total Penerimaan"].sum()
                    rate = (rec / pen * 100) if pen > 0 else 0.0
                    tl_summary.append({
                        "Nama TL": tl_name, "Perusahaan": tl_grp.iloc[0].get("Perusahaan", ""),
                        "rate": rate, "pen": pen, "rec": rec,
                    })

                df_tl_grp = pd.DataFrame(tl_summary).sort_values(by="rate", ascending=False).reset_index(drop=True)
                total_tls = len(df_tl_grp)

                for idx_rank, tl_r in df_tl_grp.iterrows():
                    tl_name = str(tl_r["Nama TL"]).strip()
                    perush_name = str(tl_r["Perusahaan"]).upper()

                    if "MPT05" in tl_name.upper() or "MPT05" in perush_name:
                        if tl_name not in mpt_tl_dict:
                            mpt_tl_dict[tl_name] = {"Nama TL": tl_name, "Status TL": "Aktif", "rates": [], "ranks_pct": [], "history": {}}

                        rec_rate_val = tl_r["rate"]
                        actual_rank = idx_rank + 1
                        mpt_tl_dict[tl_name]["rates"].append(rec_rate_val)
                        mpt_tl_dict[tl_name]["ranks_pct"].append(actual_rank / total_tls if total_tls > 0 else 1.0)
                        mpt_tl_dict[tl_name]["history"][p_label] = f"{rec_rate_val:.2f}% (Rank #{actual_rank}/{total_tls}) | {p}-{b}"

        tl_history_rows = []
        p_columns = list(tl_period_records.keys())
        for tl_name, data in mpt_tl_dict.items():
            avg_rate = sum(data["rates"]) / len(data["rates"]) if data["rates"] else 0.0
            tot_tl_m = len(data["rates"])

            if tot_tl_m <= 1: trend_label = "Data Baru 🆕"
            elif all(pct <= 0.30 for pct in data["ranks_pct"]): trend_label = "Konsisten TOP 🌟"
            elif all(pct >= 0.70 for pct in data["ranks_pct"]): trend_label = "Konsisten BOTTOM ⚠️"
            elif all(data["rates"][i] <= data["rates"][i+1] for i in range(tot_tl_m - 1)): trend_label = "Terus Meningkat 📈"
            elif all(data["rates"][i] >= data["rates"][i+1] for i in range(tot_tl_m - 1)): trend_label = "Terus Menurun 📉"
            else: trend_label = "Fluktuatif 🔄"

            row_item = {
                "Nama TL": data["Nama TL"], "Status TL": data["Status TL"],
                "Rata-rata Rate (All-Time)": f"{avg_rate:.2f}%", "Analisis Kinerja": trend_label,
                "_order_val": RANK_ORDER.get(trend_label, 9), "_avg_sort": avg_rate,
            }
            for p_col in p_columns: row_item[p_col] = data["history"].get(p_col, "-")
            tl_history_rows.append(row_item)

        df_hist_tl = pd.DataFrame(tl_history_rows)
        if df_hist_tl.empty: st.info("Tidak ada data TL MPT05.")
        else:
            df_hist_tl = df_hist_tl.sort_values(by=["_order_val", "_avg_sort"], ascending=[True, False]).reset_index(drop=True)
            df_hist_tl.insert(0, "Ranking (Avg)", [f"#{i+1}" for i in range(len(df_hist_tl))])
            df_hist_tl.drop(columns=["_order_val", "_avg_sort"], inplace=True)
            display_full_table(translate_columns(df_hist_tl))

# =========================================================
# MENU 11: TRACK RECORD PERFORMANCE GABUNGAN (TIDAK DOBEL OKTOBER)
# =========================================================
with tab_track_comp:
    st.caption("Perbandingan Recovery Rate dan Ranking antar Perusahaan bersebelahan lintas bulan (Bulan Final + MTD Berjalan + Daily).")
    
    comp_period_data = {}
    sorted_all_months = sorted(all_bulanan_dict.keys())
    past_month_keys = sorted_all_months[:-1] if len(sorted_all_months) > 1 else []

    # 1. Tampilkan hanya bulan yang sudah final (Bulan 07, 08, 09)
    for b_code in past_month_keys:
        df_b_p = process_department_weekly_data(all_bulanan_dict[b_code])
        if not df_b_p.empty:
            comp_period_data[f"Bulan {b_code}/2026 (2026年{int(b_code)}月份)"] = df_b_p

    # 2. Bulan Berjalan Terakhir (10. Okt-LD & 10. Okt-SK) HANYA tampil sekali di MTD Berjalan + Daily Hari Ini
    if "Monthly" in weekly_dict:
        df_curr_m = process_department_weekly_data(weekly_dict["Monthly"])
        if not df_curr_m.empty:
            acc_rows = []
            for (p_c, b_c, comp_n), grp in df_curr_m.groupby(["Project", "Bucket", "Perusahaan"]):
                pen_f = grp["Total Penambahan"].sum()
                rec_f = grp["Total Penerimaan"].sum()

                d_val = daily_comp_map.get((p_c, b_c, comp_n), {"pen": 0.0, "rec": 0.0})
                pen_tot = pen_f + d_val["pen"]
                rec_tot = rec_f + d_val["rec"]

                acc_rows.append({
                    "Project": p_c,
                    "Bucket": b_c,
                    "Nama TL": grp.iloc[0]["Nama TL"],
                    "Perusahaan": comp_n,
                    "Total Penambahan": pen_tot,
                    "Total Penerimaan": rec_tot,
                })
            comp_period_data["MTD Berjalan (当月至今)"] = pd.DataFrame(acc_rows)

    if not comp_period_data:
        st.info("Belum ada data bulanan di folder `rawdata/Bulanan/`.")
    else:
        def get_compact_company_table(df_subset):
            rows = []
            for comp_name, grp in df_subset.groupby("Perusahaan"):
                pen = grp["Total Penambahan"].sum()
                rec = grp["Total Penerimaan"].sum()
                rate = (rec / pen * 100) if pen > 0 else 0.0
                first_r = grp.iloc[0]
                rows.append({
                    "Project": first_r.get("Project", ""),
                    "Bucket": first_r.get("Bucket", ""),
                    "Perusahaan": comp_name,
                    "_rate_num": rate,
                    "Recovery Rate": f"{rate:.2f}%",
                })
            df_res = pd.DataFrame(rows)
            if df_res.empty: return pd.DataFrame()
            df_res = df_res.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
            df_res["Ranking"] = [f"#{i+1}" for i in range(len(df_res))]
            return df_res[["Project", "Bucket", "Perusahaan", "Recovery Rate", "Ranking"]]

        target_groups = [("SK", "D0"), ("SK", "S1"), ("LD", "S0"), ("LD", "S1"), ("LD-BA", "S1"), ("LD-C", "C")]
        period_keys = list(comp_period_data.keys())

        for proj, buck in target_groups:
            has_data = any(
                not comp_period_data[k][(comp_period_data[k]["Project"] == proj) & (comp_period_data[k]["Bucket"] == buck)].empty
                for k in period_keys
            )
            if not has_data: continue

            st.markdown(f'<div class="clean-group-title">Project (项目): {proj} &nbsp;|&nbsp; Bucket (阶段): {buck}</div>', unsafe_allow_html=True)
            cols = st.columns(len(period_keys))
            for idx, p_label in enumerate(period_keys):
                with cols[idx]:
                    st.markdown(f'<div class="month-header-clean">{p_label}</div>', unsafe_allow_html=True)
                    df_period_all = comp_period_data[p_label]
                    df_p_b = df_period_all[(df_period_all["Project"] == proj) & (df_period_all["Bucket"] == buck)]
                    if df_p_b.empty:
                        st.caption("Tidak ada data (暂无数据)")
                    else:
                        tbl_comp = get_compact_company_table(df_p_b)
                        if not tbl_comp.empty:
                            display_full_table(translate_columns(tbl_comp))
            st.write("")