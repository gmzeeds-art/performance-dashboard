import os
import glob
import re
from datetime import datetime
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# ---------------------------------------------------------
# KONFIGURASI HALAMAN
# ---------------------------------------------------------
st.set_page_config(
    page_title="Performance Collection Automation",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    [data-testid="stSidebar"] { display: none; }
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        max-width: 99%;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; margin-bottom: 0.6rem; }
    .stTabs [data-baseweb="tab"] {
        padding: 7px 15px;
        font-weight: 600;
        font-size: 0.90rem;
    }
    .clean-group-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #38bdf8;
        padding-top: 10px;
        padding-bottom: 4px;
        border-bottom: 1px solid #334155;
        margin-bottom: 8px;
    }
    .sub-table-header {
        font-size: 0.85rem;
        font-weight: 700;
        padding: 4px 6px;
        border-radius: 4px;
        margin-bottom: 4px;
        text-align: center;
    }
    .header-bulanan {
        background-color: #0f172a;
        color: #38bdf8;
        border: 1px solid #1e293b;
    }
    .header-harian {
        background-color: #1e1b4b;
        color: #a5b4fc;
        border: 1px solid #312e81;
    }
    .header-wl {
        background-color: #450a0a;
        color: #f87171;
        border: 1px solid #7f1d1d;
    }
    .month-header-clean {
        background-color: #1e293b;
        border: 1px solid #334155;
        padding: 4px 8px;
        border-radius: 4px;
        text-align: center;
        font-weight: bold;
        font-size: 0.88rem;
        color: #38bdf8;
        margin-bottom: 5px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAWDATA_DIR = os.path.join(BASE_DIR, "rawdata")
MASTER_FILE = os.path.join(BASE_DIR, "data_karyawan.xlsx")
os.makedirs(RAWDATA_DIR, exist_ok=True)

# ---------------------------------------------------------
# DICTIONARY TRANSLASI JUDUL KOLOM KE MANDARIN (DENGAN <br>)
# ---------------------------------------------------------
HEADER_TRANSLATIONS = {
    "ID Pengguna": "ID Pengguna<br>(用户ID)",
    "Nama Pengguna": "Nama Pengguna<br>(用户名)",
    "Perusahaan": "Perusahaan<br>(公司)",
    "Status Kerja": "Status Kerja<br>(工作状态)",
    "Jenis Tim": "Jenis Tim<br>(团队类型)",
    "Jenis Kasus": "Jenis Kasus<br>(案件类型)",
    "Bucket": "Bucket<br>(逾期阶段)",
    "Hiredate": "Hiredate<br>(入职日期)",
    "Project": "Project<br>(项目)",
    "Apakah Ikut Berkompetisi": "Ikut Kompetisi<br>(是否参赛)",
    "Rank": "Rank<br>(排名)",
    "Ranking": "Ranking<br>(排名)",
    "Ranking (Avg)": "Ranking Avg<br>(平均排名)",
    "Ranking Interval": "Ranking Interval<br>(排名区间)",
    "Nama TL": "Nama TL<br>(组长姓名)",
    "ID Agent": "ID Agent<br>(坐席工号)",
    "Total Agent": "Total Agent<br>(总坐席数)",
    "Total Penambahan (Monthly)": "Total Penambahan (M)<br>(月累计新增本金)",
    "Total Penerimaan (Monthly)": "Total Penerimaan (M)<br>(月累计回款总额)",
    "Total Penambahan (Periode)": "Penambahan (Periode)<br>(周期新增本金)",
    "Total Penerimaan (Periode)": "Penerimaan (Periode)<br>(周期回款总额)",
    "Recovery Rate (Monthly)": "Recovery Rate (M)<br>(月催回率)",
    "Recovery Rate (%)": "Recovery Rate<br>(催回率)",
    "Recovery Rate": "Recovery Rate<br>(催回率)",
    "Gap": "Gap<br>(差距率)",
    "Gap Amount": "Gap Amount<br>(差距金额)",
    "Penambahan (Hari Ini)": "Penambahan (Today)<br>(今日新增本金)",
    "Penerimaan (Hari Ini)": "Penerimaan (Today)<br>(今日回款总额)",
    "Recovery Rate (Hari Ini)": "Recovery Rate (Today)<br>(今日催回率)",
    "Status WL": "Status WL<br>(警告信状态)",
    "Peringkat di Bucket": "Rank / Total<br>(组内排名/总数)",
    "Persentil Rank": "Persentil<br>(排名百分比)",
    "Periode Evaluasi": "Periode<br>(考核周期)",
    "Jumlah Panggilan": "Panggilan<br>(呼叫次数)",
    "Jumlah Panggilan (Avg)": "Panggilan (Avg)<br>(平均呼叫次数)",
    "Waktu Kerja Rata-rata": "Waktu Kerja<br>(平均工作时长)",
    "Waktu Kerja Rata-rata (Avg)": "Waktu Kerja (Avg)<br>(平均工作时长)",
    "Jumlah SMS": "SMS<br>(短信发送量)",
    "Jumlah SMS (Avg)": "SMS (Avg)<br>(平均短信量)",
    "Jumlah Pengiriman Template WABA": "Template WABA<br>(WABA模板发送量)",
    "Jumlah Pengiriman Template WABA (Avg)": "Template WABA (Avg)<br>(平均WABA发送量)",
    "Status TL": "Status TL<br>(组长状态)",
    "Rata-rata Rate (All-Time)": "Avg Rate (All-Time)<br>(历史平均催回率)",
    "Analisis Kinerja": "Analisis Kinerja<br>(绩效分析)",
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
# RENDER TABEL RESPONSIVE, AUTO WRAP & FREEZE HEADER
# ---------------------------------------------------------
def display_full_table(df_input, row_count=None):
    df = df_input.data if hasattr(df_input, "data") else df_input

    html = ['<div style="width: 100%; overflow-x: auto; margin-bottom: 0.6rem; border: 1px solid #334155; border-radius: 4px;">']
    html.append(
        '<table style="width: 100%; border-collapse: separate; border-spacing: 0; '
        'font-family: -apple-system, BlinkMacSystemFont, sans-serif; '
        'font-size: 0.81rem; text-align: center;">'
    )
    
    html.append('<thead><tr>')
    for col in df.columns:
        col_header = str(col).replace("\n", "<br>")
        html.append(
            f'<th style="position: sticky; top: 0; z-index: 5; background-color: #1e293b; color: #f8fafc; '
            f'padding: 7px 4px; font-weight: 600; white-space: normal; line-height: 1.2; vertical-align: middle; '
            f'border-bottom: 2px solid #475569; border-right: 1px solid #334155; word-break: break-word;">{col_header}</th>'
        )
    html.append('</tr></thead><tbody>')
    
    for _, row in df.iterrows():
        row_str = " ".join([str(v) for v in row.values]).upper()
        is_mpt05 = "MPT05" in row_str
        is_resign = "RESIGN" in row_str
        is_wl = "WL" in row_str or "WARNING" in row_str
        
        if is_wl:
            row_style = "background-color: rgba(239, 68, 68, 0.22); font-weight: bold; color: #fca5a5;"
        elif is_resign:
            row_style = "background-color: rgba(148, 163, 184, 0.12); color: #94a3b8;"
        elif is_mpt05:
            row_style = "background-color: rgba(234, 179, 8, 0.35); font-weight: bold; color: inherit;"
        else:
            row_style = "background-color: transparent; color: inherit;"
            
        html.append(f'<tr style="{row_style}">')
        for val in row.values:
            html.append(f'<td style="padding: 5px 4px; border-bottom: 1px solid #334155; border-right: 1px solid #334155; vertical-align: middle; white-space: nowrap;">{val}</td>')
        html.append('</tr>')
        
    html.append('</tbody></table></div>')
    st.markdown("".join(html), unsafe_allow_html=True)

# ---------------------------------------------------------
# HELPER PARSER & FORMATTER
# ---------------------------------------------------------
def clean_numeric(val):
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace(",", "")
    val_str = re.sub(r"[^\d.-]", "", val_str)
    try:
        return float(val_str) if val_str != "" else 0.0
    except ValueError:
        return 0.0

def format_num(val, is_currency=False):
    if pd.isna(val) or val == 0:
        return "0"
    if is_currency:
        return f"{int(round(val)):,}".replace(",", ".")
    return f"{val:,.2f}".rstrip("0").rstrip(".") if isinstance(val, float) else str(val)

def normalize_clean(val):
    if pd.isna(val):
        return ""
    return " ".join(str(val).strip().upper().split())

def map_status_kerja(st_raw):
    s = str(st_raw).strip().lower()
    if "mengundurkan" in s or "resign" in s or "tidak aktif" in s or "non aktif" in s:
        return "Resign"
    elif "aktif" in s or "istirahat" in s:
        return "Aktif"
    return "Resign" if s != "" and s != "nan" else "-"

def get_ranking_interval(rank, total):
    if total <= 0:
        return "-"
    pct = rank / total
    if pct <= 0.10:
        return "TOP10%"
    elif pct <= 0.20:
        return "10-20%"
    elif pct <= 0.30:
        return "20-30%"
    elif pct <= 0.50:
        return "30-50%"
    elif pct <= 0.70:
        return "50-70%"
    elif pct <= 0.80:
        return "70-80%"
    else:
        return "BOTTOM20%"

# ---------------------------------------------------------
# 1. LOAD DATABASE KARYAWAN & TAB MASTER
# ---------------------------------------------------------
@st.cache_data(ttl=300, show_spinner=False)
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
# 2. SCAN & BACA FILE RAWDATA (AKURAT MTD VS BULAN FINAL)
# ---------------------------------------------------------
METRIC_SUM_COLS = ["Total Penambahan Hari Ini", "Total Penerimaan Hari Ini"]
METRIC_AVG_COLS = ["Jumlah Panggilan", "Waktu Kerja Rata-rata", "Jumlah SMS", "Jumlah Pengiriman Template WABA"]
ALL_METRICS = METRIC_SUM_COLS + METRIC_AVG_COLS

@st.cache_data(ttl=300, show_spinner=False)
def read_excel_standard(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = [str(c).strip() for c in df.columns]
        
        fname = os.path.basename(file_path).upper()
        detected_proj = ""
        if "_SK" in fname or "-SK" in fname or " SK" in fname:
            detected_proj = "SK"
        elif "LD-BA" in fname or "-BA" in fname:
            detected_proj = "LD-BA"
        elif "LD-C" in fname or "-C" in fname:
            detected_proj = "LD-C"
        elif "_LD" in fname or "-LD" in fname or " LD" in fname:
            detected_proj = "LD"
        df["_source_project"] = detected_proj

        for col in ALL_METRICS:
            if col in df.columns:
                df[col] = df[col].apply(clean_numeric)
            else:
                df[col] = 0.0
        return df
    except Exception:
        return pd.DataFrame()

def scan_rawdata_files(folder_path):
    all_files = glob.glob(os.path.join(folder_path, "*.xlsx")) + glob.glob(os.path.join(folder_path, "*.xls"))
    final_files = {}
    raw_mtd_files = []
    raw_daily_files = []
    wl_files = {}

    for f in all_files:
        clean_name = os.path.basename(f).upper()
        m_wl = re.match(r"^(\d+-\d+)[-_ ]?.*", clean_name)
        m_start_digits = re.match(r"^(\d{6})[-_].*", clean_name)

        if m_wl:
            wl_key = m_wl.group(1)
            wl_files.setdefault(wl_key, []).append(f)
        elif m_start_digits:
            periode = m_start_digits.group(1)
            final_files.setdefault(periode, []).append(f)
        elif "MTD" in clean_name:
            raw_mtd_files.append(f)
            m_digits = re.search(r"(\d{6})", clean_name)
            if m_digits:
                final_files.setdefault(m_digits.group(1), []).append(f)
        else:
            raw_daily_files.append(f)

    # Ambil berkas MTD terbaru untuk masing-masing SK dan LD
    mtd_files = []
    sk_mtd = sorted([f for f in raw_mtd_files if "SK" in os.path.basename(f).upper()], key=lambda x: os.path.basename(x).upper())
    ld_mtd = sorted([f for f in raw_mtd_files if "LD" in os.path.basename(f).upper()], key=lambda x: os.path.basename(x).upper())
    if sk_mtd:
        mtd_files.append(sk_mtd[-1])
    if ld_mtd:
        mtd_files.append(ld_mtd[-1])

    # Ambil berkas harian terbaru untuk masing-masing SK dan LD
    daily_files = []
    sk_dailies = sorted([f for f in raw_daily_files if "SK" in os.path.basename(f).upper()], key=lambda x: os.path.basename(x).upper())
    ld_dailies = sorted([f for f in raw_daily_files if "LD" in os.path.basename(f).upper()], key=lambda x: os.path.basename(x).upper())
    if sk_dailies:
        daily_files.append(sk_dailies[-1])
    if ld_dailies:
        daily_files.append(ld_dailies[-1])

    return final_files, mtd_files, daily_files, wl_files

# ---------------------------------------------------------
# PROSES UTAMA
# ---------------------------------------------------------
df_karyawan, df_master = load_database_and_master(MASTER_FILE)
final_dict, mtd_files, daily_files, wl_dict = scan_rawdata_files(RAWDATA_DIR)

col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown("### **Performance Collection Automation (催收绩效自动化系统)**")
with col_h2:
    if st.button("Refresh Data (刷新数据)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

if df_karyawan is None:
    st.error("File `data_karyawan.xlsx` tidak ditemukan.")
    st.stop()

tl_col = "Nama TL" if "Nama TL" in df_karyawan.columns else "Perusahaan"
bucket_col = "Bucket" if "Bucket" in df_karyawan.columns else "Bucket"
project_col = "Project" if "Project" in df_karyawan.columns else "Project"

# Mapping Database Karyawan
karyawan_dict = {}
for _, row in df_karyawan.iterrows():
    id_akun = normalize_clean(row.get("ID Pengguna", ""))
    nama_orang = normalize_clean(row.get("Nama Pengguna", ""))
    proj_exact = str(row.get(project_col, "")).strip().upper()
    raw_status = row.get("Status Kerja", "")

    info = {
        "Nama TL": str(row.get(tl_col, "")).strip(),
        "Bucket": str(row.get(bucket_col, "")).strip(),
        "Project": proj_exact,
        "Status Kerja": map_status_kerja(raw_status),
        "ID_Akun_Resmi": str(row.get("ID Pengguna", "")).strip(),
        "Nama_Orang_Resmi": str(row.get("Nama Pengguna", "")).strip(),
    }

    if id_akun and proj_exact:
        karyawan_dict[(id_akun, proj_exact)] = info
    if nama_orang and proj_exact:
        karyawan_dict[(nama_orang, proj_exact)] = info

    if "LD" in proj_exact:
        if id_akun and (id_akun, "LD") not in karyawan_dict:
            karyawan_dict[(id_akun, "LD")] = info
        if nama_orang and (nama_orang, "LD") not in karyawan_dict:
            karyawan_dict[(nama_orang, "LD")] = info

    if id_akun and id_akun not in karyawan_dict:
        karyawan_dict[id_akun] = info
    if nama_orang and nama_orang not in karyawan_dict:
        karyawan_dict[nama_orang] = info

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

def process_performance_data(mtd_file_list, daily_file_list, periode_label=""):
    if not mtd_file_list and not daily_file_list:
        return pd.DataFrame()

    mtd_dfs = [read_excel_standard(f) for f in mtd_file_list]
    df_mtd_all = pd.concat([d for d in mtd_dfs if not d.empty], ignore_index=True) if mtd_dfs else pd.DataFrame()

    daily_dfs = [read_excel_standard(f) for f in daily_file_list]
    df_daily_all = pd.concat([d for d in daily_dfs if not d.empty], ignore_index=True) if daily_dfs else pd.DataFrame()

    # 1. BACA DATA HARIAN (KUNCI HANYA 1 ENTRI UNIK AGAR TIDAK TERDUPLIKASI)
    daily_map = {}
    if not df_daily_all.empty:
        for _, r in df_daily_all.iterrows():
            id_d = normalize_clean(r.get("Nama Pengguna", ""))
            nm_d = normalize_clean(r.get("Nama", ""))
            p_src = normalize_clean(r.get("_source_project", ""))

            acc_key = id_d if id_d else nm_d
            if not acc_key:
                continue

            vals = {m: clean_numeric(r.get(m, 0)) for m in ALL_METRICS}
            vals["_raw_id"] = str(r.get("Nama Pengguna", "")).strip()
            vals["_raw_nama"] = str(r.get("Nama", "")).strip()
            vals["_proj"] = p_src

            k_tuple = (acc_key, p_src)
            if k_tuple not in daily_map:
                daily_map[k_tuple] = vals
            else:
                daily_map[k_tuple]["Total Penambahan Hari Ini"] += vals["Total Penambahan Hari Ini"]
                daily_map[k_tuple]["Total Penerimaan Hari Ini"] += vals["Total Penerimaan Hari Ini"]

    # 2. BACA DATA MTD (KUNCI HANYA 1 ENTRI UNIK PER AGEN)
    mtd_map = {}
    if not df_mtd_all.empty:
        for _, r in df_mtd_all.iterrows():
            id_m = normalize_clean(r.get("Nama Pengguna", ""))
            nm_m = normalize_clean(r.get("Nama", ""))
            p_src = normalize_clean(r.get("_source_project", ""))

            acc_key = id_m if id_m else nm_m
            if not acc_key:
                continue

            pen_val = clean_numeric(r.get("Total Penambahan Hari Ini", 0))
            rec_val = clean_numeric(r.get("Total Penerimaan Hari Ini", 0))

            k_tuple = (acc_key, p_src)
            if k_tuple not in mtd_map:
                mtd_map[k_tuple] = {
                    "Total Penambahan Hari Ini": pen_val,
                    "Total Penerimaan Hari Ini": rec_val,
                    "_raw_id": str(r.get("Nama Pengguna", "")).strip(),
                    "_raw_nama": str(r.get("Nama", "")).strip(),
                    "_proj": p_src,
                    "_count": 1
                }
                for m in METRIC_AVG_COLS:
                    mtd_map[k_tuple][m] = clean_numeric(r.get(m, 0))
            else:
                mtd_map[k_tuple]["Total Penambahan Hari Ini"] += pen_val
                mtd_map[k_tuple]["Total Penerimaan Hari Ini"] += rec_val
                mtd_map[k_tuple]["_count"] += 1
                for m in METRIC_AVG_COLS:
                    mtd_map[k_tuple][m] += clean_numeric(r.get(m, 0))

        for k_tuple in mtd_map:
            cnt = mtd_map[k_tuple]["_count"]
            if cnt > 1:
                for m in METRIC_AVG_COLS:
                    mtd_map[k_tuple][m] /= cnt

    # 3. GABUNGKAN KUNCI UNIK AGEN
    all_keys = set(mtd_map.keys()).union(set(daily_map.keys()))

    records = []
    for (acc_key, proj_key) in all_keys:
        m_vals = mtd_map.get((acc_key, proj_key), {})
        d_vals = daily_map.get((acc_key, proj_key), {})

        k_info = karyawan_dict.get((acc_key, proj_key)) or {}
        if not k_info:
            alt_id = normalize_clean(d_vals.get("_raw_nama", "") or m_vals.get("_raw_nama", ""))
            k_info = karyawan_dict.get((alt_id, proj_key)) or karyawan_dict.get(acc_key, {})

        tl_val = k_info.get("Nama TL", "Unknown")
        bucket_val = k_info.get("Bucket", "Unknown")
        project_val = k_info.get("Project", proj_key or "Unknown")
        status_val = k_info.get("Status Kerja", "Aktif")

        is_comp, perush_master = get_master_status(tl_val, bucket_val, project_val)

        if "092026" in str(periode_label) and tl_val == "W-MPT05-2":
            is_comp = "Tidak"

        pen_mtd = m_vals.get("Total Penambahan Hari Ini", 0.0)
        rec_mtd = m_vals.get("Total Penerimaan Hari Ini", 0.0)
        pen_daily = d_vals.get("Total Penambahan Hari Ini", 0.0)
        rec_daily = d_vals.get("Total Penerimaan Hari Ini", 0.0)

        row_item = {
            "ID Agent": d_vals.get("_raw_id", "") or m_vals.get("_raw_id", "") or k_info.get("ID_Akun_Resmi", "") or acc_key,
            "Nama Pengguna": d_vals.get("_raw_nama", "") or m_vals.get("_raw_nama", "") or k_info.get("Nama_Orang_Resmi", "") or acc_key,
            "Status Kerja": status_val,
            "Nama TL": tl_val,
            "Bucket": bucket_val,
            "Project": project_val,
            "Perusahaan": perush_master,
            "Apakah Ikut Berkompetisi": is_comp,
            "Total Penambahan Hari Ini_MTD": pen_mtd,
            "Total Penerimaan Hari Ini_MTD": rec_mtd,
            "Total Penambahan Hari Ini_Daily": pen_daily,
            "Total Penerimaan Hari Ini_Daily": rec_daily,
            "Total Penambahan Hari Ini_Total": pen_mtd + pen_daily,
            "Total Penerimaan Hari Ini_Total": rec_mtd + rec_daily,
        }

        for m in METRIC_AVG_COLS:
            row_item[f"{m}_Daily"] = d_vals.get(m, 0.0)
            row_item[f"{m}_MTD"] = m_vals.get(m, 0.0)

        records.append(row_item)

    return pd.DataFrame(records)

df_active_performance = process_performance_data(mtd_files, daily_files)

df_karyawan_status = df_karyawan.copy()
df_karyawan_status["Status Kerja"] = df_karyawan_status["Status Kerja"].apply(map_status_kerja)
k_comp_res = df_karyawan_status.apply(
    lambda r: get_master_status(r.get(tl_col, ""), r.get(bucket_col, ""), r.get(project_col, ""))[0],
    axis=1,
)
df_karyawan_status["Apakah Ikut Berkompetisi"] = k_comp_res

# ---------------------------------------------------------
# DUAL TABLES HELPER
# ---------------------------------------------------------
def build_dual_tables(df_subset, group_col_name, label_entity):
    rows_monthly = []
    rows_daily = []

    for ent_name, group in df_subset.groupby(group_col_name):
        first_row = group.iloc[0]
        p_val = first_row.get("Project", "")
        b_val = first_row.get("Bucket", "")

        tot_pen = group["Total Penambahan Hari Ini_Total"].sum()
        tot_rec = group["Total Penerimaan Hari Ini_Total"].sum()
        rec_rate_monthly = (tot_rec / tot_pen * 100) if tot_pen > 0 else 0.0

        rows_monthly.append({
            "Project": p_val,
            "Bucket": b_val,
            label_entity: ent_name,
            "Total Agent": len(group),
            "_tot_pen": tot_pen,
            "Total Penambahan (Monthly)": format_num(tot_pen, is_currency=True),
            "Total Penerimaan (Monthly)": format_num(tot_rec, is_currency=True),
            "_rate_num": rec_rate_monthly,
            "Recovery Rate (Monthly)": f"{rec_rate_monthly:.2f}%",
        })

        daily_pen = group["Total Penambahan Hari Ini_Daily"].sum()
        daily_rec = group["Total Penerimaan Hari Ini_Daily"].sum()
        rec_rate_daily = (daily_rec / daily_pen * 100) if daily_pen > 0 else 0.0

        rows_daily.append({
            "Project": p_val,
            "Bucket": b_val,
            label_entity: ent_name,
            "_daily_pen": daily_pen,
            "Penambahan (Hari Ini)": format_num(daily_pen, is_currency=True),
            "Penerimaan (Hari Ini)": format_num(daily_rec, is_currency=True),
            "_rate_num": rec_rate_daily,
            "Recovery Rate (Hari Ini)": f"{rec_rate_daily:.2f}%",
            "Jumlah Panggilan (Avg)": f"{group['Jumlah Panggilan_Daily'].mean():.2f}",
            "Waktu Kerja Rata-rata (Avg)": f"{group['Waktu Kerja Rata-rata_Daily'].mean():.2f}",
            "Jumlah SMS (Avg)": f"{group['Jumlah SMS_Daily'].mean():.2f}",
            "Jumlah Pengiriman Template WABA (Avg)": f"{group['Jumlah Pengiriman Template WABA_Daily'].mean():.2f}",
        })

    df_m = pd.DataFrame(rows_monthly)
    if not df_m.empty:
        df_m = df_m.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
        gaps_pct, gaps_amt = ["-"], ["-"]
        for i in range(1, len(df_m)):
            diff = df_m.loc[i - 1, "_rate_num"] - df_m.loc[i, "_rate_num"]
            pen_c = df_m.loc[i, "_tot_pen"]
            if diff > 0 and pen_c > 0:
                gaps_pct.append(f"+{diff:.2f}%")
                gaps_amt.append(f"-{format_num((diff / 100.0) * pen_c, is_currency=True)}")
            else:
                gaps_pct.append("0.00%")
                gaps_amt.append("-")
        df_m["Gap"] = gaps_pct
        df_m["Gap Amount"] = gaps_amt
        df_m["Rank"] = [f"#{i+1}" for i in range(len(df_m))]
        m_cols = ["Rank", "Project", "Bucket", label_entity, "Total Agent", "Total Penambahan (Monthly)", "Total Penerimaan (Monthly)", "Recovery Rate (Monthly)", "Gap", "Gap Amount"]
        df_m = df_m[m_cols]

    df_d = pd.DataFrame(rows_daily)
    if not df_d.empty:
        df_d = df_d.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
        gaps_pct_d, gaps_amt_d = ["-"], ["-"]
        for i in range(1, len(df_d)):
            diff_d = df_d.loc[i - 1, "_rate_num"] - df_d.loc[i, "_rate_num"]
            pen_d_cur = df_d.loc[i, "_daily_pen"]
            if diff_d > 0 and pen_d_cur > 0:
                gaps_pct_d.append(f"+{diff_d:.2f}%")
                gaps_amt_d.append(f"-{format_num((diff_d / 100.0) * pen_d_cur, is_currency=True)}")
            else:
                gaps_pct_d.append("0.00%")
                gaps_amt_d.append("-")
        df_d["Gap"] = gaps_pct_d
        df_d["Gap Amount"] = gaps_amt_d
        df_d["Rank"] = [f"#{i+1}" for i in range(len(df_d))]
        d_cols = ["Rank", label_entity, "Penambahan (Hari Ini)", "Penerimaan (Hari Ini)", "Recovery Rate (Hari Ini)", "Gap", "Gap Amount", "Jumlah Panggilan (Avg)", "Waktu Kerja Rata-rata (Avg)", "Jumlah SMS (Avg)", "Jumlah Pengiriman Template WABA (Avg)"]
        df_d = df_d[d_cols]

    return df_m, df_d

TARGET_BUCKETS = [
    ("S1 - SK", "S1", "SK"),
    ("D0 - SK", "D0", "SK"),
    ("S1 - LD", "S1", "LD"),
    ("S1 - LD-BA", "S1", "LD-BA"),
    ("S0 - LD", "S0", "LD"),
    ("C - LD-C", "C", "LD-C"),
    ("Semua Bucket", None, None),
]

def render_performance_with_tabs(df_source, group_col_name, label_entity):
    if df_source.empty:
        st.info("Tidak ada data.")
        return

    sub_tabs = st.tabs([label for label, _, _ in TARGET_BUCKETS])

    def render_bucket_dual_tables(df_target, title_text=None):
        if title_text:
            st.markdown(f'<div class="clean-group-title">{title_text}</div>', unsafe_allow_html=True)
        df_m, df_d = build_dual_tables(df_target, group_col_name, label_entity)
        if not df_m.empty and not df_d.empty:
            col_left, col_right = st.columns([1, 1], gap="small")
            with col_left:
                st.markdown('<div class="sub-table-header header-bulanan">Performa Bulanan (月度累计绩效)</div>', unsafe_allow_html=True)
                df_m_trans = translate_columns(df_m)
                display_full_table(df_m_trans)
            with col_right:
                st.markdown('<div class="sub-table-header header-harian">Performa Harian (今日实时绩效)</div>', unsafe_allow_html=True)
                df_d_trans = translate_columns(df_d)
                display_full_table(df_d_trans)

    for tab, (tab_label, b_filter, p_filter) in zip(sub_tabs, TARGET_BUCKETS):
        with tab:
            if b_filter and p_filter:
                filtered_df = df_source[(df_source["Bucket"] == b_filter) & (df_source["Project"] == p_filter)]
                if filtered_df.empty:
                    st.info("Tidak ada data.")
                    continue
                render_bucket_dual_tables(filtered_df)
            else:
                projects = sorted(df_source["Project"].dropna().unique().tolist())
                for p in projects:
                    p_df = df_source[df_source["Project"] == p]
                    p_buckets = sorted(p_df["Bucket"].dropna().unique().tolist())
                    for b in p_buckets:
                        b_df = p_df[p_df["Bucket"] == b]
                        if b_df.empty:
                            continue
                        render_bucket_dual_tables(b_df, f"Project: {p} &nbsp;|&nbsp; Bucket: {b} ({len(b_df)} Total Agent)")

# ---------------------------------------------------------
# WARNING LETTER HELPER
# ---------------------------------------------------------
def get_tl_wl_cutoff(total_teams):
    if total_teams <= 1:
        return 0
    elif total_teams in [2, 3, 4]:
        return 1
    elif total_teams == 5:
        return 2
    elif total_teams == 6:
        return 3
    else:
        return total_teams // 2

def calculate_warning_letter_data(df_period_all, periode_name=""):
    if df_period_all.empty:
        return pd.DataFrame(), pd.DataFrame()

    df_comp = df_period_all[df_period_all["Apakah Ikut Berkompetisi"] == "Ya"].copy()
    if df_comp.empty:
        return pd.DataFrame(), pd.DataFrame()

    wl_agents = []
    wl_tls = []

    for (p, b), grp in df_comp.groupby(["Project", "Bucket"]):
        grp_ag = grp.copy()
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
                        "Peringkat di Bucket": f"Rank #{rank_ag} dari {total_ag} Agen",
                        "Persentil Rank": f"Bottom {int(round((1 - pct_ag + (1/total_ag))*100))}%",
                        "Nama TL": row_ag.get("Nama TL", ""),
                        "ID Agent": row_ag.get("ID Agent", ""),
                        "Nama Pengguna": row_ag.get("Nama Pengguna", ""),
                        "Status Kerja": row_ag.get("Status Kerja", ""),
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
# TAB UTAMA DASHBOARD
# ---------------------------------------------------------
tab_karyawan, tab_perf_team, tab_perf_comp, tab_perf_agent, tab_final, tab_chart_final, tab_productivity, tab_wl, tab_mpt_history, tab_tl_history = st.tabs([
    "Seluruh Karyawan",
    "Performance Team",
    "Performance Perusahaan",
    "Agent Performance",
    "Data Bulanan Final",
    "📈 Grafik Performa Perusahaan",
    "🎯 Produktivitas Harian Agen",
    "Warning Letter",
    "Rekam Jejak Agen MPT05",
    "Rekam Jejak TL MPT05",
])

comp_agents = (
    df_active_performance[df_active_performance["Apakah Ikut Berkompetisi"] == "Ya"].copy()
    if not df_active_performance.empty
    else pd.DataFrame()
)

# =========================================================
# TAB 1: SELURUH KARYAWAN
# =========================================================
with tab_karyawan:
    st.caption("Database Karyawan & Status Kompetisi (员工数据库与参赛状态)")
    col_t1, _ = st.columns([1, 3])
    with col_t1:
        f_stat = st.radio("Status Kompetisi (参赛状态):", ["Semua", "Ya", "Tidak"], horizontal=True, key="f_stat_radio")
    t_view = df_karyawan_status.copy()
    if "Ikut berkompetisi atau tidak" in t_view.columns:
        t_view = t_view.drop(columns=["Ikut berkompetisi atau tidak"])
    if f_stat != "Semua":
        t_view = t_view[t_view["Apakah Ikut Berkompetisi"] == f_stat]
    t_view_trans = translate_columns(t_view)
    display_full_table(t_view_trans, len(t_view_trans))

# =========================================================
# TAB 2: PERFORMANCE TEAM
# =========================================================
with tab_perf_team:
    st.caption("Performance Team (团队组长绩效排行榜) — Tabel Bulanan & Harian")
    render_performance_with_tabs(comp_agents, "Nama TL", "Nama TL")

# =========================================================
# TAB 3: PERFORMANCE PERUSAHAAN
# =========================================================
with tab_perf_comp:
    st.caption("Performance Perusahaan (各催收机构绩效排行榜) — Tabel Bulanan & Harian")
    render_performance_with_tabs(comp_agents, "Perusahaan", "Perusahaan")

# =========================================================
# TAB 4: AGENT PERFORMANCE
# =========================================================
with tab_perf_agent:
    st.caption("Agent Performance (全员坐席绩效排行榜) — Tabel Bulanan & Harian")
    if comp_agents.empty:
        st.info("Belum ada data agen.")
    else:
        ag_sub_tabs = st.tabs([label for label, _, _ in TARGET_BUCKETS])

        def render_agent_dual_tables(target_df, title_text=None):
            if title_text:
                st.markdown(f'<div class="clean-group-title">{title_text}</div>', unsafe_allow_html=True)

            rows_m = []
            rows_d = []
            for _, r in target_df.iterrows():
                st_kerja = r.get("Status Kerja", "Aktif")
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
                    "_tot_pen": pen_m,
                    "Total Penambahan (Monthly)": format_num(pen_m, is_currency=True),
                    "Total Penerimaan (Monthly)": format_num(rec_m, is_currency=True),
                    "_rate_num": rate_m,
                    "Recovery Rate (%)": f"{rate_m:.2f}%",
                })

                pen_d = r["Total Penambahan Hari Ini_Daily"]
                rec_d = r["Total Penerimaan Hari Ini_Daily"]
                rate_d = (rec_d / pen_d * 100) if pen_d > 0 else 0.0

                rows_d.append({
                    "ID Agent": r.get("ID Agent", ""),
                    "Nama Pengguna": r.get("Nama Pengguna", ""),
                    "Status Kerja": st_kerja,
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
                cols_am = ["Rank", "Ranking Interval", "Project", "Bucket", "Nama TL", "ID Agent", "Nama Pengguna", "Status Kerja", "Total Penambahan (Monthly)", "Total Penerimaan (Monthly)", "Recovery Rate (%)", "Gap", "Gap Amount"]
                df_am = df_am[cols_am]

            df_ad = pd.DataFrame(rows_d)
            if not df_ad.empty:
                df_ad = df_ad.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
                total_nd = len(df_ad)
                gaps_d, gaps_amt_d = ["-"], ["-"]
                for i in range(1, total_nd):
                    diff_d = df_ad.loc[i - 1, "_rate_num"] - df_ad.loc[i, "_rate_num"]
                    pen_dc = df_ad.loc[i, "_daily_pen"]
                    if diff_d > 0 and pen_dc > 0:
                        gaps_d.append(f"+{diff_d:.2f}%")
                        gaps_amt_d.append(f"-{format_num((diff_d / 100.0) * pen_dc, is_currency=True)}")
                    else:
                        gaps_d.append("0.00%")
                        gaps_amt_d.append("-")

                df_ad["Gap"] = gaps_d
                df_ad["Gap Amount"] = gaps_amt_d
                df_ad["Rank"] = [f"#{i+1}" for i in range(total_nd)]
                cols_ad = ["Rank", "ID Agent", "Nama Pengguna", "Status Kerja", "Penambahan (Hari Ini)", "Penerimaan (Hari Ini)", "Recovery Rate (Hari Ini)", "Gap", "Gap Amount", "Jumlah Panggilan", "Waktu Kerja Rata-rata", "Jumlah SMS", "Jumlah Pengiriman Template WABA"]
                df_ad = df_ad[cols_ad]

            col_left, col_right = st.columns([1, 1], gap="small")
            with col_left:
                st.markdown('<div class="sub-table-header header-bulanan">Agent Bulanan (月度坐席绩效)</div>', unsafe_allow_html=True)
                display_full_table(translate_columns(df_am))
            with col_right:
                st.markdown('<div class="sub-table-header header-harian">Agent Harian (今日实时绩效)</div>', unsafe_allow_html=True)
                display_full_table(translate_columns(df_ad))

        for ag_tab, (ag_label, b_filter, p_filter) in zip(ag_sub_tabs, TARGET_BUCKETS):
            with ag_tab:
                if b_filter and p_filter:
                    view_df = comp_agents[(comp_agents["Bucket"] == b_filter) & (comp_agents["Project"] == p_filter)].copy()
                    if view_df.empty:
                        st.info("Tidak ada data.")
                        continue
                    render_agent_dual_tables(view_df)
                else:
                    projects = sorted(comp_agents["Project"].dropna().unique().tolist())
                    for p in projects:
                        p_df = comp_agents[comp_agents["Project"] == p]
                        p_buckets = sorted(p_df["Bucket"].dropna().unique().tolist())
                        for b in p_buckets:
                            b_df = p_df[p_df["Bucket"] == b]
                            if b_df.empty:
                                continue
                            render_agent_dual_tables(b_df, f"Project: {p} &nbsp;|&nbsp; Bucket: {b} ({len(b_df)} Total Agent)")

# =========================================================
# TAB 5: DATA BULANAN FINAL
# =========================================================
all_final_periods = sorted(final_dict.keys())
periods_data = {}

for p_code in all_final_periods:
    month_label = f"Bulan {p_code[:2]}/{p_code[2:]} ({p_code[2:]}年{int(p_code[:2])}月份)"
    df_p_proc = process_performance_data(final_dict[p_code], [], periode_label=p_code)
    if not df_p_proc.empty:
        df_p_comp = df_p_proc[df_p_proc["Apakah Ikut Berkompetisi"] == "Ya"].copy()
        periods_data[month_label] = df_p_comp

if not comp_agents.empty:
    periods_data["MTD Berjalan (当月至今)"] = comp_agents.copy()

with tab_final:
    st.caption("Perbandingan Recovery Rate dan Ranking antar Perusahaan bersebelahan (Bulan 7, 8, dan 9 MTD).")

    if not periods_data:
        st.info("Belum ada file bulanan final.")
    else:
        def get_compact_company_table(df_subset):
            rows = []
            for comp_name, grp in df_subset.groupby("Perusahaan"):
                pen = grp["Total Penambahan Hari Ini_Total"].sum()
                rec = grp["Total Penerimaan Hari Ini_Total"].sum()
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
            if df_res.empty:
                return pd.DataFrame()
            df_res = df_res.sort_values(by="_rate_num", ascending=False).reset_index(drop=True)
            df_res["Ranking"] = [f"#{i+1}" for i in range(len(df_res))]
            return df_res[["Project", "Bucket", "Perusahaan", "Recovery Rate", "Ranking"]]

        target_groups = [("SK", "D0"), ("SK", "S1"), ("LD", "S0"), ("LD", "S1"), ("LD-BA", "S1"), ("LD-C", "C")]
        period_keys = list(periods_data.keys())

        for proj, buck in target_groups:
            has_data = any(
                not periods_data[k][(periods_data[k]["Project"] == proj) & (periods_data[k]["Bucket"] == buck)].empty
                for k in period_keys
            )
            if not has_data:
                continue

            st.markdown(
                f'<div class="clean-group-title">'
                f'Project (项目): {proj} &nbsp;|&nbsp; Bucket (阶段): {buck}'
                f'</div>', 
                unsafe_allow_html=True
            )
            cols = st.columns(len(period_keys))

            for idx, p_label in enumerate(period_keys):
                with cols[idx]:
                    st.markdown(
                        f'<div class="month-header-clean">{p_label}</div>', 
                        unsafe_allow_html=True
                    )
                    df_period_all = periods_data[p_label]
                    df_p_b = df_period_all[(df_period_all["Project"] == proj) & (df_period_all["Bucket"] == buck)]

                    if df_p_b.empty:
                        st.caption("Tidak ada data (暂无数据)")
                    else:
                        tbl_comp = get_compact_company_table(df_p_b)
                        if not tbl_comp.empty:
                            tbl_comp_trans = translate_columns(tbl_comp)
                            display_full_table(tbl_comp_trans)
            st.write("")

# =========================================================
# TAB 6: GRAFIK PERFORMA PERUSAHAAN (SESUAI IMAGE 1)
# =========================================================
with tab_chart_final:
    st.caption("Visualisasi Tren Kenaikan & Penurunan Recovery Rate Antar Perusahaan per Divisi/Bucket (月度催收机构趋势图)")

    if not periods_data:
        st.info("Belum ada data bulanan final untuk ditampilkan dalam grafik.")
    else:
        target_groups = [("SK", "D0"), ("SK", "S1"), ("LD", "S0"), ("LD", "S1"), ("LD-BA", "S1"), ("LD-C", "C")]
        period_keys = list(periods_data.keys())

        short_labels = []
        for pk in period_keys:
            if "MTD" in pk:
                short_labels.append("MTD Berjalan")
            else:
                m_match = re.search(r"(\d{2})", pk)
                short_labels.append(f"Bln {m_match.group(1)}" if m_match else pk)

        line_colors = ["#2563eb", "#ea580c", "#16a34a", "#dc2626", "#9333ea", "#0891b2", "#d97706"]

        for proj, buck in target_groups:
            has_data = any(
                not periods_data[k][(periods_data[k]["Project"] == proj) & (periods_data[k]["Bucket"] == buck)].empty
                for k in period_keys
            )
            if not has_data:
                continue

            all_companies = set()
            for pk in period_keys:
                df_slice = periods_data[pk][(periods_data[pk]["Project"] == proj) & (periods_data[pk]["Bucket"] == buck)]
                for comp in df_slice["Perusahaan"].dropna().unique():
                    if comp and comp != "Non-Kompetisi":
                        all_companies.add(comp)

            sorted_companies = sorted(list(all_companies))
            if not sorted_companies:
                continue

            company_rates = {c: [] for c in sorted_companies}
            for c in sorted_companies:
                for pk in period_keys:
                    df_slice = periods_data[pk][(periods_data[pk]["Project"] == proj) & (periods_data[pk]["Bucket"] == buck) & (periods_data[pk]["Perusahaan"] == c)]
                    if not df_slice.empty:
                        pen = df_slice["Total Penambahan Hari Ini_Total"].sum()
                        rec = df_slice["Total Penerimaan Hari Ini_Total"].sum()
                        rate = (rec / pen * 100.0) if pen > 0 else 0.0
                    else:
                        rate = 0.0
                    company_rates[c].append(rate)

            fig, ax = plt.subplots(figsize=(9, 4.8), dpi=130)
            table_cells = []
            row_labels = []

            for idx, c in enumerate(sorted_companies):
                vals = company_rates[c]
                clr = line_colors[idx % len(line_colors)]
                
                ax.plot(short_labels, vals, marker="o", markersize=6, linewidth=2.8, label=c, color=clr)

                for x_idx, y_val in enumerate(vals):
                    if y_val > 0:
                        ax.text(
                            x_idx,
                            y_val + 0.35,
                            f"{y_val:.2f}%".replace(".", ","),
                            ha="center",
                            va="bottom",
                            fontsize=9,
                            fontweight="bold",
                            color="#1e293b",
                        )

                table_cells.append([f"{v:.2f}%".replace(".", ",") if v > 0 else "-" for v in vals])
                row_labels.append(c)

            chart_title = f"{proj} - {buck}"
            ax.set_title(chart_title, fontsize=13, fontweight="bold", pad=14, color="#0f172a")
            ax.set_ylabel("Recovery Rate", fontsize=10, fontweight="bold")
            ax.grid(axis="y", linestyle="--", alpha=0.5, color="#94a3b8")
            ax.tick_params(axis="x", labelsize=10)

            all_vals = [v for vals in company_rates.values() for v in vals if v > 0]
            if all_vals:
                min_y = max(0, min(all_vals) - 4)
                max_y = max(all_vals) + 5
                ax.set_ylim(min_y, max_y)

            cell_table = ax.table(
                cellText=table_cells,
                rowLabels=row_labels,
                colLabels=short_labels,
                loc="bottom",
                bbox=[0.0, -0.42, 1.0, 0.30]
            )
            cell_table.auto_set_font_size(False)
            cell_table.set_fontsize(8.5)
            cell_table.scale(1, 1.25)
            plt.subplots_adjust(bottom=0.32)

            st.pyplot(fig)
            plt.close(fig)
            st.markdown("---")

# =========================================================
# TAB 7: PRODUKTIVITAS HARIAN AGEN (MPT05 DAILY PRODUCTIVITY)
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
            input_pwd = st.text_input("Masukkan Password Admin:", type="password", key="input_target_pwd")
        with col_btn_lock:
            st.write("")
            if not st.session_state["target_unlocked"]:
                if st.button("🔓 Buka Kunci Edit", use_container_width=True):
                    if input_pwd == ADMIN_TARGET_PASSWORD:
                        st.session_state["target_unlocked"] = True
                        st.success("Akses edit terbuka!")
                        st.rerun()
                    else:
                        st.error("Password salah!")
            else:
                if st.button("🔒 Kunci Kembali", use_container_width=True):
                    st.session_state["target_unlocked"] = False
                    st.info("Pengaturan dikunci kembali.")
                    st.rerun()

        is_editable = st.session_state["target_unlocked"]
        if not is_editable:
            st.warning("⚠️ Form target terkunci (Read-Only). Masukkan password di atas untuk mengubah nilai target.")
        else:
            st.success("✅ Mode Edit Aktif. Anda dapat mengubah nilai target di bawah ini.")

        target_configs = {}
        for b_name in TARGET_KEYS:
            st.markdown(f"**Divisi: `{b_name}`**")
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            with c1:
                t_call = st.number_input("Target Call", min_value=0, value=300, step=10, key=f"t_call_{b_name}", disabled=not is_editable)
            with c2:
                t_work = st.number_input("Waktu Kerja (Jam)", min_value=0.0, value=8.0, step=0.5, key=f"t_work_{b_name}", disabled=not is_editable)
            with c3:
                t_cov_call = st.number_input("Tutup Call (%)", min_value=0.0, max_value=100.0, value=80.0, step=5.0, key=f"t_covcall_{b_name}", disabled=not is_editable)
            with c4:
                t_cov_sms = st.number_input("Tutup SMS (%)", min_value=0.0, max_value=200.0, value=95.0, step=5.0, key=f"t_covsms_{b_name}", disabled=not is_editable)
            with c5:
                t_waba = st.number_input("Target WABA", min_value=0, value=0, step=5, key=f"t_waba_{b_name}", disabled=not is_editable)
            with c6:
                t_rec = st.number_input("Target Rec Rate (%)", min_value=0.0, max_value=100.0, value=15.0, step=1.0, key=f"t_rec_{b_name}", disabled=not is_editable)
                
            target_configs[b_name] = {
                "call": t_call,
                "work": t_work,
                "cov_call": t_cov_call,
                "cov_sms": t_cov_sms,
                "waba": t_waba,
                "rec_rate": t_rec
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
                proj_detected = "SK" if "SK" in fname else ("LD" if "LD" in fname else "")
                df_temp["_source_project"] = proj_detected
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

                id_clean = normalize_clean(id_raw)
                nm_clean = normalize_clean(nm_raw)

                k_info = karyawan_dict.get((id_clean, proj_src)) or karyawan_dict.get((nm_clean, proj_src)) or karyawan_dict.get(id_clean) or karyawan_dict.get(nm_clean) or {}
                
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

                jml_calls = clean_numeric(r.get("Jumlah Panggilan", 0))
                work_time = clean_numeric(r.get("Waktu Kerja Rata-rata", 0))
                cov_call = clean_numeric(r.get("Tingkat Penutupan Panggilan", 0))
                cov_sms = clean_numeric(r.get("Tingkat Penutupan SMS", 0))
                waba_sent = clean_numeric(r.get("Jumlah Pengiriman Template WABA", 0))

                agent_productivity_list.append({
                    "Project": proj_name,
                    "Bucket": bucket_name,
                    "Nama TL": tl_name,
                    "ID Agent": k_info.get("ID_Akun_Resmi", id_raw),
                    "Nama Pengguna": k_info.get("Nama_Orang_Resmi", nm_raw),
                    "Status Kerja": k_info.get("Status Kerja", "Aktif"),
                    "Calls": jml_calls,
                    "Work_Time": work_time,
                    "Cov_Call": cov_call,
                    "Cov_SMS": cov_sms,
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

                    tg = target_configs.get(div_key, {"call": 300, "work": 8.0, "cov_call": 80.0, "cov_sms": 95.0, "waba": 0, "rec_rate": 15.0})

                    team_header = (
                        f'<div style="background-color: #1e3a1e; color: #86efac; padding: 7px 14px; '
                        f'border-radius: 6px; font-weight: bold; font-size: 1rem; margin-top: 15px; margin-bottom: 8px;">'
                        f'👥 Team: {tl_team} &nbsp;|&nbsp; Divisi: {div_key} &nbsp;'
                        f'(Target Call: {tg["call"]} | Waktu: {tg["work"]} Jam | SMS: {tg["cov_sms"]}%)'
                        f'</div>'
                    )
                    st.markdown(team_header, unsafe_allow_html=True)

                    html_tbl = [
                        '<div style="width: 100%; overflow-x: auto; margin-bottom: 1.2rem; border: 1px solid #334155; border-radius: 4px;">'
                        '<table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: -apple-system, sans-serif; font-size: 0.82rem; text-align: center;">'
                        '<thead><tr>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px; border-right: 1px solid #047857;">No</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px; border-right: 1px solid #047857;">ID Agent</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px; border-right: 1px solid #047857;">Nama Agent</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px; border-right: 1px solid #047857;">Calls (Panggilan)</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px; border-right: 1px solid #047857;">Waktu Kerja (Jam)</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px; border-right: 1px solid #047857;">Tutup Call (%)</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px; border-right: 1px solid #047857;">Tutup SMS (%)</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px; border-right: 1px solid #047857;">Template WABA</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px; border-right: 1px solid #047857;">Penambahan</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px; border-right: 1px solid #047857;">Penerimaan</th>'
                        '<th style="background-color: #064e3b; color: #ecfdf5; padding: 7px 4px;">Recovery Rate</th>'
                        '</tr></thead><tbody>'
                    ]

                    group_sorted = group.sort_values(by="Rec_Rate", ascending=False).reset_index(drop=True)

                    for idx, row in group_sorted.iterrows():
                        c_call = "color: #ef4444; font-weight: bold;" if row["Calls"] < tg["call"] else "color: inherit;"
                        c_work = "color: #ef4444; font-weight: bold;" if row["Work_Time"] < tg["work"] else "color: inherit;"
                        c_covcall = "color: #ef4444; font-weight: bold;" if row["Cov_Call"] < tg["cov_call"] else "color: inherit;"
                        c_covsms = "color: #ef4444; font-weight: bold;" if row["Cov_SMS"] < tg["cov_sms"] else "color: inherit;"
                        c_waba = "color: #ef4444; font-weight: bold;" if row["WABA"] < tg["waba"] else "color: inherit;"
                        c_rec = "color: #ef4444; font-weight: bold;" if row["Rec_Rate"] < tg["rec_rate"] else "color: inherit;"

                        row_html = (
                            f'<tr style="border-bottom: 1px solid #1e293b;">'
                            f'<td style="padding: 5px 4px; border-right: 1px solid #334155;">{idx+1}</td>'
                            f'<td style="padding: 5px 4px; border-right: 1px solid #334155; font-weight: 600;">{row["ID Agent"]}</td>'
                            f'<td style="padding: 5px 4px; border-right: 1px solid #334155;">{row["Nama Pengguna"]}</td>'
                            f'<td style="padding: 5px 4px; border-right: 1px solid #334155; {c_call}">{int(row["Calls"]):,}</td>'
                            f'<td style="padding: 5px 4px; border-right: 1px solid #334155; {c_work}">{row["Work_Time"]:.2f}</td>'
                            f'<td style="padding: 5px 4px; border-right: 1px solid #334155; {c_covcall}">{row["Cov_Call"]:.2f}%</td>'
                            f'<td style="padding: 5px 4px; border-right: 1px solid #334155; {c_covsms}">{row["Cov_SMS"]:.2f}%</td>'
                            f'<td style="padding: 5px 4px; border-right: 1px solid #334155; {c_waba}">{int(row["WABA"]):,}</td>'
                            f'<td style="padding: 5px 4px; border-right: 1px solid #334155;">{format_num(row["Penambahan"], is_currency=True)}</td>'
                            f'<td style="padding: 5px 4px; border-right: 1px solid #334155;">{format_num(row["Penerimaan"], is_currency=True)}</td>'
                            f'<td style="padding: 5px 4px; {c_rec}">{row["Rec_Rate"]:.2f}%</td>'
                            f'</tr>'
                        )
                        html_tbl.append(row_html)

                    tot_pen = group_sorted["Penambahan"].sum()
                    tot_rec = group_sorted["Penerimaan"].sum()
                    avg_rate = (tot_rec / tot_pen * 100.0) if tot_pen > 0 else 0.0
                    avg_calls = group_sorted["Calls"].mean()
                    avg_work = group_sorted["Work_Time"].mean()
                    avg_sms = group_sorted["Cov_SMS"].mean()
                    avg_covcall = group_sorted["Cov_Call"].mean()
                    tot_waba = group_sorted["WABA"].sum()

                    total_html = (
                        f'<tr style="background-color: #064e3b; color: #ecfdf5; font-weight: bold; border-top: 2px solid #059669;">'
                        f'<td colspan="3" style="padding: 6px 4px; text-align: center; border-right: 1px solid #047857;">TOTAL / RATA-RATA TIM ({len(group_sorted)} AGENT)</td>'
                        f'<td style="padding: 6px 4px; border-right: 1px solid #047857;">{int(avg_calls):,}</td>'
                        f'<td style="padding: 6px 4px; border-right: 1px solid #047857;">{avg_work:.2f}</td>'
                        f'<td style="padding: 6px 4px; border-right: 1px solid #047857;">{avg_covcall:.2f}%</td>'
                        f'<td style="padding: 6px 4px; border-right: 1px solid #047857;">{avg_sms:.2f}%</td>'
                        f'<td style="padding: 6px 4px; border-right: 1px solid #047857;">{int(tot_waba):,}</td>'
                        f'<td style="padding: 6px 4px; border-right: 1px solid #047857;">{format_num(tot_pen, is_currency=True)}</td>'
                        f'<td style="padding: 6px 4px; border-right: 1px solid #047857;">{format_num(tot_rec, is_currency=True)}</td>'
                        f'<td style="padding: 6px 4px;">{avg_rate:.2f}%</td>'
                        f'</tr>'
                    )
                    html_tbl.append(total_html)
                    html_tbl.append('</tbody></table></div>')

                    st.markdown("".join(html_tbl), unsafe_allow_html=True)

# =========================================================
# TAB 8: WARNING LETTER (DENGAN BUTTON DOWNLOAD EXCEL)
# =========================================================
with tab_wl:
    st.caption("Daftar Warning Letter untuk Agen dan Team Leader (警告信名单)")

    if not wl_dict:
        st.info("Belum ada file mingguan Warning Letter di folder `rawdata` (contoh: `1-7-SK.xlsx`, `8-14-LD.xlsx`).")
    else:
        import io

        wl_period_keys = sorted(wl_dict.keys(), key=lambda x: [int(n) for n in re.findall(r"\d+", x)] or [0])
        
        # 1. Pilihan Periode Evaluasi & Tombol Download
        col_w1, col_w2, col_w3 = st.columns([2, 2.5, 1.5])
        with col_w1:
            selected_wl_period = st.selectbox("Pilih Periode Mingguan (Agent WL):", wl_period_keys, key="sel_wl_period")
        
        files_for_period = wl_dict.get(selected_wl_period, [])
        df_wl_raw = process_performance_data(files_for_period, [])
        df_wl_ag, _ = calculate_warning_letter_data(df_wl_raw, f"Periode {selected_wl_period}")

        df_wl_tl = pd.DataFrame()
        label_tl_period = selected_wl_period
        if len(wl_period_keys) >= 2:
            p_idx = wl_period_keys.index(selected_wl_period)
            pair_start = (p_idx // 2) * 2
            pair_keys = wl_period_keys[pair_start : pair_start + 2]
            
            combined_tl_files = []
            for pk in pair_keys:
                combined_tl_files.extend(wl_dict.get(pk, []))
            
            label_tl_period = " & ".join(pair_keys)
            with col_w2:
                st.info(f"Evaluasi TL 2 Mingguan: Menggabungkan Periode **{label_tl_period}**")
            
            df_wl_tl_raw = process_performance_data(combined_tl_files, [])
            _, df_wl_tl = calculate_warning_letter_data(df_wl_tl_raw, f"2-Minggu ({label_tl_period})")
        else:
            with col_w2:
                st.caption("ℹ️ Warning Letter TL dievaluasi setiap 2 minggu")

        # ---------------------------------------------------------
        # 2. FILTER HANYA KOLOM YANG TERTAMPIL DI TABEL LAYAR
        # ---------------------------------------------------------
        cols_display_ag = [
            "Project", "Bucket", "Rank", "ID Agent", "Nama Pengguna", 
            "Nama TL", "Peringkat di Bucket", "Total Penambahan (Periode)", 
            "Total Penerimaan (Periode)", "Recovery Rate (%)", "Status WL"
        ]
        cols_display_tl = [
            "Project", "Bucket", "Rank", "Nama TL", "Perusahaan", 
            "Peringkat di Bucket", "Total Agent", "Total Penambahan (Periode)", 
            "Total Penerimaan (Periode)", "Recovery Rate (%)", "Status WL"
        ]

        # Ambil hanya kolom yang tertampil
        df_export_ag = pd.DataFrame()
        if df_wl_ag is not None and not df_wl_ag.empty:
            exist_ag = [c for c in cols_display_ag if c in df_wl_ag.columns]
            df_export_ag = df_wl_ag[exist_ag].rename(columns={"Peringkat di Bucket": "Rank / Total"})

        df_export_tl = pd.DataFrame()
        if df_wl_tl is not None and not df_wl_tl.empty:
            exist_tl = [c for c in cols_display_tl if c in df_wl_tl.columns]
            df_export_tl = df_wl_tl[exist_tl].rename(columns={"Peringkat di Bucket": "Rank / Total"})

        # Buat file Excel (.xlsx) di memory buffer
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            if not df_export_ag.empty:
                df_export_ag.to_excel(writer, sheet_name="WL_Agent_MPT05", index=False)
            else:
                pd.DataFrame({"Keterangan": ["Tidak ada agen MPT05 yang terkena WL"]}).to_excel(writer, sheet_name="WL_Agent_MPT05", index=False)

            if not df_export_tl.empty:
                df_export_tl.to_excel(writer, sheet_name="WL_TL_MPT05", index=False)
            else:
                pd.DataFrame({"Keterangan": ["Tidak ada TL MPT05 yang terkena WL"]}).to_excel(writer, sheet_name="WL_TL_MPT05", index=False)

        # Tombol Download Excel
        with col_w3:
            st.write("")  # Penyeimbang margin vertikal agar sejajar
            st.download_button(
                label="📥 Unduh Data WL (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=f"Warning_Letter_MPT05_Periode_{selected_wl_period}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # ---------------------------------------------------------
        # 3. TAMPILAN TABEL SUB-TAB SESUAI DIVISI / BUCKET
        # ---------------------------------------------------------
        wl_sub_tabs = st.tabs([label for label, _, _ in TARGET_BUCKETS])

        def render_wl_tables(df_target_ag, df_target_tl, title_text=None):
            if title_text:
                st.markdown(f'<div class="clean-group-title">{title_text}</div>', unsafe_allow_html=True)

            col_l, col_r = st.columns([1, 1], gap="small")
            with col_l:
                st.markdown('<div class="sub-table-header header-wl">⚠️ Warning Letter Agent MPT05 (Bottom 30%)</div>', unsafe_allow_html=True)
                if df_target_ag is None or df_target_ag.empty:
                    st.success("🎉 Tidak ada agen MPT05 yang terkena Warning Letter di bucket ini!")
                else:
                    cols_ag = ["Rank", "ID Agent", "Nama Pengguna", "Nama TL", "Peringkat di Bucket", "Total Penambahan (Periode)", "Total Penerimaan (Periode)", "Recovery Rate (%)", "Status WL"]
                    cols_exist = [c for c in cols_ag if c in df_target_ag.columns]
                    ag_view = df_target_ag[cols_exist]
                    display_full_table(translate_columns(ag_view))

            with col_r:
                st.markdown('<div class="sub-table-header header-wl">🚨 Warning Letter TL MPT05 (Evaluasi 2 Minggu)</div>', unsafe_allow_html=True)
                if df_target_tl is None or df_target_tl.empty:
                    st.success("🎉 Tidak ada Team Leader MPT05 yang terkena Warning Letter di bucket ini!")
                else:
                    cols_tl = ["Rank", "Nama TL", "Perusahaan", "Peringkat di Bucket", "Total Agent", "Total Penambahan (Periode)", "Total Penerimaan (Periode)", "Recovery Rate (%)", "Status WL"]
                    cols_exist_tl = [c for c in cols_tl if c in df_target_tl.columns]
                    tl_view = df_target_tl[cols_exist_tl]
                    display_full_table(translate_columns(tl_view))

        for tab_w, (w_label, b_filt, p_filt) in zip(wl_sub_tabs, TARGET_BUCKETS):
            with tab_w:
                if b_filt and p_filt:
                    ag_b = df_wl_ag[(df_wl_ag["Bucket"] == b_filt) & (df_wl_ag["Project"] == p_filt)] if (df_wl_ag is not None and not df_wl_ag.empty) else pd.DataFrame()
                    tl_b = df_wl_tl[(df_wl_tl["Bucket"] == b_filt) & (df_wl_tl["Project"] == p_filt)] if (df_wl_tl is not None and not df_wl_tl.empty) else pd.DataFrame()
                    render_wl_tables(ag_b, tl_b)
                else:
                    projects_ag = df_wl_ag["Project"].tolist() if (df_wl_ag is not None and not df_wl_ag.empty) else []
                    projects_tl = df_wl_tl["Project"].tolist() if (df_wl_tl is not None and not df_wl_tl.empty) else []
                    all_p = sorted(list(set(projects_ag + projects_tl)))
                    
                    if not all_p:
                        st.info("Belum ada data Warning Letter.")
                    for p_item in all_p:
                        ag_p = df_wl_ag[df_wl_ag["Project"] == p_item] if (df_wl_ag is not None and not df_wl_ag.empty) else pd.DataFrame()
                        tl_p = df_wl_tl[df_wl_tl["Project"] == p_item] if (df_wl_tl is not None and not df_wl_tl.empty) else pd.DataFrame()
                        
                        buckets_ag = ag_p["Bucket"].tolist() if not ag_p.empty else []
                        buckets_tl = tl_p["Bucket"].tolist() if not tl_p.empty else []
                        b_list = sorted(list(set(buckets_ag + buckets_tl)))
                        
                        for b_item in b_list:
                            ag_pb = ag_p[ag_p["Bucket"] == b_item] if not ag_p.empty else pd.DataFrame()
                            tl_pb = tl_p[tl_p["Bucket"] == b_item] if not tl_p.empty else pd.DataFrame()
                            render_wl_tables(ag_pb, tl_pb, f"Project: {p_item} &nbsp;|&nbsp; Bucket: {b_item}")

# =========================================================
# HELPER DATA HISTORIS (UNTUK AGENT & TL)
# =========================================================
all_period_records = {}
for p_code in all_final_periods:
    m_label = f"Bln {p_code[:2]}"
    df_p_proc = process_performance_data(final_dict[p_code], [], periode_label=p_code)
    if not df_p_proc.empty:
        all_period_records[m_label] = df_p_proc[df_p_proc["Apakah Ikut Berkompetisi"] == "Ya"].copy()

if not comp_agents.empty:
    cur_month_str = datetime.now().strftime("%m")
    all_period_records[f"Bln {cur_month_str} (MTD)"] = comp_agents.copy()

RANK_ORDER = {
    "Konsisten TOP 🌟": 1,
    "Terus Meningkat 📈": 2,
    "Fluktuatif 🔄": 3,
    "Terus Menurun 📉": 4,
    "Konsisten BOTTOM ⚠️": 5,
    "Data Baru 🆕": 6,
}

# =========================================================
# TAB 9: REKAM JEJAK HISTORIS AGEN MPT05
# =========================================================
with tab_mpt_history:
    st.caption("Memantau kinerja agen MPT05. Konsisten TOP HANYA jika di SEMUA bulan masuk dalam Top 30% bucket masing-masing.")

    if not all_period_records:
        st.info("Belum ada data historis yang dapat diproses.")
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
                        nm_ag = str(ag.get("Nama Pengguna", "")).strip()
                        st_kerja = ag.get("Status Kerja", "Aktif")

                        if id_ag not in mpt_agents_dict:
                            mpt_agents_dict[id_ag] = {
                                "ID Agent": id_ag,
                                "Nama Pengguna": nm_ag,
                                "Status Kerja": st_kerja,
                                "rates": [],
                                "is_top_30_list": [],
                                "is_bottom_30_list": [],
                                "history": {},
                            }

                        rec_rate_val = ag["_rate"]
                        actual_rank = idx_rank + 1
                        
                        pct_in_bucket = actual_rank / total_in_bucket if total_in_bucket > 0 else 1.0
                        is_top30 = pct_in_bucket <= 0.30
                        is_bottom30 = pct_in_bucket > 0.70

                        mpt_agents_dict[id_ag]["rates"].append(rec_rate_val)
                        mpt_agents_dict[id_ag]["is_top_30_list"].append(is_top30)
                        mpt_agents_dict[id_ag]["is_bottom_30_list"].append(is_bottom30)
                        
                        txt_summary = f"{rec_rate_val:.2f}% (Rank #{actual_rank}/{total_in_bucket}) | {p}-{b} | {ag.get('Nama TL', '-')}"
                        mpt_agents_dict[id_ag]["history"][p_label] = txt_summary

        history_rows = []
        p_columns = list(all_period_records.keys())

        for id_ag, data in mpt_agents_dict.items():
            avg_rate = sum(data["rates"]) / len(data["rates"]) if data["rates"] else 0.0
            total_active_months = len(data["rates"])

            if total_active_months <= 1:
                trend_label = "Data Baru 🆕"
            else:
                is_always_top_30 = all(data["is_top_30_list"])
                is_always_bottom_30 = all(data["is_bottom_30_list"])
                is_improving = all(data["rates"][i] < data["rates"][i+1] for i in range(total_active_months - 1))
                is_declining = all(data["rates"][i] > data["rates"][i+1] for i in range(total_active_months - 1))

                if is_always_top_30:
                    trend_label = "Konsisten TOP 🌟"
                elif is_always_bottom_30:
                    trend_label = "Konsisten BOTTOM ⚠️"
                elif is_improving:
                    trend_label = "Terus Meningkat 📈"
                elif is_declining:
                    trend_label = "Terus Menurun 📉"
                else:
                    trend_label = "Fluktuatif 🔄"

            row_item = {
                "ID Agent": data["ID Agent"],
                "Nama Pengguna": data["Nama Pengguna"],
                "Status Kerja": data["Status Kerja"],
                "Rata-rata Rate (All-Time)": f"{avg_rate:.2f}%",
                "Analisis Kinerja": trend_label,
                "_order_val": RANK_ORDER.get(trend_label, 9),
                "_avg_sort": avg_rate,
            }

            for p_col in p_columns:
                row_item[p_col] = data["history"].get(p_col, "-")

            history_rows.append(row_item)

        df_hist_ag = pd.DataFrame(history_rows)
        if df_hist_ag.empty:
            st.info("Tidak ada data agen MPT05.")
        else:
            df_hist_ag = df_hist_ag.sort_values(by=["_order_val", "_avg_sort"], ascending=[True, False]).reset_index(drop=True)
            df_hist_ag.insert(0, "Ranking (Avg)", [f"#{i+1}" for i in range(len(df_hist_ag))])
            df_hist_ag.drop(columns=["_order_val", "_avg_sort"], inplace=True)

            df_hist_ag_trans = translate_columns(df_hist_ag)
            display_full_table(df_hist_ag_trans, len(df_hist_ag_trans))

# =========================================================
# TAB 10: REKAM JEJAK HISTORIS TEAM LEADER (TL) MPT05
# =========================================================
with tab_tl_history:
    st.caption("Memantau kinerja seluruh Team Leader MPT05 & W-MPT05 lintas bulan (组长月度绩效跟踪).")

    if not all_period_records:
        st.info("Belum ada data historis yang dapat diproses.")
    else:
        mpt_tl_dict = {}

        for p_label, df_p in all_period_records.items():
            for (p, b), grp in df_p.groupby(["Project", "Bucket"]):
                tl_summary = []
                for tl_name, tl_grp in grp.groupby("Nama TL"):
                    pen = tl_grp["Total Penambahan Hari Ini_Total"].sum()
                    rec = tl_grp["Total Penerimaan Hari Ini_Total"].sum()
                    rate = (rec / pen * 100) if pen > 0 else 0.0
                    tl_summary.append({
                        "Nama TL": tl_name,
                        "Perusahaan": tl_grp.iloc[0].get("Perusahaan", ""),
                        "rate": rate,
                        "agent_count": len(tl_grp),
                        "has_active_agent": (tl_grp["Status Kerja"] == "Aktif").any(),
                    })

                df_tl_grp = pd.DataFrame(tl_summary)
                df_tl_grp = df_tl_grp.sort_values(by="rate", ascending=False).reset_index(drop=True)
                total_tls = len(df_tl_grp)

                for idx_rank, tl_r in df_tl_grp.iterrows():
                    tl_name = str(tl_r["Nama TL"]).strip()
                    perush_name = str(tl_r["Perusahaan"]).upper()

                    if "MPT05" in tl_name.upper() or "MPT05" in perush_name:
                        st_tl = "Aktif" if tl_r.get("has_active_agent", True) else "Resign"

                        if tl_name not in mpt_tl_dict:
                            mpt_tl_dict[tl_name] = {
                                "Nama TL": tl_name,
                                "Status TL": st_tl,
                                "rates": [],
                                "ranks_pct": [],
                                "history": {},
                            }
                        else:
                            if st_tl == "Aktif":
                                mpt_tl_dict[tl_name]["Status TL"] = "Aktif"

                        rec_rate_val = tl_r["rate"]
                        actual_rank = idx_rank + 1
                        mpt_tl_dict[tl_name]["rates"].append(rec_rate_val)
                        mpt_tl_dict[tl_name]["ranks_pct"].append(actual_rank / total_tls if total_tls > 0 else 1.0)

                        txt_tl_sum = f"{rec_rate_val:.2f}% (Rank #{actual_rank}) | {p}-{b} | {tl_r['agent_count']} Agent"
                        mpt_tl_dict[tl_name]["history"][p_label] = txt_tl_sum

        tl_history_rows = []
        p_columns = list(all_period_records.keys())

        for tl_name, data in mpt_tl_dict.items():
            avg_rate = sum(data["rates"]) / len(data["rates"]) if data["rates"] else 0.0

            if len(data["rates"]) <= 1:
                trend_label = "Data Baru 🆕"
            else:
                is_always_top = all(pct <= 0.30 for pct in data["ranks_pct"])
                is_always_bottom = all(pct >= 0.70 for pct in data["ranks_pct"])
                is_improving = all(data["rates"][i] <= data["rates"][i+1] for i in range(len(data["rates"])-1))
                is_declining = all(data["rates"][i] >= data["rates"][i+1] for i in range(len(data["rates"])-1))

                if is_always_top:
                    trend_label = "Konsisten TOP 🌟"
                elif is_always_bottom:
                    trend_label = "Konsisten BOTTOM ⚠️"
                elif is_improving:
                    trend_label = "Terus Meningkat 📈"
                elif is_declining:
                    trend_label = "Terus Menurun 📉"
                else:
                    trend_label = "Fluktuatif 🔄"

            row_item = {
                "Nama TL": data["Nama TL"],
                "Status TL": data["Status TL"],
                "Rata-rata Rate (All-Time)": f"{avg_rate:.2f}%",
                "Analisis Kinerja": trend_label,
                "_order_val": RANK_ORDER.get(trend_label, 9),
                "_avg_sort": avg_rate,
            }

            for p_col in p_columns:
                row_item[p_col] = data["history"].get(p_col, "-")

            tl_history_rows.append(row_item)

        df_hist_tl = pd.DataFrame(tl_history_rows)
        if df_hist_tl.empty:
            st.info("Tidak ada data TL MPT05.")
        else:
            df_hist_tl = df_hist_tl.sort_values(by=["_order_val", "_avg_sort"], ascending=[True, False]).reset_index(drop=True)
            df_hist_tl.insert(0, "Ranking (Avg)", [f"#{i+1}" for i in range(len(df_hist_tl))])
            df_hist_tl.drop(columns=["_order_val", "_avg_sort"], inplace=True)

            df_hist_tl_trans = translate_columns(df_hist_tl)
            display_full_table(df_hist_tl_trans, len(df_hist_tl_trans))