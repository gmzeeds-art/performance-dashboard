import os
import glob
import re
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
    # Umum & Karyawan
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
    
    # Leaderboard & Tabel
    "Rank": "Rank<br>(排名)",
    "Ranking": "Ranking<br>(排名)",
    "Ranking (Avg)": "Ranking Avg<br>(平均排名)",
    "Ranking Interval": "Ranking Interval<br>(排名区间)",
    "Nama TL": "Nama TL<br>(组长姓名)",
    "ID Agent": "ID Agent<br>(坐席工号)",
    "Total Agent": "Total Agent<br>(总坐席数)",
    
    # Metrik Keuangan
    "Total Penambahan (Monthly)": "Total Penambahan (M)<br>(月累计新增本金)",
    "Total Penerimaan (Monthly)": "Total Penerimaan (M)<br>(月累计回款总额)",
    "Recovery Rate (Monthly)": "Recovery Rate (M)<br>(月催回率)",
    "Recovery Rate (%)": "Recovery Rate<br>(催回率)",
    "Recovery Rate": "Recovery Rate<br>(催回率)",
    "Gap": "Gap<br>(差距率)",
    "Gap Amount": "Gap Amount<br>(差距金额)",
    "Penambahan (Hari Ini)": "Penambahan (Today)<br>(今日新增本金)",
    "Penerimaan (Hari Ini)": "Penerimaan (Today)<br>(今日回款总额)",
    "Recovery Rate (Hari Ini)": "Recovery Rate (Today)<br>(今日催回率)",
    
    # Metrik Operasional Harian
    "Jumlah Panggilan": "Panggilan<br>(呼叫次数)",
    "Jumlah Panggilan (Avg)": "Panggilan (Avg)<br>(平均呼叫次数)",
    "Waktu Kerja Rata-rata": "Waktu Kerja<br>(平均工作时长)",
    "Waktu Kerja Rata-rata (Avg)": "Waktu Kerja (Avg)<br>(平均工作时长)",
    "Jumlah SMS": "SMS<br>(短信发送量)",
    "Jumlah SMS (Avg)": "SMS (Avg)<br>(平均短信量)",
    "Jumlah Pengiriman Template WABA": "Template WABA<br>(WABA模板发送量)",
    "Jumlah Pengiriman Template WABA (Avg)": "Template WABA (Avg)<br>(平均WABA发送量)",
    
    # Rekam Jejak
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
# FUNGSI RENDER TABEL RESPONSIVE, AUTO WRAP & FREEZE HEADER
# ---------------------------------------------------------
def display_full_table(df_input, row_count=None):
    df = df_input.data if hasattr(df_input, "data") else df_input

    # Container dengan scroll vertikal dan horizontal (max-height membuat tabel bisa di-scroll dengan header beku)
    html = ['<div style="width: 100%; max-height: 72vh; overflow-y: auto; overflow-x: auto; margin-bottom: 0.6rem; border: 1px solid #334155; border-radius: 4px;">']
    html.append(
        '<table style="width: 100%; border-collapse: separate; border-spacing: 0; '
        'font-family: -apple-system, BlinkMacSystemFont, sans-serif; '
        'font-size: 0.81rem; text-align: center;">'
    )
    
    # Header Wrap Text + FREEZE PANES (Sticky Top)
    html.append('<thead><tr>')
    for col in df.columns:
        col_header = str(col).replace("\n", "<br>")
        html.append(
            f'<th style="position: sticky; top: 0; z-index: 5; background-color: #1e293b; color: #f8fafc; '
            f'padding: 7px 4px; font-weight: 600; white-space: normal; line-height: 1.2; vertical-align: middle; '
            f'border-bottom: 2px solid #475569; border-right: 1px solid #334155; word-break: break-word;">{col_header}</th>'
        )
    html.append('</tr></thead><tbody>')
    
    # Rows & Highlight MPT05
    for _, row in df.iterrows():
        row_str = " ".join([str(v) for v in row.values]).upper()
        is_mpt05 = "MPT05" in row_str
        is_resign = "RESIGN" in row_str
        
        if is_resign:
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
    # 1. Kategori Resign (Mengundurkan diri & Tidak aktif)
    if "mengundurkan" in s or "resign" in s or "tidak aktif" in s or "non aktif" in s:
        return "Resign"
    # 2. Kategori Aktif (Aktif, Aktif bekerja, & Istirahat)
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
# 2. SCAN & BACA FILE RAWDATA
# ---------------------------------------------------------
METRIC_SUM_COLS = ["Total Penambahan Hari Ini", "Total Penerimaan Hari Ini"]
METRIC_AVG_COLS = ["Jumlah Panggilan", "Waktu Kerja Rata-rata", "Jumlah SMS", "Jumlah Pengiriman Template WABA"]
ALL_METRICS = METRIC_SUM_COLS + METRIC_AVG_COLS


def read_excel_standard(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = [str(c).strip() for c in df.columns]
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
    mtd_files = []
    daily_files = []

    for f in all_files:
        clean_name = os.path.basename(f).upper()
        m_final = re.match(r"^(\d{6})[-_].*", clean_name)
        if m_final:
            periode = m_final.group(1)
            final_files.setdefault(periode, []).append(f)
        elif "MTD" in clean_name:
            mtd_files.append(f)
        else:
            daily_files.append(f)

    return final_files, mtd_files, daily_files


# ---------------------------------------------------------
# PROSES UTAMA
# ---------------------------------------------------------
df_karyawan, df_master = load_database_and_master(MASTER_FILE)
final_dict, mtd_files, daily_files = scan_rawdata_files(RAWDATA_DIR)

col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown("### **Performance Collection Automation (催收绩效自动化系统)**")
with col_h2:
    if st.button("Refresh Data (刷新数据)", use_container_width=True):
        st.rerun()

if df_karyawan is None:
    st.error("File `data_karyawan.xlsx` tidak ditemukan.")
    st.stop()

tl_col = "Nama TL" if "Nama TL" in df_karyawan.columns else "Perusahaan"
bucket_col = "Bucket" if "Bucket" in df_karyawan.columns else "Bucket"
project_col = "Project" if "Project" in df_karyawan.columns else "Project"

# ---------------------------------------------------------
# MAPPING DATABASE KARYAWAN DENGAN SISTEM SILANG (CROSS-MATCH)
# ---------------------------------------------------------
karyawan_dict = {}
for _, row in df_karyawan.iterrows():
    # Kolom A di data_karyawan adalah ID Akun (Login)
    id_akun = normalize_clean(row.get("ID Pengguna", ""))
    # Kolom B di data_karyawan adalah Nama Asli Orang
    nama_orang = normalize_clean(row.get("Nama Pengguna", ""))
    raw_status = row.get("Status Kerja", "")

    info = {
        "Nama TL": str(row.get(tl_col, "")).strip(),
        "Bucket": str(row.get(bucket_col, "")).strip(),
        "Project": str(row.get(project_col, "")).strip(),
        "Status Kerja": map_status_kerja(raw_status),
        "ID_Akun_Resmi": str(row.get("ID Pengguna", "")).strip(),
        "Nama_Orang_Resmi": str(row.get("Nama Pengguna", "")).strip(),
    }

    # Daftarkan kedua kunci ke kamus pencarian
    if id_akun:
        karyawan_dict[id_akun] = info
    if nama_orang:
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


def process_performance_data(mtd_file_list, daily_file_list):
    if not mtd_file_list and not daily_file_list:
        return pd.DataFrame()

    mtd_dfs = [read_excel_standard(f) for f in mtd_file_list]
    df_mtd_all = pd.concat([d for d in mtd_dfs if not d.empty], ignore_index=True) if mtd_dfs else pd.DataFrame()

    daily_dfs = [read_excel_standard(f) for f in daily_file_list]
    df_daily_all = pd.concat([d for d in daily_dfs if not d.empty], ignore_index=True) if daily_dfs else pd.DataFrame()

    # 1. Peta data harian (Key: ID Akun Login & Nama Orang)
    daily_map = {}
    if not df_daily_all.empty:
        for _, r in df_daily_all.iterrows():
            id_akun_d = normalize_clean(r.get("Nama Pengguna", ""))
            nama_d = normalize_clean(r.get("Nama", ""))
            vals = {m: clean_numeric(r.get(m, 0)) for m in ALL_METRICS}
            vals["_raw_id"] = str(r.get("Nama Pengguna", "")).strip()
            vals["_raw_nama"] = str(r.get("Nama", "")).strip()
            if id_akun_d:
                daily_map[id_akun_d] = vals
            if nama_d:
                daily_map[nama_d] = vals

    # 2. Peta data MTD (Key: ID Akun Login & Nama Orang)
    mtd_map = {}
    if not df_mtd_all.empty:
        for _, r in df_mtd_all.iterrows():
            id_akun_m = normalize_clean(r.get("Nama Pengguna", ""))
            nama_m = normalize_clean(r.get("Nama", ""))
            vals = {
                "Total Penambahan Hari Ini": clean_numeric(r.get("Total Penambahan Hari Ini", 0)),
                "Total Penerimaan Hari Ini": clean_numeric(r.get("Total Penerimaan Hari Ini", 0)),
                "_raw_id": str(r.get("Nama Pengguna", "")).strip(),
                "_raw_nama": str(r.get("Nama", "")).strip(),
            }
            for m in METRIC_AVG_COLS:
                vals[m] = clean_numeric(r.get(m, 0))
            if id_akun_m:
                mtd_map[id_akun_m] = vals
            if nama_m:
                mtd_map[nama_m] = vals

    # 3. Kumpulkan SEMUA akun unik dari MTD maupun Daily (Union)
    all_account_keys = set()
    if not df_mtd_all.empty:
        for _, r in df_mtd_all.iterrows():
            k_id = normalize_clean(r.get("Nama Pengguna", ""))
            k_nm = normalize_clean(r.get("Nama", ""))
            if k_id:
                all_account_keys.add(k_id)
            elif k_nm:
                all_account_keys.add(k_nm)

    if not df_daily_all.empty:
        for _, r in df_daily_all.iterrows():
            k_id = normalize_clean(r.get("Nama Pengguna", ""))
            k_nm = normalize_clean(r.get("Nama", ""))
            if k_id:
                all_account_keys.add(k_id)
            elif k_nm:
                all_account_keys.add(k_nm)

    records = []
    for key_acc in all_account_keys:
        m_vals = mtd_map.get(key_acc, {})
        d_vals = daily_map.get(key_acc, {})

        # Cari info karyawan di database (cross-matching ID atau Nama)
        k_info = karyawan_dict.get(key_acc) or {}
        if not k_info:
            alt_id = d_vals.get("_raw_nama", "") or m_vals.get("_raw_nama", "")
            k_info = karyawan_dict.get(normalize_clean(alt_id), {})

        tl_val = k_info.get("Nama TL", "Unknown")
        bucket_val = k_info.get("Bucket", "Unknown")
        project_val = k_info.get("Project", "Unknown")
        status_val = k_info.get("Status Kerja", "Aktif")

        is_comp, perush_master = get_master_status(tl_val, bucket_val, project_val)

        pen_mtd = m_vals.get("Total Penambahan Hari Ini", 0.0)
        rec_mtd = m_vals.get("Total Penerimaan Hari Ini", 0.0)
        pen_daily = d_vals.get("Total Penambahan Hari Ini", 0.0)
        rec_daily = d_vals.get("Total Penerimaan Hari Ini", 0.0)

        id_agent_display = (
            d_vals.get("_raw_id", "")
            or m_vals.get("_raw_id", "")
            or k_info.get("ID_Akun_Resmi", "")
            or key_acc
        )
        nama_orang_display = (
            d_vals.get("_raw_nama", "")
            or m_vals.get("_raw_nama", "")
            or k_info.get("Nama_Orang_Resmi", "")
            or key_acc
        )

        row_item = {
            "ID Agent": id_agent_display,
            "Nama Pengguna": nama_orang_display,
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
# HELPER HITUNG DUA TABEL (BULANAN & HARIAN)
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


# ---------------------------------------------------------
# FUNGSI RENDER PERFORMANCE DENGAN DUA TABEL
# ---------------------------------------------------------
TARGET_BUCKETS = [
    ("S1 - SK", "S1", "SK"),
    ("D0 - SK", "D0", "SK"),
    ("S1 - LD", "S1", "LD"),
    ("S0 - LD", "S0", "LD"),
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
# TABS UTAMA DASHBOARD
# ---------------------------------------------------------
tab_karyawan, tab_perf_team, tab_perf_comp, tab_perf_agent, tab_final, tab_mpt_history, tab_tl_history = st.tabs([
    "Seluruh Karyawan",
    "Performance Team",
    "Performance Perusahaan",
    "Agent Performance",
    "Data Bulanan Final",
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

                        # 1. Tabel Bulanan: Tetap masukkan SEMUA agen (Aktif & Resign)
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

                        # 2. Tabel Harian: HANYA masukkan agen yang masih AKTIF (Resign dikecualikan)
                        if st_kerja == "Aktif":
                            pen_d = r["Total Penambahan Hari Ini_Daily"]
                            rec_d = r["Total Penerimaan Hari Ini_Daily"]
                            rate_d = (rec_d / pen_d * 100) if pen_d > 0 else 0.0

                            rows_d.append({
                                "ID Agent": r.get("ID Agent", ""),
                                "Nama Pengguna": r.get("Nama Pengguna", ""),
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

                    # Proses Tabel Bulanan
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

                    # Proses Tabel Harian
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
                        cols_ad = ["Rank", "ID Agent", "Nama Pengguna", "Penambahan (Hari Ini)", "Penerimaan (Hari Ini)", "Recovery Rate (Hari Ini)", "Gap", "Gap Amount", "Jumlah Panggilan", "Waktu Kerja Rata-rata", "Jumlah SMS", "Jumlah Pengiriman Template WABA"]
                        df_ad = df_ad[cols_ad]

                    # Tampilkan Berdampingan
                    col_left, col_right = st.columns([1, 1], gap="small")
                    with col_left:
                        st.markdown('<div class="sub-table-header header-bulanan">Agent Bulanan (月度坐席绩效)</div>', unsafe_allow_html=True)
                        display_full_table(translate_columns(df_am))
                    with col_right:
                        st.markdown('<div class="sub-table-header header-harian">Agent Harian - Aktif (今日在职坐席实时绩效)</div>', unsafe_allow_html=True)
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
with tab_final:
    st.caption("Perbandingan Recovery Rate dan Ranking antar Perusahaan bersebelahan (Bulan 7, 8, dan 9 MTD).")

    all_final_periods = sorted(final_dict.keys())
    periods_data = {}

    for p_code in all_final_periods:
        month_label = f"Bulan {p_code[:2]}/{p_code[2:]} ({p_code[2:]}年{int(p_code[:2])}月份)"
        df_p_proc = process_performance_data(final_dict[p_code], [])
        if not df_p_proc.empty:
            df_p_comp = df_p_proc[df_p_proc["Apakah Ikut Berkompetisi"] == "Ya"].copy()
            periods_data[month_label] = df_p_comp

    if not comp_agents.empty:
        periods_data["MTD Berjalan (当月至今)"] = comp_agents.copy()

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

        target_groups = [("SK", "D0"), ("SK", "S1"), ("LD", "S0"), ("LD", "S1")]
        period_keys = list(periods_data.keys())

        for proj, buck in target_groups:
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
# HELPER DATA HISTORIS (UNTUK AGENT & TL)
# =========================================================
all_period_records = {}
for p_code in all_final_periods:
    m_label = f"Bln {p_code[:2]}"
    df_p_proc = process_performance_data(final_dict[p_code], [])
    if not df_p_proc.empty:
        all_period_records[m_label] = df_p_proc[df_p_proc["Apakah Ikut Berkompetisi"] == "Ya"].copy()

if not comp_agents.empty:
    all_period_records["Bln 09 (MTD)"] = comp_agents.copy()

RANK_ORDER = {
    "Konsisten TOP 🌟": 1,
    "Terus Meningkat 📈": 2,
    "Fluktuatif 🔄": 3,
    "Terus Menurun 📉": 4,
    "Konsisten BOTTOM ⚠️": 5,
    "Data Baru 🆕": 6,
}


# =========================================================
# TAB 6: REKAM JEJAK HISTORIS AGEN MPT05 (STRICT TOP 30% RULE)
# =========================================================
with tab_mpt_history:
    st.caption("Memantau kinerja agen MPT05. Konsisten TOP HANYA jika di SEMUA bulan masuk dalam Top 30% bucket masing-masing.")

    if not all_period_records:
        st.info("Belum ada data historis yang dapat diproses.")
    else:
        mpt_agents_dict = {}

        # 1. Kumpulkan data dan hitung ranking persentil presisi per bucket per bulan
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
                                "is_top_30_list": [],      # True jika rank <= 30% bucket
                                "is_bottom_30_list": [],   # True jika rank > 70% bucket
                                "history": {},
                            }

                        rec_rate_val = ag["_rate"]
                        actual_rank = idx_rank + 1
                        
                        # Hitung persentil murni terhadap total agen di bucket tersebut
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

        # 2. Klasifikasikan Kinerja berdasarkan aturan Top 30%
        for id_ag, data in mpt_agents_dict.items():
            avg_rate = sum(data["rates"]) / len(data["rates"]) if data["rates"] else 0.0
            total_active_months = len(data["rates"])

            if total_active_months <= 1:
                trend_label = "Data Baru 🆕"
            else:
                # WAJIB SEMUA BULAN (all == True) berada di dalam Top 30%
                is_always_top_30 = all(data["is_top_30_list"])
                
                # WAJIB SEMUA BULAN berada di dalam Bottom 30%
                is_always_bottom_30 = all(data["is_bottom_30_list"])
                
                # Tren kenaikan/penurunan rate antar bulan
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
# TAB 7: REKAM JEJAK HISTORIS TEAM LEADER (TL) MPT05
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
                        # Logika Status TL: Jika memimpin agen aktif di periode ini, TL pasti Aktif
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