"""
wcl_web_app.py
================================================================================
Web Tools Interaktif untuk WCL NR26 Report Generator & Smart ISD Integrator.
Versi: 2.2 (Desain Mewah & Simpel, Filter Band, Smart Baseline Fallback, Multi-Device)
================================================================================
"""

import os
import io
import socket
import datetime
import importlib
import pandas as pd
import streamlit as st

import wcl_processor
importlib.reload(wcl_processor)
from wcl_processor import WCLProcessor, ISDMatcher, resolve_custom_dates, parse_date_val

# ── Konfigurasi Halaman Streamlit ─────────────────────────────────────────────
st.set_page_config(
    page_title="IOH WCL NR26 Studio",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded"
)

def get_network_ips():
    try:
        hostname = socket.gethostname()
        ips = socket.gethostbyname_ex(hostname)[2]
        clean_ips = [ip for ip in ips if not ip.startswith('127.')]
        return clean_ips or ['localhost']
    except:
        return ['localhost']

local_ips = get_network_ips()
primary_ip = local_ips[0] if local_ips else 'localhost'

# ── Custom CSS: Modern, Clean & Luxurious UI ─────────────────────────────────
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Plus Jakarta Sans', sans-serif;
    }}
    
    /* Header Mewah */
    .hero-container {{
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 50%, #1E3A8A 100%);
        border-radius: 16px;
        padding: 24px 28px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 16px;
    }}
    .hero-title {{
        font-size: 26px;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin: 0;
        background: linear-gradient(90deg, #FFFFFF 0%, #93C5FD 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }}
    .hero-subtitle {{
        font-size: 13px;
        color: #94A3B8;
        margin-top: 4px;
    }}
    .network-badge {{
        background: rgba(255, 255, 255, 0.08);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 30px;
        padding: 6px 14px;
        font-size: 12px;
        color: #E2E8F0;
        display: inline-flex;
        align-items: center;
        gap: 8px;
    }}
    .network-dot {{
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background-color: #10B981;
        box-shadow: 0 0 8px #10B981;
    }}

    /* Card Elegan */
    .luxury-card {{
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.03);
    }}
    .step-pill {{
        background: #EEF2FF;
        color: #4338CA;
        font-weight: 700;
        font-size: 11px;
        padding: 4px 10px;
        border-radius: 20px;
        display: inline-block;
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    .card-title {{
        font-size: 17px;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }}

    /* Stat Box Mewah */
    .stat-box {{
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 12px 16px;
        text-align: center;
    }}
    .stat-number {{
        font-size: 22px;
        font-weight: 800;
        color: #1E3A8A;
    }}
    .stat-label {{
        font-size: 11px;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
    }}

    /* Primary Action Button */
    .stButton>button {{
        background: linear-gradient(135deg, #2563EB 0%, #4338CA 100%) !important;
        color: white !important;
        border: none !important;
        padding: 12px 24px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        border-radius: 10px !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.25) !important;
        transition: all 0.2s ease-in-out !important;
    }}
    .stButton>button:hover {{
        transform: translateY(-1px);
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.35) !important;
    }}
</style>
""", unsafe_allow_html=True)


def init_session_state():
    if 'kpi_processor' not in st.session_state:
        st.session_state.kpi_processor = None
    if 'kpi_metadata' not in st.session_state:
        st.session_state.kpi_metadata = None
    if 'selected_cluster' not in st.session_state:
        st.session_state.selected_cluster = ""
    if 'isd_matcher' not in st.session_state:
        st.session_state.isd_matcher = ISDMatcher()
    if 'isd_df' not in st.session_state:
        st.session_state.isd_df = None
    if 'generated_report' not in st.session_state:
        st.session_state.generated_report = None
    if 'band_mode' not in st.session_state:
        st.session_state.band_mode = "nr26_baseline_nr21"

init_session_state()


# ══════════════════════════════════════════════════════════════════════════════
# HERO HEADER MEWAH
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="hero-container">
    <div>
        <div class="hero-title">IOH 5G NR26 Studio</div>
        <div class="hero-subtitle">Sistem Otomatisasi Analisis KPI NR26, Agregasi TWAMP, dan Integrasi Cerdas ISD (Site ID + Sector)</div>
    </div>
    <div class="network-badge">
        <div class="network-dot"></div>
        <span>Akses Laptop Lain: <b>http://{primary_ip}:8501</b></span>
    </div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR: SUMBER DATA
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📂 Data Sources")
    
    is_local_env = os.path.exists("D:\\")
    default_kpi_idx = 1 if is_local_env else 0
    default_twamp_idx = 2 if is_local_env else 0

    kpi_input_type = st.radio(
        "Metode Input KPI 5G:",
        ["Upload File (Single / Multi)", "Path Lokal / Folder (Batch)"],
        index=default_kpi_idx
    )

    kpi_sources_list = []
    if kpi_input_type == "Upload File (Single / Multi)":
        uploaded_kpis = st.file_uploader(
            "Upload File KPI (.xlsx):",
            type=["xlsx", "xlsm"],
            accept_multiple_files=True,
            help="Bisa pilih 1 file atau banyak file sekaligus"
        )
        if uploaded_kpis:
            kpi_sources_list = list(uploaded_kpis)
            st.caption(f"✅ {len(kpi_sources_list)} file terpilih.")
    else:
        default_kpi_path = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\MSH\KOTA DEPOK_01_N04_A\KPI 5G Daily.xlsx"
        kpi_path_input = st.text_input("Path File / Folder KPI:", value=default_kpi_path)
        if kpi_path_input and os.path.exists(kpi_path_input):
            if os.path.isdir(kpi_path_input):
                files = [os.path.join(kpi_path_input, f) for f in os.listdir(kpi_path_input) if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$')]
                kpi_sources_list = files
                st.caption(f"📁 Folder: {len(files)} file Excel ditemukan.")
            else:
                kpi_sources_list = [kpi_path_input]
                st.caption(f"📄 File tunggal ({round(os.path.getsize(kpi_path_input)/(1024*1024), 2)} MB)")
        elif kpi_path_input:
            st.error("❌ Path KPI tidak ditemukan!")

    st.markdown("---")
    st.markdown("### 📶 Sumber TWAMP")
    twamp_mode = st.radio(
        "Pilihan TWAMP:",
        ["Sheet 'twamp' Internal KPI", "Upload File Terpisah (Single/Batch)", "Path File / Folder Lokal"],
        index=default_twamp_idx
    )

    twamp_sources_list = None
    if twamp_mode == "Upload File Terpisah (Single/Batch)":
        uploaded_twamps = st.file_uploader("Upload File TWAMP (.xlsx):", type=["xlsx", "xlsm"], accept_multiple_files=True)
        if uploaded_twamps:
            twamp_sources_list = list(uploaded_twamps)
    elif twamp_mode == "Path File / Folder Lokal":
        default_twamp_path = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\MSH\KOTA DEPOK_01_N04_A\Twamp.xlsx"
        twamp_path_input = st.text_input("Path File TWAMP:", value=default_twamp_path)
        if twamp_path_input and os.path.exists(twamp_path_input):
            if os.path.isdir(twamp_path_input):
                files = [os.path.join(twamp_path_input, f) for f in os.listdir(twamp_path_input) if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$')]
                twamp_sources_list = files
            else:
                twamp_sources_list = [twamp_path_input]
        elif twamp_path_input:
            st.error("❌ Path TWAMP tidak ditemukan!")

    st.markdown("---")
    if kpi_sources_list:
        if st.button("🔄 Muat & Analisis Data", use_container_width=True):
            with st.spinner("Memindai skema KPI & TWAMP..."):
                try:
                    proc = WCLProcessor(kpi_sources=kpi_sources_list, twamp_sources=twamp_sources_list)
                    meta = proc.inspect_sources()
                    st.session_state.kpi_processor = proc
                    st.session_state.kpi_metadata = meta
                    if meta['clusters']:
                        st.session_state.selected_cluster = meta['clusters'][0]
                    st.success(f"Siap! Ditemukan {len(meta['clusters'])} cluster.")
                except Exception as e:
                    st.error(f"Gagal membaca data: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# WORKFLOW KARTU TIGA TAHAP: SIMPLE & MEWAH
# ══════════════════════════════════════════════════════════════════════════════
meta = st.session_state.kpi_metadata
avail_dates = meta.get('dates', []) if meta else []
avail_twamp_dates = meta.get('twamp_dates', []) if meta else []

# ── KARTU 1: KONFIGURASI TARGET & FILTER BAND ─────────────────────────────────
st.markdown("""
<div class="luxury-card">
    <div class="step-pill">Tahap 1</div>
    <div class="card-title">🎯 Target Cluster & Filter Band</div>
</div>
""", unsafe_allow_html=True)

c1, c2, c3 = st.columns([1.2, 1.3, 1])

with c1:
    if meta and meta.get('clusters'):
        selected_cluster = st.selectbox(
            "Pilih Cluster Target:",
            options=meta['clusters'],
            index=0
        )
        st.session_state.selected_cluster = selected_cluster
    else:
        st.selectbox("Cluster Target:", ["(Muat file KPI terlebih dahulu)"], disabled=True)

with c2:
    # ── FILTER BAND BARU ──────────────────────────────────────────────────────
    band_options = {
        "🌟 NR26 + Baseline NR21 (Otomatis isi Before jika NR26 belum on-air)": "nr26_baseline_nr21",
        "🔵 Hanya NR26 (5G2600)": "nr26_only",
        "🟢 Hanya NR21 (5G2100)": "nr21_only",
        "🟣 Semua Band (NR21 & NR26 Terpisah)": "all_bands_separate"
    }
    sel_band_label = st.selectbox(
        "Filter Band / Skema Komparasi:",
        options=list(band_options.keys()),
        index=0,
        help="Fitur 'NR26 + Baseline NR21' otomatis mengatasi sel Before yang kosong jika pada tanggal Before cluster belum on-air NR26!"
    )
    band_mode = band_options[sel_band_label]
    st.session_state.band_mode = band_mode

with c3:
    all_sites = meta.get('sites', []) if meta else []
    site_mode = st.radio("Filter Site:", ["Semua Site di Cluster", "Pilih Site Tertentu"], horizontal=True)
    site_filter = None
    if site_mode == "Pilih Site Tertentu":
        chosen_sites = st.multiselect("Pilih Site ID:", options=all_sites, default=[])
        site_filter = chosen_sites if chosen_sites else None


# ── KARTU 2: KONFIGURASI TANGGAL CUSTOM (BEFORE, AFTER, TWAMP) ────────────────
st.markdown("""
<div class="luxury-card">
    <div class="step-pill">Tahap 2</div>
    <div class="card-title">📅 Rentang Tanggal Analisis (Custom Range)</div>
</div>
""", unsafe_allow_html=True)

cb1, cb2, cb3 = st.columns(3)

if not avail_dates:
    st.info("💡 Silakan muat file KPI di sidebar sebelah kiri untuk memilih tanggal.")
    bef_dates_final = []
    aft_dates_final = []
    tw_dates_final = []
else:
    d_min = avail_dates[0]
    d_max = avail_dates[-1]

    # BEFORE
    with cb1:
        st.markdown("**🔵 BEFORE (WCL Pre-Optim)**")
        b_mode = st.selectbox("Preset BEFORE:", ["Rentang Tanggal Custom (Start & End)", "3 Hari Pertama (Otomatis)", "Checklist Tanggal"], key="b_mode_s")
        if b_mode == "Rentang Tanggal Custom (Start & End)":
            col_b1, col_b2 = st.columns(2)
            with col_b1: b_start = st.date_input("Mulai:", value=d_min, min_value=d_min, max_value=d_max, key="bs_s")
            with col_b2: b_end = st.date_input("Sampai:", value=avail_dates[min(6, len(avail_dates)-1)], min_value=d_min, max_value=d_max, key="be_s")
            if b_start > b_end: b_start, b_end = b_end, b_start
            bef_dates_final = [d for d in avail_dates if b_start <= d <= b_end]
        elif b_mode == "3 Hari Pertama (Otomatis)":
            bef_dates_final = avail_dates[:3]
        else:
            bef_dates_final = st.multiselect("Pilih:", options=avail_dates, default=avail_dates[:min(3, len(avail_dates))], key="bm_s")
        st.caption(f"Terpilih ({len(bef_dates_final)} hari): `{[d.strftime('%d-%b') for d in bef_dates_final]}`")

    # AFTER
    with cb2:
        st.markdown("**🟢 AFTER (Post-Optim)**")
        a_mode = st.selectbox("Preset AFTER:", ["Rentang Tanggal Custom (Start & End)", "3 Hari Terakhir (Otomatis)", "Checklist Tanggal"], key="a_mode_s")
        rem_dates = [d for d in avail_dates if d not in bef_dates_final]
        if a_mode == "Rentang Tanggal Custom (Start & End)":
            col_a1, col_a2 = st.columns(2)
            def_as = rem_dates[0] if rem_dates else d_max
            def_ae = rem_dates[min(6, len(rem_dates)-1)] if rem_dates else d_max
            with col_a1: a_start = st.date_input("Mulai:", value=def_as, min_value=d_min, max_value=d_max, key="as_s")
            with col_a2: a_end = st.date_input("Sampai:", value=def_ae, min_value=d_min, max_value=d_max, key="ae_s")
            if a_start > a_end: a_start, a_end = a_end, a_start
            aft_dates_final = [d for d in avail_dates if a_start <= d <= a_end]
        elif a_mode == "3 Hari Terakhir (Otomatis)":
            aft_dates_final = rem_dates[:3] if rem_dates else avail_dates[-3:]
        else:
            aft_dates_final = st.multiselect("Pilih:", options=avail_dates, default=rem_dates[:min(3, len(rem_dates))] if rem_dates else [], key="am_s")
        st.caption(f"Terpilih ({len(aft_dates_final)} hari): `{[d.strftime('%d-%b') for d in aft_dates_final]}`")

    # TWAMP
    with cb3:
        st.markdown("**🟡 TWAMP**")
        tw_pool = avail_twamp_dates if avail_twamp_dates else avail_dates
        tw_mode = st.selectbox("Preset TWAMP:", ["7 Hari Terakhir (Otomatis)", "Rentang Tanggal Custom"], key="tw_mode_s")
        if tw_mode == "7 Hari Terakhir (Otomatis)":
            tw_dates_final = tw_pool[-7:]
        else:
            tw_dates_final = tw_pool[-min(7, len(tw_pool)):]
        st.caption(f"Terpilih ({len(tw_dates_final)} hari): `{[d.strftime('%d-%b') for d in tw_dates_final]}`")

    st.session_state.bef_dates_final = bef_dates_final
    st.session_state.aft_dates_final = aft_dates_final
    st.session_state.tw_dates_final  = tw_dates_final


# ── KARTU 3: DATA ISD & GENERATE REPORT ───────────────────────────────────────
st.markdown("""
<div class="luxury-card">
    <div class="step-pill">Tahap 3</div>
    <div class="card-title">📐 Integrasi Data ISD (Site + Sector) & Generate Laporan</div>
</div>
""", unsafe_allow_html=True)

with st.expander("📝 Pengaturan & Input Data ISD (Buka untuk Ubah / Paste)", expanded=False):
    isd_method = st.radio("Metode Input ISD:", ["Paste Tabel Langsung", "Upload File Excel/CSV", "Gunakan Contoh Bawaan"], horizontal=True)
    isd_df_loaded = None

    if isd_method == "Paste Tabel Langsung":
        sample_paste = (
            "Site ID\tSector\tISD\n"
            "13DPK0120\tsector 1\t450\n"
            "13DPK0120\tsector 2\t520\n"
            "13DPK0120\tsector 3\t480\n"
            "13DPK0134\tsector 1\t600\n"
            "13DPK0134\tsector 2\t580\n"
            "13DPK0134\tsector 3\t620\n"
            "13DPK0146\tsector 1\t410\n"
            "13DPK0146\tsector 2\t430\n"
            "13DPK0146\tsector 3\t490\n"
            "13DPK0147\tsector 1\t550\n"
            "13DPK0159\tsector 1\t510\n"
            "13DPK0167\tsector 1\t470"
        )
        paste_text = st.text_area("Paste data dari Excel di sini:", value=sample_paste, height=140)
        if paste_text.strip():
            try:
                sep = '\t' if '\t' in paste_text else ','
                isd_df_loaded = pd.read_csv(io.StringIO(paste_text), sep=sep)
            except Exception as e:
                st.error(f"Format tabel tidak valid: {e}")

    elif isd_method == "Upload File Excel/CSV":
        isd_file = st.file_uploader("Upload File ISD:", type=["xlsx", "xls", "csv"])
        if isd_file:
            try:
                if isd_file.name.lower().endswith('.csv'):
                    isd_df_loaded = pd.read_csv(isd_file)
                else:
                    xl = pd.ExcelFile(isd_file)
                    sh = st.selectbox("Pilih Sheet:", xl.sheet_names)
                    isd_df_loaded = xl.parse(sh)
            except Exception as e:
                st.error(f"Gagal membaca file ISD: {e}")

    else:
        sample_data = {
            "Site ID": ["13DPK0120", "13DPK0120", "13DPK0120", "13DPK0134", "13DPK0134", "13DPK0134", "13DPK0146", "13DPK0146", "13DPK0146", "13DPK0147", "13DPK0159", "13DPK0167"],
            "Sector": ["sector 1", "sector 2", "sector 3", "sector 1", "sector 2", "sector 3", "sector 1", "sector 2", "sector 3", "sector 1", "sector 1", "sector 1"],
            "ISD (m)": [450, 520, 480, 600, 580, 620, 410, 430, 490, 550, 510, 470]
        }
        isd_df_loaded = pd.DataFrame(sample_data)

    if isd_df_loaded is not None:
        st.session_state.isd_df = isd_df_loaded
        matcher = ISDMatcher(isd_df_loaded)
        st.session_state.isd_matcher = matcher
        st.dataframe(isd_df_loaded.head(6), use_container_width=True, height=140)

# Live Preview & Status Matched ISD
proc = st.session_state.kpi_processor
target_cluster = st.session_state.selected_cluster

if proc and target_cluster:
    try:
        try:
            cells_preview = proc.get_cluster_cells_preview(target_cluster, site_filter=site_filter, band_filter=band_mode)
        except TypeError:
            # Re-instantiate proc with fresh class in case of stale cache
            import importlib
            import wcl_processor
            importlib.reload(wcl_processor)
            proc = wcl_processor.WCLProcessor(kpi_sources=proc.kpi_sources, twamp_sources=proc.twamp_sources)
            st.session_state.kpi_processor = proc
            try:
                cells_preview = proc.get_cluster_cells_preview(target_cluster, site_filter=site_filter, band_filter=band_mode)
            except TypeError:
                cells_preview = proc.get_cluster_cells_preview(target_cluster, site_filter=site_filter)

        matcher = st.session_state.isd_matcher
        preview_df = matcher.match_preview(cells_preview)

        total_cells = len(preview_df)
        matched_cells = (preview_df['Status'] == 'Matched').sum()
        missing_cells = total_cells - matched_cells
        match_rate = (matched_cells / total_cells * 100) if total_cells > 0 else 0

        # Ringkasan Stat Baris
        m1, m2, m3, m4 = st.columns(4)
        with m1: st.markdown(f'<div class="stat-box"><div class="stat-number">{total_cells}</div><div class="stat-label">Total Cell</div></div>', unsafe_allow_html=True)
        with m2: st.markdown(f'<div class="stat-box"><div class="stat-number" style="color:#059669;">{matched_cells}</div><div class="stat-label">ISD Cocok</div></div>', unsafe_allow_html=True)
        with m3: st.markdown(f'<div class="stat-box"><div class="stat-number" style="color:#DC2626;">{missing_cells}</div><div class="stat-label">ISD Kosong</div></div>', unsafe_allow_html=True)
        with m4: st.markdown(f'<div class="stat-box"><div class="stat-number" style="color:#4338CA;">{match_rate:.0f}%</div><div class="stat-label">Match Rate</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Gagal memuat status cell: {e}")

# Tombol Eksekusi Besar Mewah
col_run, col_out = st.columns([1.5, 1.5])
with col_run:
    btn_generate = st.button("🚀 GENERATE WCL REPORT SEKARANG", use_container_width=True)

with col_out:
    out_dir_def = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\MSH\KOTA DEPOK_01_N04_A\Output" if is_local_env else ""
    save_folder = st.text_input("Simpan juga ke folder lokal:", value=out_dir_def)

if btn_generate:
    if not proc or not target_cluster:
        st.error("Pilih data sumber dan cluster target terlebih dahulu!")
    else:
        bef_d = st.session_state.get('bef_dates_final', None)
        aft_d = st.session_state.get('aft_dates_final', None)
        tw_d  = st.session_state.get('tw_dates_final', None)

        progress_bar = st.progress(0)
        status_text = st.empty()
        log_exp = st.expander("📜 Log Pemrosesan Real-Time", expanded=True)
        log_box = log_exp.empty()
        log_records = []

        def progress_cb(pct, msg):
            progress_bar.progress(pct)
            status_text.markdown(f"**Status:** {msg}")
            log_records.append(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")
            log_box.code("\n".join(log_records))

        out_fname = f"WCL NR26 {target_cluster}_NEW.xlsx"
        out_fpath = os.path.join(save_folder, out_fname) if (save_folder and os.path.exists(save_folder)) else None

        try:
            proc.isd_matcher = st.session_state.isd_matcher
            try:
                res = proc.generate_report(
                    target_cluster=target_cluster,
                    site_filter=site_filter,
                    band_filter=band_mode,
                    before_dates=bef_d,
                    after_dates=aft_d,
                    twamp_dates=tw_d,
                    output_file=out_fpath,
                    progress_callback=progress_cb
                )
            except TypeError:
                import importlib
                import wcl_processor
                importlib.reload(wcl_processor)
                proc = wcl_processor.WCLProcessor(kpi_sources=proc.kpi_sources, twamp_sources=proc.twamp_sources)
                proc.isd_matcher = st.session_state.isd_matcher
                st.session_state.kpi_processor = proc
                res = proc.generate_report(
                    target_cluster=target_cluster,
                    site_filter=site_filter,
                    band_filter=band_mode,
                    before_dates=bef_d,
                    after_dates=aft_d,
                    twamp_dates=tw_d,
                    output_file=out_fpath,
                    progress_callback=progress_cb
                )

            st.session_state.generated_report = res
            st.balloons()
            st.success(f"🎉 Sukses! Laporan berhasil dibuat: {len(res['sheets'])} sheet, {res['total_cells']} cell ({res['isd_matched']} ISD terisi).")
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses laporan: {e}")
            import traceback
            st.code(traceback.format_exc())

# Tombol Download
gen_res = st.session_state.generated_report
if gen_res and gen_res.get('bytes'):
    st.markdown("---")
    dl_cl = gen_res.get('cluster', 'OUTPUT')
    dl_name = f"WCL NR26 {dl_cl}_NEW.xlsx"
    
    cdl1, cdl2 = st.columns([1, 2])
    with cdl1:
        st.download_button(
            label=f"📥 Download {dl_name}",
            data=gen_res['bytes'],
            file_name=dl_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with cdl2:
        st.markdown(f"**Periode BEFORE ({len(gen_res.get('wcl_dates', []))} hari):** `{', '.join(gen_res.get('wcl_dates', []))}`")
        st.markdown(f"**Periode AFTER ({len(gen_res.get('post_dates', []))} hari):** `{', '.join(gen_res.get('post_dates', []))}`")
