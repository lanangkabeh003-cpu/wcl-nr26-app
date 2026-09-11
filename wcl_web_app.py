"""
wcl_web_app.py
================================================================================
Web Tools Interaktif untuk WCL NR26 Report Generator & Smart ISD Integrator.
Versi: 2.0 (Single/Batch KPI & TWAMP, Multi-Schema, Custom Date Ranges, Multi-Device)
================================================================================
"""

import os
import io
import socket
import datetime
import pandas as pd
import streamlit as st

from wcl_processor import WCLProcessor, ISDMatcher, resolve_custom_dates, parse_date_val

# ── Konfigurasi Halaman Streamlit ─────────────────────────────────────────────
st.set_page_config(
    page_title="WCL NR26 Generator - Smart ISD",
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

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        font-size: 24px;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 2px;
    }
    .sub-header {
        font-size: 13px;
        color: #4B5563;
        margin-bottom: 12px;
    }
    .net-banner {
        background-color: #EFF6FF;
        border: 1px solid #BFDBFE;
        border-radius: 8px;
        padding: 8px 14px;
        margin-bottom: 16px;
        font-size: 13px;
        color: #1E40AF;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .stButton>button {
        font-weight: 600;
        border-radius: 6px;
    }
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

init_session_state()


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR: SUMBER DATA KPI (SINGLE / BATCH) & TWAMP (SINGLE / BATCH)
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.image("https://img.icons8.com/color/96/5g.png", width=56)
    st.title("WCL Configuration")

    # Informasi Akses Jaringan Multi-Device
    st.info(f"🌐 **Buka di Laptop Lain:**\n`http://{primary_ip}:8501`\n*(Pastikan satu Wi-Fi/LAN)*")
    st.markdown("---")

    # ── 1. SUMBER DATA KPI (SINGLE ATAU BATCH) ────────────────────────────────
    st.subheader("1. Sumber Data KPI 5G")
    is_local_env = os.path.exists("D:\\")
    default_kpi_idx = 1 if is_local_env else 0
    default_twamp_idx = 2 if is_local_env else 0

    kpi_input_type = st.radio(
        "Tipe Input KPI:",
        ["Upload File (Single / Multi)", "Path Lokal / Folder (Batch)"],
        index=default_kpi_idx
    )

    kpi_sources_list = []

    if kpi_input_type == "Upload File (Single / Multi)":
        uploaded_kpis = st.file_uploader(
            "Pilih File KPI (.xlsx):",
            type=["xlsx", "xlsm"],
            accept_multiple_files=True,
            help="Bisa upload 1 file (single) atau banyak file sekaligus (batch)"
        )
        if uploaded_kpis:
            kpi_sources_list = list(uploaded_kpis)
            st.caption(f"✅ {len(kpi_sources_list)} file KPI terpilih.")
    else:
        default_kpi_path = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\MSH\KOTA DEPOK_01_N04_A\KPI 5G Daily.xlsx"
        kpi_path_input = st.text_input(
            "Path File atau Folder KPI:",
            value=default_kpi_path,
            help="Masukkan path file .xlsx atau path folder untuk batch scanning seluruh file di folder tersebut"
        )
        if kpi_path_input and os.path.exists(kpi_path_input):
            if os.path.isdir(kpi_path_input):
                files = [os.path.join(kpi_path_input, f) for f in os.listdir(kpi_path_input) if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$')]
                kpi_sources_list = files
                st.caption(f"📁 Folder: Ditemukan {len(files)} file Excel KPI.")
            else:
                kpi_sources_list = [kpi_path_input]
                st.caption(f"📄 File tunggal ({round(os.path.getsize(kpi_path_input)/(1024*1024), 2)} MB)")
        elif kpi_path_input:
            st.error("❌ Path KPI tidak ditemukan!")

    # ── 2. SUMBER DATA TWAMP (SEPARATE FILE / BATCH / INTERNAL) ───────────────
    st.subheader("2. Sumber Data TWAMP")
    twamp_mode = st.radio(
        "Pilihan Sumber TWAMP:",
        ["Gunakan Sheet 'twamp' Internal KPI", "Upload File TWAMP Terpisah (Single/Batch)", "Path File / Folder TWAMP Lokal"],
        index=default_twamp_idx
    )

    twamp_sources_list = None

    if twamp_mode == "Upload File TWAMP Terpisah (Single/Batch)":
        uploaded_twamps = st.file_uploader(
            "Pilih File TWAMP (.xlsx):",
            type=["xlsx", "xlsm"],
            accept_multiple_files=True,
            key="uploader_twamp"
        )
        if uploaded_twamps:
            twamp_sources_list = list(uploaded_twamps)
            st.caption(f"✅ {len(twamp_sources_list)} file TWAMP terpilih.")

    elif twamp_mode == "Path File / Folder TWAMP Lokal":
        default_twamp_path = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\MSH\KOTA DEPOK_01_N04_A\Twamp.xlsx"
        twamp_path_input = st.text_input("Path File atau Folder TWAMP:", value=default_twamp_path)
        if twamp_path_input and os.path.exists(twamp_path_input):
            if os.path.isdir(twamp_path_input):
                files = [os.path.join(twamp_path_input, f) for f in os.listdir(twamp_path_input) if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$')]
                twamp_sources_list = files
                st.caption(f"📁 Folder: Ditemukan {len(files)} file TWAMP.")
            else:
                twamp_sources_list = [twamp_path_input]
                st.caption(f"📄 File TWAMP ({round(os.path.getsize(twamp_path_input)/(1024*1024), 2)} MB)")
        elif twamp_path_input:
            st.error("❌ Path TWAMP tidak ditemukan!")

    # Tombol Muat & Analisis
    if kpi_sources_list:
        if st.button("🔄 Muat & Analisis Data Sumber", use_container_width=True, type="primary"):
            with st.spinner("Memindai dan mengidentifikasi skema file KPI & TWAMP..."):
                try:
                    proc = WCLProcessor(kpi_sources=kpi_sources_list, twamp_sources=twamp_sources_list)
                    meta = proc.inspect_sources()
                    st.session_state.kpi_processor = proc
                    st.session_state.kpi_metadata = meta
                    if meta['clusters']:
                        st.session_state.selected_cluster = meta['clusters'][0]
                    st.success(f"Berhasil! Ditemukan {len(meta['clusters'])} cluster & {len(meta['dates'])} tanggal.")
                except Exception as e:
                    st.error(f"Gagal memuat sumber data: {e}")

    st.markdown("---")

    # ── 3. CLUSTER & SITE FILTER ──────────────────────────────────────────────
    meta = st.session_state.kpi_metadata
    if meta and meta.get('clusters'):
        st.subheader("3. Target Cluster & Site")
        selected_cluster = st.selectbox(
            "Pilih Cluster Target:",
            options=meta['clusters'],
            index=0
        )
        st.session_state.selected_cluster = selected_cluster

        all_sites = meta.get('sites', [])
        filter_mode = st.radio("Filter Site ID:", ["Semua Site di Cluster", "Pilih Site Spesifik"], index=0)
        
        site_filter = None
        if filter_mode == "Pilih Site Spesifik":
            chosen_sites = st.multiselect("Pilih Site ID:", options=all_sites, default=[])
            manual_sites = st.text_area("Atau ketik/paste Site ID (pisahkan koma/spasi):", placeholder="13DPK0120, 13DPK0134")
            if manual_sites:
                extra_sites = [s.strip().upper() for s in manual_sites.replace('\n', ',').replace(' ', ',').split(',') if s.strip()]
                chosen_sites = list(set(chosen_sites + extra_sites))
            site_filter = chosen_sites if chosen_sites else None
    else:
        site_filter = None
        st.info("Muat data KPI terlebih dahulu untuk memilih cluster.")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN AREA
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="main-header">📡 IOH 5G NR26 WCL Report Generator & Smart ISD Integrator</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="net-banner">'
    f'<span>💻 <b>Web Access:</b> Buka dari browser laptop Anda (<code>http://localhost:8501</code>) atau dari laptop rekan di jaringan yang sama: <code>http://{primary_ip}:8501</code></span>'
    f'</div>',
    unsafe_allow_html=True
)

tab_isd, tab_dates, tab_generate = st.tabs([
    "📐 1. Input & Pencocokan ISD (Site + Sector)",
    "📅 2. Konfigurasi Tanggal Custom (BEFORE, AFTER, TWAMP)",
    "🚀 3. Eksekusi & Download Report"
])


# ──────────────────────────────────────────────────────────────────────────────
# TAB 1: INPUT & PENCOCOKAN ISD
# ──────────────────────────────────────────────────────────────────────────────
with tab_isd:
    st.subheader("Pengaturan Data ISD (Inter-Site Distance)")
    st.info(
        "💡 **Pencocokan Otomatis Kombinasi Site ID & Sector:**\n"
        "- Mendukung format **2 kolom terpisah** (Kolom `Site_ID` dan Kolom `Sector`).\n"
        "- Mendukung format **1 kolom gabungan** (contoh: `13DPK0208 sector 1`, `13DPK0208_1`, `13DPK0208 1`).\n"
        "Nilai ISD akan otomatis diisikan ke kolom `ISD m` di sheet `TA Avg`."
    )

    isd_input_method = st.radio(
        "Pilih Metode Memasukkan Data ISD:",
        ["Copy & Paste Tabel Data Langsung", "Upload File ISD (.xlsx / .csv)", "Gunakan Data Contoh / Bawaan"],
        horizontal=True
    )

    isd_df_loaded = None

    if isd_input_method == "Copy & Paste Tabel Data Langsung":
        st.markdown("**Paste teks dari Excel (tab-separated atau comma-separated):**")
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
        paste_text = st.text_area("Paste Data Tabel ISD di sini:", value=sample_paste, height=180)
        if paste_text.strip():
            try:
                sep = '\t' if '\t' in paste_text else ','
                isd_df_loaded = pd.read_csv(io.StringIO(paste_text), sep=sep)
                st.caption(f"✅ Terbaca {len(isd_df_loaded)} baris data dari paste.")
            except Exception as e:
                st.error(f"Gagal membaca format tabel: {e}")

    elif isd_input_method == "Upload File ISD (.xlsx / .csv)":
        isd_file = st.file_uploader("Upload File ISD:", type=["xlsx", "xls", "csv"], key="isd_uploader_v2")
        if isd_file is not None:
            try:
                if isd_file.name.lower().endswith('.csv'):
                    isd_df_loaded = pd.read_csv(isd_file)
                else:
                    xl = pd.ExcelFile(isd_file)
                    selected_isd_sheet = st.selectbox("Pilih Sheet ISD:", xl.sheet_names)
                    isd_df_loaded = xl.parse(selected_isd_sheet)
                st.success(f"Berhasil membaca file ISD: {len(isd_df_loaded)} baris data.")
            except Exception as e:
                st.error(f"Gagal membaca file ISD: {e}")

    elif isd_input_method == "Gunakan Data Contoh / Bawaan":
        sample_data = {
            "Site ID": ["13DPK0120", "13DPK0120", "13DPK0120", "13DPK0134", "13DPK0134", "13DPK0134", "13DPK0146", "13DPK0146", "13DPK0146", "13DPK0147", "13DPK0159", "13DPK0167"],
            "Sector": ["sector 1", "sector 2", "sector 3", "sector 1", "sector 2", "sector 3", "sector 1", "sector 2", "sector 3", "sector 1", "sector 1", "sector 1"],
            "ISD (m)": [450, 520, 480, 600, 580, 620, 410, 430, 490, 550, 510, 470]
        }
        isd_df_loaded = pd.DataFrame(sample_data)

    if isd_df_loaded is not None:
        st.session_state.isd_df = isd_df_loaded

        with st.expander("🛠️ Penyesuaian Pemetaan Kolom ISD (Auto-Detected)", expanded=False):
            cols = list(isd_df_loaded.columns)
            c1, c2, c3, c4 = st.columns(4)
            with c1: sel_site_col = st.selectbox("Kolom Site ID:", ["(Auto)"] + cols, index=0)
            with c2: sel_sec_col = st.selectbox("Kolom Sector:", ["(Auto)", "(Tidak Ada / Tergabung)"] + cols, index=0)
            with c3: sel_isd_col = st.selectbox("Kolom Nilai ISD (meter):", ["(Auto)"] + cols, index=0)
            with c4: sel_comb_col = st.selectbox("Kolom Gabungan (Site+Sector):", ["(Auto)", "(Tidak Ada)"] + cols, index=0)

        site_arg = None if sel_site_col == "(Auto)" else sel_site_col
        sec_arg  = None if sel_sec_col in ["(Auto)", "(Tidak Ada / Tergabung)"] else sel_sec_col
        isd_arg  = None if sel_isd_col == "(Auto)" else sel_isd_col
        comb_arg = None if sel_comb_col in ["(Auto)", "(Tidak Ada)"] else sel_comb_col

        matcher = ISDMatcher(isd_df_loaded, site_col=site_arg, sector_col=sec_arg, isd_col=isd_arg, combined_col=comb_arg)
        st.session_state.isd_matcher = matcher

        st.markdown("##### Tabel Data ISD yang Dimuat:")
        st.dataframe(isd_df_loaded.head(10), use_container_width=True, height=160)

    # Live Preview Matching
    st.markdown("---")
    st.subheader("🔍 Live Preview Pencocokan ISD untuk Cluster Target")

    proc = st.session_state.kpi_processor
    target_cluster = st.session_state.selected_cluster

    if proc and target_cluster:
        try:
            cells_preview = proc.get_cluster_cells_preview(target_cluster, site_filter=site_filter)
            matcher = st.session_state.isd_matcher
            preview_df = matcher.match_preview(cells_preview)

            total_cells = len(preview_df)
            matched_cells = (preview_df['Status'] == 'Matched').sum()
            missing_cells = total_cells - matched_cells
            match_rate = (matched_cells / total_cells * 100) if total_cells > 0 else 0

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Cell 5G26 Target", f"{total_cells} cell")
            m2.metric("ISD Cocok (Matched)", f"{matched_cells} cell")
            m3.metric("ISD Belum Ada", f"{missing_cells} cell")
            m4.metric("Persentase Kecocokan", f"{match_rate:.1f}%")

            def highlight_status(val):
                color = '#D1FAE5' if val == 'Matched' else '#FEE2E2'
                text_color = '#065F46' if val == 'Matched' else '#991B1B'
                return f'background-color: {color}; color: {text_color}; font-weight: bold;'

            st.dataframe(
                preview_df.style.map(highlight_status, subset=['Status']),
                use_container_width=True,
                height=260
            )
            if missing_cells == 0 and total_cells > 0:
                st.success("🎉 Seluruh cell pada cluster ini berhasil dicocokkan dengan nilai ISD!")
        except Exception as e:
            st.error(f"Gagal menampilkan preview cell: {e}")
    else:
        st.info("Silakan muat file KPI dan pilih cluster pada sidebar sebelah kiri.")


# ──────────────────────────────────────────────────────────────────────────────
# TAB 2: KONFIGURASI TANGGAL CUSTOM (BEFORE, AFTER, TWAMP)
# ──────────────────────────────────────────────────────────────────────────────
with tab_dates:
    st.subheader("Konfigurasi Periode Tanggal Custom")
    st.markdown("Atur rentang tanggal secara bebas untuk **BEFORE (WCL Pre-Optim)**, **AFTER (Post-Optim)**, dan **TWAMP**.")

    meta = st.session_state.kpi_metadata
    avail_dates = meta.get('dates', []) if meta else []
    avail_twamp_dates = meta.get('twamp_dates', []) if meta else []

    if not avail_dates:
        st.warning("Silakan muat file KPI pada sidebar untuk membaca rentang tanggal yang tersedia.")
    else:
        c_bef, c_aft, c_tw = st.columns(3)

        d_min = avail_dates[0]
        d_max = avail_dates[-1]

        # ── BEFORE DATES ──────────────────────────────────────────────────────
        with c_bef:
            st.markdown("#### 🔵 BEFORE (Pre-Optim)")
            bef_mode = st.radio(
                "Mode BEFORE:",
                ["Rentang Tanggal (Start & End)", "3 Hari Pertama (Otomatis)", "Pilih Tanggal Spesifik (Multiselect)"],
                key="bef_mode_v2"
            )
            bef_dates_final = []

            if bef_mode == "Rentang Tanggal (Start & End)":
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    b_start = st.date_input("Start Date BEFORE:", value=d_min, min_value=d_min, max_value=d_max, key="b_start_inp")
                with col_b2:
                    default_b_end = avail_dates[min(2, len(avail_dates)-1)]
                    b_end = st.date_input("End Date BEFORE:", value=default_b_end, min_value=d_min, max_value=d_max, key="b_end_inp")

                if b_start > b_end:
                    b_start, b_end = b_end, b_start
                bef_dates_final = [d for d in avail_dates if b_start <= d <= b_end]

            elif bef_mode == "3 Hari Pertama (Otomatis)":
                bef_dates_final = avail_dates[:3]
            else:
                bef_dates_final = st.multiselect("Checklist Tanggal BEFORE:", options=avail_dates, default=avail_dates[:min(3, len(avail_dates))], key="bef_multi_v2")

            st.success(f"**BEFORE ({len(bef_dates_final)} hari):** {[d.strftime('%d-%b') for d in bef_dates_final]}")

        # ── AFTER DATES ───────────────────────────────────────────────────────
        with c_aft:
            st.markdown("#### 🟢 AFTER (Post-Optim)")
            aft_mode = st.radio(
                "Mode AFTER:",
                ["Rentang Tanggal (Start & End)", "3 Hari Terakhir (Otomatis)", "Pilih Tanggal Spesifik (Multiselect)"],
                key="aft_mode_v2"
            )
            aft_dates_final = []

            rem_dates = [d for d in avail_dates if d not in bef_dates_final]
            if aft_mode == "Rentang Tanggal (Start & End)":
                col_a1, col_a2 = st.columns(2)
                def_a_start = rem_dates[0] if rem_dates else d_max
                def_a_end = rem_dates[min(2, len(rem_dates)-1)] if rem_dates else d_max
                with col_a1:
                    a_start = st.date_input("Start Date AFTER:", value=def_a_start, min_value=d_min, max_value=d_max, key="a_start_inp")
                with col_a2:
                    a_end = st.date_input("End Date AFTER:", value=def_a_end, min_value=d_min, max_value=d_max, key="a_end_inp")

                if a_start > a_end:
                    a_start, a_end = a_end, a_start
                aft_dates_final = [d for d in avail_dates if a_start <= d <= a_end]

            elif aft_mode == "3 Hari Terakhir (Otomatis)":
                aft_dates_final = rem_dates[:3] if rem_dates else avail_dates[-3:]
            else:
                def_sel_aft = rem_dates[:min(3, len(rem_dates))] if rem_dates else []
                aft_dates_final = st.multiselect("Checklist Tanggal AFTER:", options=avail_dates, default=def_sel_aft, key="aft_multi_v2")

            st.success(f"**AFTER ({len(aft_dates_final)} hari):** {[d.strftime('%d-%b') for d in aft_dates_final]}")

        # ── TWAMP DATES ───────────────────────────────────────────────────────
        with c_tw:
            st.markdown("#### 🟡 TWAMP")
            tw_pool = avail_twamp_dates if avail_twamp_dates else avail_dates
            tw_mode = st.radio(
                "Mode TWAMP:",
                ["7 Hari Terakhir (Otomatis)", "Rentang Tanggal (Start & End)", "Pilih Tanggal Spesifik"],
                key="tw_mode_v2"
            )
            tw_dates_final = []

            if tw_mode == "7 Hari Terakhir (Otomatis)":
                tw_dates_final = tw_pool[-7:]
            elif tw_mode == "Rentang Tanggal (Start & End)":
                tw_min = tw_pool[0]
                tw_max = tw_pool[-1]
                col_t1, col_t2 = st.columns(2)
                with col_t1: t_start = st.date_input("Start Date TWAMP:", value=tw_min, min_value=tw_min, max_value=tw_max, key="t_start_inp")
                with col_t2: t_end   = st.date_input("End Date TWAMP:", value=tw_max, min_value=tw_min, max_value=tw_max, key="t_end_inp")
                if t_start > t_end: t_start, t_end = t_end, t_start
                tw_dates_final = [d for d in tw_pool if t_start <= d <= t_end]
            else:
                tw_dates_final = st.multiselect("Checklist Tanggal TWAMP:", options=tw_pool, default=tw_pool[-min(7, len(tw_pool)):], key="tw_multi_v2")

            st.success(f"**TWAMP ({len(tw_dates_final)} hari):** {[d.strftime('%d-%b') for d in tw_dates_final]}")

        st.session_state.bef_dates_final = bef_dates_final
        st.session_state.aft_dates_final = aft_dates_final
        st.session_state.tw_dates_final  = tw_dates_final


# ──────────────────────────────────────────────────────────────────────────────
# TAB 3: EKSEKUSI & DOWNLOAD REPORT
# ──────────────────────────────────────────────────────────────────────────────
with tab_generate:
    st.subheader("Generate & Unduh Laporan WCL NR26")
    st.markdown("Klik tombol di bawah ini untuk memulai proses generate seluruh sheet KPI, TA Avg (dengan data ISD), dan TWAMP.")

    col_btn, col_path = st.columns([2, 2])
    with col_btn:
        output_folder_default = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\MSH\KOTA DEPOK_01_N04_A\Output"
        custom_out_dir = st.text_input("Folder Output Lokal Komputer (Opsional):", value=output_folder_default)
        btn_run = st.button("🚀 GENERATE WCL NR26 REPORT SEKARANG", type="primary", use_container_width=True)

    if btn_run:
        proc = st.session_state.kpi_processor
        target_cluster = st.session_state.selected_cluster

        if not proc or not target_cluster:
            st.error("File KPI atau Cluster target belum dipilih! Silakan cek sidebar.")
        else:
            bef_d = st.session_state.get('bef_dates_final', None)
            aft_d = st.session_state.get('aft_dates_final', None)
            tw_d  = st.session_state.get('tw_dates_final', None)

            progress_bar = st.progress(0)
            status_text = st.empty()
            log_expander = st.expander("📜 Log Real-Time Eksekusi", expanded=True)
            log_placeholder = log_expander.empty()
            log_messages = []

            def ui_progress_callback(pct, msg):
                progress_bar.progress(pct)
                status_text.markdown(f"**Status:** {msg}")
                log_messages.append(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")
                log_placeholder.code("\n".join(log_messages))

            out_filename = f"WCL NR26 {target_cluster}_NEW.xlsx"
            out_filepath = os.path.join(custom_out_dir, out_filename) if custom_out_dir else None

            try:
                proc.isd_matcher = st.session_state.isd_matcher
                result = proc.generate_report(
                    target_cluster=target_cluster,
                    site_filter=site_filter,
                    before_dates=bef_d,
                    after_dates=aft_d,
                    twamp_dates=tw_d,
                    output_file=out_filepath,
                    progress_callback=ui_progress_callback
                )

                st.session_state.generated_report = result
                st.balloons()
                st.success(f"🎉 Sukses! Laporan berhasil disusun: {len(result['sheets'])} sheet, {result['total_cells']} cell ({result['isd_matched']} ISD terisi).")
            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses laporan: {e}")
                import traceback
                st.code(traceback.format_exc())

    gen_res = st.session_state.generated_report
    if gen_res and gen_res.get('bytes'):
        st.markdown("---")
        st.subheader("📥 Download File Laporan Excel")

        target_cl = gen_res.get('cluster', 'OUTPUT')
        download_name = f"WCL NR26 {target_cl}_NEW.xlsx"

        c_dl, c_info = st.columns([1, 2])
        with c_dl:
            st.download_button(
                label=f"⬇️ Download {download_name}",
                data=gen_res['bytes'],
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        with c_info:
            st.markdown(f"""
            - **Nama File**: `{download_name}`
            - **Sheets**: `{", ".join(gen_res.get('sheets', []))}`
            - **Total Cell 5G26**: `{gen_res.get('total_cells', 0)}`
            - **ISD Berhasil Dicocokkan**: `{gen_res.get('isd_matched', 0)}`
            - **Periode BEFORE**: `{", ".join(gen_res.get('wcl_dates', []))}`
            - **Periode AFTER**: `{", ".join(gen_res.get('post_dates', []))}`
            """)
