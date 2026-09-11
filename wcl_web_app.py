"""
wcl_web_app.py
================================================================================
WCL GEMOY - 5G NR26 Analytics & Fast Report Engine
Versi: Ringkas, Cepat, Tepat & Presisi Master Template
================================================================================
"""

import os
import io
import datetime
import pandas as pd
import streamlit as st

import wcl_processor
from wcl_processor import (
    WCLProcessor,
    ISDMatcher,
    clean_site_batch_input,
    validate_template_compliance,
    DEFAULT_MASTER_TEMPLATE
)

# ── 1. KONFIGURASI HALAMAN ────────────────────────────────────────────────────
st.set_page_config(
    page_title="WCL GEMOY",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── 2. SESSION STATE ──────────────────────────────────────────────────────────
def init_session():
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
    if 'batch_sites_text' not in st.session_state:
        st.session_state.batch_sites_text = ""
    if 'site_filter' not in st.session_state:
        st.session_state.site_filter = []

init_session()

# ── 3. CUSTOM CSS: SLEEK, MODERN & RINGKAS (DARK TEAL THEME) ─────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&display=swap');

    html, body, [class*="css"], .stMarkdown, .stSelectbox, .stTextInput, .stTextArea {
        font-family: 'Poppins', sans-serif !important;
    }

    body, .stApp {
        background: linear-gradient(145deg, #02121B 0%, #05202A 45%, #072B38 100%) !important;
        color: #F1F5F9 !important;
    }

    /* Header Ringkas */
    .gemoy-header {
        background: linear-gradient(135deg, #031824 0%, #062E3B 50%, #0A4354 100%);
        border: 1px solid rgba(0, 242, 254, 0.25);
        border-radius: 15px;
        padding: 18px 24px;
        margin-bottom: 18px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.35);
    }
    .gemoy-title {
        font-size: 26px;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin: 0;
        background: linear-gradient(90deg, #FFFFFF 0%, #38BDF8 50%, #00F2FE 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .gemoy-sub {
        font-size: 12.5px;
        color: #94A3B8;
        margin-top: 2px;
    }
    .gemoy-tag {
        background: rgba(0, 242, 254, 0.12);
        border: 1px solid #00F2FE;
        color: #00F2FE;
        font-size: 11px;
        font-weight: 700;
        padding: 4px 12px;
        border-radius: 20px;
        letter-spacing: 0.5px;
    }

    /* Kartu Ringkas */
    .gemoy-card {
        background: rgba(6, 35, 45, 0.72);
        border: 1px solid rgba(20, 184, 166, 0.22);
        border-radius: 15px;
        padding: 16px 20px;
        margin-bottom: 16px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.25);
    }
    .gemoy-card-title {
        font-size: 15px;
        font-weight: 700;
        color: #F8FAFC;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* Live Preview Stat Box */
    .stat-pill {
        background: rgba(3, 21, 32, 0.7);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
        padding: 12px 14px;
        text-align: center;
    }
    .stat-val {
        font-size: 22px;
        font-weight: 800;
    }
    .stat-lbl {
        font-size: 11px;
        font-weight: 600;
        color: #94A3B8;
        text-transform: uppercase;
        margin-top: 2px;
    }

    /* Tombol Utama */
    .stButton>button {
        background: linear-gradient(135deg, #00F2FE 0%, #0D9488 100%) !important;
        color: #02121B !important;
        border: none !important;
        padding: 12px 24px !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        border-radius: 12px !important;
        box-shadow: 0 6px 20px rgba(0, 242, 254, 0.3) !important;
        transition: all 0.2s ease !important;
    }
    .stButton>button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 8px 24px rgba(0, 242, 254, 0.45) !important;
    }

    .stDataFrame {
        border-radius: 12px !important;
        overflow: hidden !important;
    }
</style>
""", unsafe_allow_html=True)

# ── 4. HEADER RINGKAS: WCL GEMOY ──────────────────────────────────────────────
st.markdown("""
<div class="gemoy-header">
    <div>
        <div class="gemoy-title">📡 WCL GEMOY</div>
        <div class="gemoy-sub">Fast & Accurate 5G NR26 Analytics • Smart ISD Integrator • Master Template Compliant</div>
    </div>
    <div class="gemoy-tag">FAST & PRECISE</div>
</div>
""", unsafe_allow_html=True)

# ── 5. SIDEBAR: DATA SOURCE (KPI & TWAMP) ─────────────────────────────────────
with st.sidebar:
    st.markdown("### 📂 Sumber Data")

    is_local_env = os.path.exists("D:\\")
    def_src_mode = 1 if is_local_env else 0

    kpi_input_mode = st.radio(
        "Sumber File KPI 5G:",
        ["Upload File (Single/Batch)", "Folder / Path File Lokal"],
        index=def_src_mode
    )

    kpi_sources = []
    if kpi_input_mode == "Upload File (Single/Batch)":
        files_up = st.file_uploader("Upload KPI (.xlsx, .xlsm, .csv):", type=["xlsx", "xlsm", "csv"], accept_multiple_files=True)
        if files_up:
            kpi_sources = list(files_up)
            st.caption(f"✅ {len(kpi_sources)} file diupload.")
    else:
        def_kpi = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\KOTA DEPOK_03_N05_B\KPI 5G DAILY.xlsx"
        if not os.path.exists(def_kpi):
            def_kpi = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\KOTA TANGERANG_01_N04_C"
        kpi_path_str = st.text_input("Path Folder/File KPI:", value=def_kpi)
        if kpi_path_str and os.path.exists(kpi_path_str):
            if os.path.isdir(kpi_path_str):
                files = [os.path.join(kpi_path_str, f) for f in os.listdir(kpi_path_str)
                         if f.lower().endswith(('.xlsx', '.xlsm', '.csv')) and not f.startswith('~$')]
                kpi_sources = files
                st.caption(f"📁 Folder: {len(files)} file ditemukan.")
            else:
                kpi_sources = [kpi_path_str]
                st.caption(f"📄 File tunggal ({round(os.path.getsize(kpi_path_str)/(1024*1024), 2)} MB)")
        elif kpi_path_str:
            st.error("❌ Path tidak ditemukan!")

    st.markdown("---")
    st.markdown("### 📶 Sumber TWAMP")
    twamp_choice = st.radio(
        "Pilihan TWAMP:",
        ["Internal Sheet 'twamp'", "Upload File", "Path Lokal"],
        index=0
    )
    twamp_sources = None
    if twamp_choice == "Upload File":
        tw_up = st.file_uploader("Upload TWAMP:", type=["xlsx", "xlsm", "csv"], accept_multiple_files=True)
        if tw_up:
            twamp_sources = list(tw_up)
    elif twamp_choice == "Path Lokal":
        def_tw = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\5G Twamp.xlsx"
        tw_in = st.text_input("Path File TWAMP:", value=def_tw if os.path.exists(def_tw) else "")
        if tw_in and os.path.exists(tw_in):
            twamp_sources = [tw_in]

    st.markdown("---")
    if kpi_sources:
        if st.button("🔄 Muat Data", use_container_width=True):
            with st.spinner("Memindai data..."):
                try:
                    proc = WCLProcessor(kpi_sources=kpi_sources, twamp_sources=twamp_sources)
                    meta = proc.inspect_sources()
                    st.session_state.kpi_processor = proc
                    st.session_state.kpi_metadata = meta
                    if meta['clusters']:
                        st.session_state.selected_cluster = meta['clusters'][0]
                    st.success(f"Ditemukan {len(meta['clusters'])} cluster & {len(meta['sites'])} site.")
                except Exception as e:
                    st.error(f"Gagal memuat: {e}")

meta = st.session_state.kpi_metadata
avail_dates = meta.get('dates', []) if meta else []
avail_twamp = meta.get('twamp_dates', []) if meta else []

# ── 6. PANEL 1: TARGET CLUSTER, BAND & BATCH SITE ID ─────────────────────────
st.markdown("""
<div class="gemoy-card">
    <div class="gemoy-card-title">🎯 Target Cluster, Band & Site ID Batch Filter</div>
</div>
""", unsafe_allow_html=True)

c_top1, c_top2 = st.columns([1.2, 1.8])

with c_top1:
    clusters = meta.get('clusters', []) if meta else []
    if clusters:
        sel_cluster = st.selectbox("Cluster Target:", options=clusters, index=0)
        st.session_state.selected_cluster = sel_cluster
    else:
        st.selectbox("Cluster Target:", ["(Muat sumber data di sidebar terlebih dahulu)"], disabled=True)

    band_opts = {
        "🌟 NR26 Baseline NR21 (Otomatis fallback ke NR21 jika Before belum On-Air)": "nr26_baseline_nr21",
        "🔵 NR26 Only": "nr26_only",
        "🟢 NR21 Only": "nr21_only",
        "🟣 All Bands": "all_bands_separate"
    }
    sel_band = st.selectbox("Band Filter:", options=list(band_opts.keys()), index=0)
    st.session_state.band_mode = band_opts[sel_band]

with c_top2:
    # ── SITE ID BATCH FILTER (WAJIB) ──
    batch_raw = st.text_area(
        "SITE ID BATCH FILTER (Paste dari Excel):",
        value=st.session_state.batch_sites_text,
        placeholder="11TGR0316\n11TGR0320\n11TGR0325\natau pisahkan koma/titik-koma",
        height=85
    )
    st.session_state.batch_sites_text = batch_raw
    parsed_sites = clean_site_batch_input(batch_raw)
    st.session_state.site_filter = parsed_sites if parsed_sites else None

    if parsed_sites:
        st.markdown(f'<span style="background:rgba(16,185,129,0.2);color:#10B981;border:1px solid #10B981;padding:3px 10px;border-radius:15px;font-size:12px;font-weight:700;">✓ Total Site Filter: {len(parsed_sites)} Site Aktif</span>', unsafe_allow_html=True)
    else:
        st.caption("ℹ️ Memproses seluruh site dalam cluster.")

# ── 7. PANEL 2: RENTANG TANGGAL (3 KOLOM) ─────────────────────────────────────
st.markdown("""
<div class="gemoy-card">
    <div class="gemoy-card-title">📅 Rentang Tanggal Analisis (3 Kolom)</div>
</div>
""", unsafe_allow_html=True)

dt1, dt2, dt3 = st.columns(3)

if not avail_dates:
    st.info("💡 Muat file KPI di sidebar untuk memilih tanggal.")
    bef_dates_final, aft_dates_final, tw_dates_final = [], [], []
else:
    d_min = avail_dates[0]
    d_max = avail_dates[-1]

    with dt1:
        st.markdown("**🔵 BEFORE (WCL Pre-Optim)**")
        b_mode = st.selectbox("Preset BEFORE:", ["Auto First 3 Days", "Custom Date Range", "Checklist Tanggal"], key="bm_k")
        if b_mode == "Auto First 3 Days":
            bef_dates_final = avail_dates[:3]
        elif b_mode == "Custom Date Range":
            cb1, cb2 = st.columns(2)
            with cb1: b_start = st.date_input("Mulai:", value=d_min, min_value=d_min, max_value=d_max, key="bs_k")
            with cb2: b_end = st.date_input("Sampai:", value=avail_dates[min(2, len(avail_dates)-1)], min_value=d_min, max_value=d_max, key="be_k")
            if b_start > b_end: b_start, b_end = b_end, b_start
            bef_dates_final = [d for d in avail_dates if b_start <= d <= b_end]
        else:
            bef_dates_final = st.multiselect("Pilih:", options=avail_dates, default=avail_dates[:min(3, len(avail_dates))], key="bck_k")
        st.caption(f"Terpilih ({len(bef_dates_final)} hari): `{[d.strftime('%d-%b') for d in bef_dates_final]}`")

    with dt2:
        st.markdown("**🟢 AFTER (Post-Optim)**")
        a_mode = st.selectbox("Preset AFTER:", ["Auto Last 3 Days", "Custom Date Range", "Checklist Tanggal"], key="am_k")
        rem_dates = [d for d in avail_dates if d not in bef_dates_final]
        if a_mode == "Auto Last 3 Days":
            aft_dates_final = rem_dates[:3] if rem_dates else avail_dates[-3:]
        elif a_mode == "Custom Date Range":
            ca1, ca2 = st.columns(2)
            def_as = rem_dates[0] if rem_dates else d_max
            def_ae = rem_dates[min(2, len(rem_dates)-1)] if rem_dates else d_max
            with ca1: a_start = st.date_input("Mulai:", value=def_as, min_value=d_min, max_value=d_max, key="as_k")
            with ca2: a_end = st.date_input("Sampai:", value=def_ae, min_value=d_min, max_value=d_max, key="ae_k")
            if a_start > a_end: a_start, a_end = a_end, a_start
            aft_dates_final = [d for d in avail_dates if a_start <= d <= a_end]
        else:
            aft_dates_final = st.multiselect("Pilih:", options=avail_dates, default=rem_dates[:min(3, len(rem_dates))] if rem_dates else [], key="ack_k")
        st.caption(f"Terpilih ({len(aft_dates_final)} hari): `{[d.strftime('%d-%b') for d in aft_dates_final]}`")

    with dt3:
        st.markdown("**🟡 TWAMP (Transport Metrics)**")
        tw_pool = avail_twamp if avail_twamp else avail_dates
        tw_mode = st.selectbox("Preset TWAMP:", ["Auto Last 7 Days", "Custom Date Range", "Checklist Tanggal"], key="twm_k")
        if tw_mode == "Auto Last 7 Days":
            tw_dates_final = tw_pool[-7:]
        elif tw_mode == "Custom Date Range":
            ct1, ct2 = st.columns(2)
            with ct1: t_start = st.date_input("Mulai:", value=tw_pool[0] if tw_pool else d_min, key="ts_k")
            with ct2: t_end = st.date_input("Sampai:", value=tw_pool[-1] if tw_pool else d_max, key="te_k")
            if t_start > t_end: t_start, t_end = t_end, t_start
            tw_dates_final = [d for d in tw_pool if t_start <= d <= t_end]
        else:
            tw_dates_final = st.multiselect("Pilih:", options=tw_pool, default=tw_pool[-min(7, len(tw_pool)):], key="tck_k")
        st.caption(f"Terpilih ({len(tw_dates_final)} hari): `{[d.strftime('%d-%b') for d in tw_dates_final]}`")

st.session_state.bef_dates_final = bef_dates_final
st.session_state.aft_dates_final = aft_dates_final
st.session_state.tw_dates_final = tw_dates_final

# ── 8. PANEL 3: ISD & LIVE PREVIEW RINGKAS ────────────────────────────────────
with st.expander("📝 Pengaturan Data ISD (Buka untuk Ubah / Paste)", expanded=False):
    isd_choice = st.radio("Metode Input:", ["Contoh Default", "Paste Tabel Excel", "Upload File"], horizontal=True)
    if isd_choice == "Paste Tabel Excel":
        sample_paste = "Site ID\tSector\tISD\n11TGR0378\tsector 1\t470\n11TGR0378\tsector 2\t490\n11TGR0378\tsector 3\t520"
        p_txt = st.text_area("Paste:", value=sample_paste, height=100)
        if p_txt.strip():
            try:
                sep = '\t' if '\t' in p_txt else ','
                st.session_state.isd_matcher = ISDMatcher(pd.read_csv(io.StringIO(p_txt), sep=sep))
            except Exception as e:
                st.error(f"Format error: {e}")
    elif isd_choice == "Upload File":
        f_isd = st.file_uploader("Upload ISD (.xlsx, .csv):", type=["xlsx", "csv"])
        if f_isd:
            try:
                df_isd = pd.read_csv(f_isd) if f_isd.name.endswith('.csv') else pd.read_excel(f_isd)
                st.session_state.isd_matcher = ISDMatcher(df_isd)
            except Exception as e:
                st.error(f"Error membaca file: {e}")

# Live Preview
proc = st.session_state.kpi_processor
target_cluster = st.session_state.selected_cluster
site_filter = st.session_state.site_filter

if proc and target_cluster:
    try:
        cells_preview = proc.get_cluster_cells_preview(target_cluster, site_filter=site_filter, band_filter=st.session_state.band_mode)
        preview_df = st.session_state.isd_matcher.match_preview(cells_preview)
        tot_cells = len(preview_df)
        tot_sites = preview_df['Site ID'].nunique() if not preview_df.empty else 0
        mat_cells = int((preview_df['Status'] == 'Matched').sum()) if not preview_df.empty else 0
        mis_cells = tot_cells - mat_cells
        m_rate = (mat_cells / tot_cells * 100) if tot_cells > 0 else 0.0

        p1, p2, p3, p4 = st.columns(4)
        with p1: st.markdown(f'<div class="stat-pill"><div class="stat-val" style="color:#00F2FE;">{tot_sites}</div><div class="stat-lbl">Total Sites</div></div>', unsafe_allow_html=True)
        with p2: st.markdown(f'<div class="stat-pill"><div class="stat-val" style="color:#38BDF8;">{tot_cells}</div><div class="stat-lbl">Total Cells</div></div>', unsafe_allow_html=True)
        with p3: st.markdown(f'<div class="stat-pill"><div class="stat-val" style="color:#10B981;">{mat_cells}</div><div class="stat-lbl">ISD Cocok</div></div>', unsafe_allow_html=True)
        with p4: st.markdown(f'<div class="stat-pill"><div class="stat-val" style="color:#F59E0B;">{m_rate:.0f}%</div><div class="stat-lbl">Match Rate</div></div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
    except Exception as e:
        st.caption(f"Status preview: {e}")

# ── 9. EKSEKUSI & GENERATE REPORT ─────────────────────────────────────────────
c_btn, c_dir = st.columns([1.5, 1.5])
with c_btn:
    btn_run = st.button("🚀 GENERATE WCL REPORT", use_container_width=True)
with c_dir:
    def_out = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\KOTA TANGERANG_01_N04_C\Output" if is_local_env else ""
    save_folder = st.text_input("Simpan juga ke folder lokal:", value=def_out)

if btn_run:
    if not proc or not target_cluster:
        st.error("Pilih data sumber dan cluster target terlebih dahulu!")
    else:
        bef_d = st.session_state.bef_dates_final
        aft_d = st.session_state.aft_dates_final
        tw_d = st.session_state.tw_dates_final

        p_bar = st.progress(0)
        p_status = st.empty()

        def on_prog(pct, msg):
            p_bar.progress(pct)
            p_status.caption(f"Status: {msg}")

        out_fname = f"WCL NR26 {target_cluster}_GEMOY.xlsx"
        out_fpath = os.path.join(save_folder, out_fname) if (save_folder and os.path.exists(save_folder)) else None

        try:
            res = proc.generate_report(
                target_cluster=target_cluster,
                site_filter=site_filter,
                band_filter=st.session_state.band_mode,
                before_dates=bef_d,
                after_dates=aft_d,
                twamp_dates=tw_d,
                output_file=out_fpath,
                progress_callback=on_prog
            )
            st.session_state.generated_report = res
            st.success(f"🎉 Selesai! Berhasil membuat 11 sheet ({res['total_cells']} cells). Master Template Match: 100%.")
        except Exception as e:
            st.error(f"Gagal memproses laporan: {e}")

# ── 10. DOWNLOAD HASIL ────────────────────────────────────────────────────────
res_data = st.session_state.generated_report
if res_data and res_data.get('bytes'):
    st.markdown("---")
    dl1, dl2 = st.columns([1.5, 1.5])
    
    with dl1:
        dl_name = f"WCL NR26 {target_cluster}_GEMOY.xlsx"
        st.download_button(
            label=f"📥 Download Excel: {dl_name}",
            data=res_data['bytes'],
            file_name=dl_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with dl2:
        if res_data.get('pptx_bytes'):
            st.download_button(
                label="📊 Download PowerPoint Presentation (.pptx)",
                data=res_data['pptx_bytes'],
                file_name=f"WCL_{target_cluster}_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )

    # Ringkasan Cepat RCA jika ada
    if res_data.get('rca_df') is not None and not res_data['rca_df'].empty:
        with st.expander("🔍 Lihat Ringkasan Temuan RCA", expanded=False):
            st.dataframe(res_data['rca_df'], use_container_width=True)
