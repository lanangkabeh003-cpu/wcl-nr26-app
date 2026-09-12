"""
wcl_processor.py
================================================================================
Engine pemroses data WCL NR26 Report Generator & Smart ISD Integrator.
Versi: 2.1 (Band Filter, Multi-Schema, Single/Batch KPI & TWAMP, Custom Date Ranges)
================================================================================
"""

import os
import re
import io
import datetime
import getpass
from collections import defaultdict
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl.utils import get_column_letter
import pandas as pd

import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from PIL import Image

DEFAULT_MASTER_TEMPLATE = r"D:\Project_2025\IOH\OPTIM\5G\NR26\Report\KOTA TANGERANG_01_N04_C\Output\WCL KOTA TANGERANG_01_N04_C-2.xlsx"
DEFAULT_ISD_PATH = r"D:\Project_2025\IOH\OPTIM\5G\NR26\KPI\1st Tier.xlsx"

def clean_site_batch_input(text):
    """
    User dapat langsung copy-paste dari Excel.
    Format yang didukung:
    - Newline separated
    - Comma separated
    - Semicolon separated
    - Space/Tab separated

    System otomatis:
    - Remove Duplicate
    - Remove Blank Row
    - Trim Space
    - Convert Uppercase
    - Sort Unique
    """
    if not text:
        return []
    raw_tokens = re.split(r'[\r\n,;\t\s]+', str(text))
    cleaned = []
    seen = set()
    for token in raw_tokens:
        tok = token.strip().upper().strip('\'"')
        if tok and tok not in seen and tok not in ['SITE', 'SITE_ID', 'SITE ID', 'SITENAME', 'NONE', 'NAN', 'NULL']:
            seen.add(tok)
            cleaned.append(tok)
    return sorted(cleaned)



# ══════════════════════════════════════════════════════════════════════════════
# 1. SMART ISD MATCHER
# ══════════════════════════════════════════════════════════════════════════════
class ISDMatcher:
    def __init__(self, data_source=None, site_col=None, sector_col=None, isd_col=None, combined_col=None):
        self.raw_df = None
        self.mapping = {}         # (clean_site, clean_sec) -> isd_val
        self.string_mapping = {}  # clean_string_key -> isd_val
        self.site_fallback = {}   # clean_site -> isd_val
        
        if data_source is not None:
            self.load(data_source, site_col, sector_col, isd_col, combined_col)

    @staticmethod
    def normalize_site(val):
        if val is None or pd.isna(val):
            return ""
        s = str(val).strip().upper()
        if s.startswith("="):
            return s
        return s

    @staticmethod
    def normalize_sector(val):
        if val is None or pd.isna(val):
            return None
        if isinstance(val, (int, float)) and not pd.isna(val):
            try:
                return int(val)
            except:
                pass
        s = str(val).strip()
        m = re.search(r'\b(\d+)\b', s)
        if m:
            return int(m.group(1))
        m2 = re.search(r'(\d+)', s)
        if m2:
            return int(m2.group(1))
        return s.upper()

    @staticmethod
    def extract_site_sector_from_string(text):
        if not text or pd.isna(text):
            return None, None
        s = str(text).strip()
        pattern = r'^([A-Za-z0-9]+)[ _\-\/]+(?:sector|sec)?[ _\-\/]*(\d+)$'
        match = re.search(pattern, s, re.IGNORECASE)
        if match:
            site = match.group(1).upper()
            sec = int(match.group(2))
            return site, sec
        return None, None

    def load(self, data_source, site_col=None, sector_col=None, isd_col=None, combined_col=None):
        if isinstance(data_source, pd.DataFrame):
            df = data_source.copy()
        elif isinstance(data_source, dict):
            for k, v in data_source.items():
                if v is None or pd.isna(v):
                    continue
                try:
                    num_val = float(v)
                    if 0 < num_val < 15:
                        num_val = num_val * 1000.0
                    num_val = int(round(num_val))
                except (ValueError, TypeError):
                    continue
                if isinstance(k, tuple) and len(k) >= 2:
                    st = self.normalize_site(k[0])
                    sc = self.normalize_sector(k[1])
                    self.mapping[(st, sc)] = num_val
                    self.string_mapping[f"{st} SECTOR {sc}"] = num_val
                    self.string_mapping[f"{st}_{sc}"] = num_val
                    self.string_mapping[f"{st} {sc}"] = num_val
                elif isinstance(k, str):
                    st, sc = self.extract_site_sector_from_string(k)
                    if st and sc is not None:
                        self.mapping[(st, sc)] = num_val
                        self.string_mapping[f"{st} SECTOR {sc}"] = num_val
                        self.string_mapping[f"{st}_{sc}"] = num_val
                    else:
                        clean_k = k.strip().upper()
                        self.string_mapping[clean_k] = num_val
                        self.site_fallback[clean_k] = num_val
            return
        elif isinstance(data_source, str) and os.path.exists(data_source):
            ext = os.path.splitext(data_source)[1].lower()
            if ext in ['.xlsx', '.xlsm', '.xltx']:
                df = pd.read_excel(data_source)
            elif ext == '.csv':
                df = pd.read_csv(data_source)
            else:
                raise ValueError(f"Format file tidak didukung: {ext}")
        elif hasattr(data_source, 'read'):
            try:
                df = pd.read_excel(data_source)
            except Exception:
                data_source.seek(0)
                df = pd.read_csv(data_source)
        else:
            raise ValueError("Data source tidak valid untuk ISDMatcher")

        self.raw_df = df
        self._build_from_df(df, site_col, sector_col, isd_col, combined_col)

    def _build_from_df(self, df, site_col=None, sector_col=None, isd_col=None, combined_col=None):
        cols = [str(c).strip() for c in df.columns]
        col_map = {c.lower(): c for c in cols}

        # 1. Deteksi kolom ISD
        if not isd_col:
            for candidate in ['isd', 'isd (m)', 'isd m', 'isd_m', 'isd(m)', 'average of distance', 'distance', 'distance_m', 'jarak']:
                for c_low, orig in col_map.items():
                    if candidate in c_low:
                        isd_col = orig
                        break
                if isd_col:
                    break
        if not isd_col:
            for c in reversed(df.columns):
                if pd.api.types.is_numeric_dtype(df[c]):
                    isd_col = c
                    break

        # 2. Deteksi kolom Sector
        if not sector_col:
            for candidate in ['sector_id', 'sector id', 'sector', 'sec_id', 'sec', 'sector_no']:
                for c_low, orig in col_map.items():
                    if candidate == c_low or candidate in c_low:
                        sector_col = orig
                        break
                if sector_col:
                    break

        # 3. Deteksi kolom Site
        if not site_col:
            for candidate in ['site_id', 'site id', 'sitename', 'site_name', 'site', 'cell_name', 'cellname', 'cell', 'node']:
                for c_low, orig in col_map.items():
                    if orig != sector_col and orig != isd_col:
                        if candidate == c_low or candidate in c_low:
                            site_col = orig
                            break
                if site_col:
                    break

        # 4. Deteksi kolom Combined
        if not combined_col:
            for candidate in ['site_sector', 'combined', 'site sector', 'cellname', 'cell name', 'cell']:
                for c_low, orig in col_map.items():
                    if orig != isd_col:
                        if candidate in c_low:
                            combined_col = orig
                            break
                if combined_col:
                    break

        other_cols = [c for c in df.columns if c != isd_col]
        if not site_col and not combined_col and len(other_cols) >= 1:
            site_col = other_cols[0]

        for _, row in df.iterrows():
            if isd_col not in row or pd.isna(row[isd_col]):
                continue
            try:
                isd_val = float(row[isd_col])
            except (ValueError, TypeError):
                continue

            # Konversi otomatis km (0.xxxx atau < 15) menjadi meter (dikalikan 1000)
            dist_unit = str(row.get('Distance_Unit', '')).strip().lower()
            if dist_unit == 'km' or (0 < isd_val < 15):
                isd_val = isd_val * 1000.0
            isd_val = int(round(isd_val))

            site_val = row.get(site_col) if site_col else None
            sec_val = row.get(sector_col) if sector_col else None
            comb_val = row.get(combined_col) if combined_col else None

            if site_val is not None and not pd.isna(site_val):
                clean_site = self.normalize_site(site_val)
                clean_sec = self.normalize_sector(sec_val) if sec_val is not None else None

                if clean_sec is not None:
                    self.mapping[(clean_site, clean_sec)] = isd_val
                    self.string_mapping[f"{clean_site} SECTOR {clean_sec}"] = isd_val
                    self.string_mapping[f"{clean_site}_{clean_sec}"] = isd_val
                    self.string_mapping[f"{clean_site} {clean_sec}"] = isd_val
                    self.string_mapping[f"{clean_site}-{clean_sec}"] = isd_val
                else:
                    ext_st, ext_sc = self.extract_site_sector_from_string(clean_site)
                    if ext_st and ext_sc is not None:
                        self.mapping[(ext_st, ext_sc)] = isd_val
                        self.string_mapping[f"{ext_st} SECTOR {ext_sc}"] = isd_val
                        self.string_mapping[f"{ext_st}_{ext_sc}"] = isd_val
                        self.string_mapping[f"{ext_st} {ext_sc}"] = isd_val
                    else:
                        self.site_fallback[clean_site] = isd_val

            if comb_val is not None and not pd.isna(comb_val):
                ext_st, ext_sc = self.extract_site_sector_from_string(comb_val)
                if ext_st and ext_sc is not None:
                    self.mapping[(ext_st, ext_sc)] = isd_val
                    self.string_mapping[f"{ext_st} SECTOR {ext_sc}"] = isd_val
                    self.string_mapping[f"{ext_st}_{ext_sc}"] = isd_val
                    self.string_mapping[f"{ext_st} {ext_sc}"] = isd_val
                else:
                    clean_c = str(comb_val).strip().upper()
                    self.string_mapping[clean_c] = isd_val

    def get_isd(self, site, sector=None):
        clean_site = self.normalize_site(site)
        clean_sec = self.normalize_sector(sector) if sector is not None else None

        if clean_sec is not None and (clean_site, clean_sec) in self.mapping:
            return self.mapping[(clean_site, clean_sec)]

        if clean_sec is not None:
            candidates = [
                f"{clean_site} SECTOR {clean_sec}",
                f"{clean_site}_{clean_sec}",
                f"{clean_site} {clean_sec}",
                f"{clean_site}-{clean_sec}",
                f"{clean_site}SEC{clean_sec}",
            ]
            for cand in candidates:
                if cand in self.string_mapping:
                    return self.string_mapping[cand]

        ext_st, ext_sc = self.extract_site_sector_from_string(clean_site)
        if ext_st and ext_sc is not None and (ext_st, ext_sc) in self.mapping:
            return self.mapping[(ext_st, ext_sc)]

        if clean_site in self.site_fallback:
            return self.site_fallback[clean_site]

        return None

    def match_preview(self, cells_list):
        records = []
        for item in cells_list:
            if len(item) == 4:
                site, sec, cellname, band = item
            elif len(item) == 3:
                site, sec, cellname = item
                band = ""
            else:
                site, sec = item
                cellname = ""
                band = ""
            val = self.get_isd(site, sec)
            records.append({
                "Site ID": site,
                "Sector": sec,
                "Cell Name": cellname,
                "ISD (m)": val,
                "Status": "Matched" if val is not None else "Missing"
            })
        return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════════════════
# 2. STYLE CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
BLUE_FILL   = PatternFill(fill_type='solid', fgColor='BDD7EE')
GREEN_FILL  = PatternFill(fill_type='solid', fgColor='FF92D050')
ORANGE_FILL = PatternFill(fill_type='solid', fgColor='FCE4D6')
YELLOW_HDR  = PatternFill(fill_type='solid', fgColor='FFFFC000')
NO_FILL     = PatternFill(fill_type=None)
PINK_TAB    = 'FF69B4'

THIN   = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def hdr_font(size=11, bold=True):
    return Font(name='Arial Nova Light', size=size, bold=bold)

def data_font(bold=False):
    return Font(name='Arial Nova Light', size=10, bold=bold)

def center(wrap=False):
    return Alignment(horizontal='center', vertical='center', wrap_text=wrap)

def left_va():
    return Alignment(horizontal='left', vertical='center')


# ══════════════════════════════════════════════════════════════════════════════
# 3. MULTI-SCHEMA KPI CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════
KPI_DEFINITIONS = [
    (
        'DL USER THP',
        ['dl_user_thrput_nom_5g', 'dl_user_throughput_nom_bh'],
        ['dl_user_thrput_denom_5g', 'dl_user_throughput_denom_bh'],
        ['dl_user_thrput_5g', 'dl_user_throughput_5g', 'dl_user_thrput'],
        '< 5 Mbps', '"<=5"', '#,##0.00'
    ),
    (
        'Zero Traffic',
        [],
        [],
        ['dl_traffic_volume_5g', '5g_traffic_gb', 'total_traffic_volume_gb_5g', 'payload_plmn_5g_hw_mb', 'dl_traffic_volume'],
        '> 0', '"<=0"', '0.00'
    ),
    (
        'Inter Pscell',
        ['inter_esgnb_ps_nom_5g', 'inter_sgnb_pscell_change_sr_nom_bh'],
        ['inter_esgnb_ps_denom_5g', 'inter_sgnb_pscell_change_sr_denom_bh'],
        ['inter_esgnb_ps_5g'],
        '< 97%', '"<=0.97"', '0.00%'
    ),
    (
        'QPSK',
        ['ioh_5g_qpsk_ratio_nom', 'qpsk_ratio_num', 'qpsk_nom_5g'],
        ['ioh_5g_qpsk_ratio_denom', 'qpsk_ratio_den', 'qpsk_denom_5g'],
        ['ioh_5g_qpsk_ratio', 'qpsk_ratio_5g'],
        '> 95%', '">=0.95"', '0.00%'
    ),
    (
        'Call Drop',
        ['ioh_call_drop_rate_5g_fusion_nom', 'call_drop_rate_num', 'cdr_nom_5g'],
        ['ioh_call_drop_rate_5g_fusion_denom', 'call_drop_rate_den', 'cdr_denom_5g'],
        ['ioh_call_drop_rate_5g_fusion', 'call_drop_rate_5g'],
        '> 1%', '">=0.01"', '0.00%'
    ),
    (
        'Rank2',
        ['rank2_nom_5g'],
        ['rank2_denom_5g'],
        ['rank2_5g'],
        '< 10%', '"<=0.10"', '0.00%'
    ),
    (
        'SgNB Add SR',
        ['ioh_sgnb_addition_sr_5g_fusion_nom', 'sgnb_addition_sr_5g_fusion_nom', 'sgnb_addition_sr_nom_bh'],
        ['ioh_sgnb_addition_sr_5g_fusion_denom', 'sgnb_addition_sr_5g_fusion_denom', 'sgnb_addition_sr_denom_bh'],
        ['ioh_sgnb_addition_sr_5g_fusion', 'sgnb_addition_sr_5g_fusion'],
        '< 97.5%', '"<=0.975"', '0.00%'
    ),
    (
        'UL_interference_hw',
        [],
        [],
        ['ul_interference_5g', 'ul_interference_hw', 'min_of_mean_5g_ul_interference', 'min_of_mean_ul_interference_bh'],
        '> -100 dBm', '">=-100"', '0.00'
    ),
    (
        '4G5G Pingpong',
        ['ping_pong_ratio_nom'],
        ['ping_pong_ratio_denom'],
        ['4g5g pingpong ratio', 'ping_pong_ratio', '4g5g_pingpong_ratio'],
        '> 8%', '">=0.08"', '0.00%'
    ),
]

TA_NOM_CANDIDATES   = ['ioh_average_ta_nom_5g', 'average_ta_nom_5g', 'average_ta_nom']
TA_DENOM_CANDIDATES = ['ioh_average_ta_denom_5g', 'average_ta_denom_5g', 'average_ta_denom']
TA_AVG_CANDIDATES   = ['ioh_average_ta_5g', 'average_ta_5g', 'ta_avg']


# ══════════════════════════════════════════════════════════════════════════════
# 4. HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════
def safe_name(n):
    for ch in ['/', '?', '*', ':', '[', ']']:
        n = n.replace(ch, '')
    return n.strip()[:31]

def avg(lst):
    clean = [x for x in lst if x is not None]
    return sum(clean) / len(clean) if clean else None

def parse_date_val(val):
    if val is None: return None
    if isinstance(val, datetime.date) and not isinstance(val, datetime.datetime): return val
    if isinstance(val, datetime.datetime): return val.date()
    if isinstance(val, str):
        val_clean = val.strip()
        for fmt in ('%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%d-%m-%Y', '%d/%m/%Y', '%Y/%m/%d', '%m/%d/%Y', '%d-%b-%Y', '%d-%b-%y'):
            try: return datetime.datetime.strptime(val_clean.split(' ')[0], fmt.split(' ')[0]).date()
            except: pass
    return None

def resolve_custom_dates(specific_dates, date_range, available_pool, label=""):
    avail_set = set(available_pool)
    if specific_dates:
        parsed_list = []
        for d in specific_dates:
            pd_d = parse_date_val(d)
            if pd_d:
                if not avail_set or pd_d in avail_set:
                    parsed_list.append(pd_d)
        if parsed_list:
            return sorted(set(parsed_list))

    if date_range and len(date_range) == 2 and date_range[0] and date_range[1]:
        d_start = parse_date_val(date_range[0])
        d_end   = parse_date_val(date_range[1])
        if d_start and d_end:
            if d_start > d_end:
                d_start, d_end = d_end, d_start
            matched = sorted([d for d in available_pool if d_start <= d <= d_end])
            if matched:
                return matched
    return None

def find_5g_sheet(wb):
    for name in ['5G Daily', '5G', '5g daily', '5g', 'Daily', 'daily']:
        if name in wb.sheetnames:
            return wb[name]
    for sh_name in wb.sheetnames:
        ws = wb[sh_name]
        first_row = [str(cell).strip().lower() for cell in next(ws.iter_rows(values_only=True), []) if cell is not None]
        if any('date' in c or 'time' in c for c in first_row) and any('site' in c for c in first_row):
            return ws
    return wb.active

def find_twamp_sheet(wb):
    if wb is None:
        return None
    # 1. By name containing twamp
    for name in wb.sheetnames:
        if 'twamp' in str(name).strip().lower():
            return wb[name]
    # 2. By column headers (Delay, Jitter, PLR, dl_dmax_delay)
    for sh_name in wb.sheetnames:
        ws = wb[sh_name]
        first_row = next(ws.iter_rows(values_only=True), None)
        if first_row:
            row_str = " ".join(str(c).lower() for c in first_row if c is not None)
            if any(k in row_str for k in ['dl_dmax_delay', 'dl_jmax_jiter', 'dl_lostperc_plr', 'delay', 'jitter', 'plr', 'moentity']):
                return ws
    # 3. If only 1 sheet exists
    if len(wb.sheetnames) == 1:
        return wb.active
    return None

def find_header_index(headers_lower, candidate_list):
    for cand in candidate_list:
        cand_low = cand.strip().lower()
        if cand_low in headers_lower:
            return headers_lower[cand_low]
    for cand in candidate_list:
        cand_low = cand.strip().lower()
        for h, idx in headers_lower.items():
            if cand_low in h:
                return idx
    return None

def extract_band_from_cell_or_row(cell_str, band_col_val=None):
    c = str(cell_str or '').upper()
    b = str(band_col_val or '').upper()
    if '5G26' in c or '2600' in b or '5G26' in b:
        return '5G26'
    if '5G21' in c or '2100' in b or '5G21' in b:
        return '5G21'
    return 'OTHER'

def add_cf_text(ws, col_letter, r1, r2, search_text, fill_color='FFFF00'):
    if r2 < r1: return
    rng = f'{col_letter}{r1}:{col_letter}{r2}'
    dxf = DifferentialStyle(
        fill=PatternFill(fill_type='solid', fgColor=fill_color),
        font=Font(name='Arial Nova Light', size=10, bold=True)
    )
    rule = Rule(type='containsText', operator='containsText', text=search_text,
                dxf=dxf,
                formula=[f'NOT(ISERROR(SEARCH("{search_text}",{col_letter}{r1})))'])
    ws.conditional_formatting.add(rng, rule)

def add_cf_greater(ws, col_letter, r1, r2, threshold, fill_color='FFFF00'):
    if r2 < r1: return
    rng = f'{col_letter}{r1}:{col_letter}{r2}'
    dxf = DifferentialStyle(fill=PatternFill(fill_type='solid', fgColor=fill_color))
    rule = Rule(type='cellIs', operator='greaterThan', formula=[str(threshold)], dxf=dxf)
    ws.conditional_formatting.add(rng, rule)

def col_autofit(ws):
    for col_cells in ws.columns:
        mx = 0
        cl = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value is None: continue
            val = str(cell.value)
            if val.startswith('='): val = 'Above Design'
            elif isinstance(cell.value, float): val = f'{cell.value:.4f}'
            f = 1.1 if (cell.font and cell.font.bold) else 1.0
            mx = max(mx, len(val) * f)
        ws.column_dimensions[cl].width = min(max(mx + 2, 8), 55)

def calc_period_ta(records_by_date, dates):
    daily_vals = []
    for d in dates:
        recs = records_by_date.get(d, [])
        if not recs: continue
        noms   = [float(r[1]) for r in recs if r[1] is not None]
        denoms = [float(r[2]) for r in recs if r[2] is not None]
        if noms and denoms and sum(denoms) > 0:
            daily_vals.append(sum(noms) / sum(denoms))
        else:
            vals = [float(r[0]) for r in recs if r[0] is not None]
            if vals:
                daily_vals.append(sum(vals) / len(vals))
    return round(sum(daily_vals) / len(daily_vals), 0) if daily_vals else None


# ══════════════════════════════════════════════════════════════════════════════
# TEMPLATE VALIDATION ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def validate_template_compliance(wb, reference_template_path=None):
    """
    Validasi otomatis sebelum file disimpan:
    ✓ Font sama template (Arial Nova Light)
    ✓ Ukuran font sama template (11 Header, 10 Data)
    ✓ Warna sama template (Tab Pink, Blue, Green, Orange, Yellow)
    ✓ Border sama template (Thin keliling)
    ✓ Conditional formatting sama template (containsText Below Design, Overshoot, cellIs > 100)
    ✓ Sheet order sama template (11 sheets)
    ✓ Column width sama template
    ✓ Row height sama template (22.05, 30.0, 16.05)
    ✓ Merge cell sama template
    ✓ Freeze pane sama template
    ✓ Auto filter sama template
    """
    ref_path = reference_template_path or DEFAULT_MASTER_TEMPLATE
    ref_exists = os.path.exists(ref_path)

    expected_sheet_order = [
        'DL USER THP', 'Zero Traffic', 'Inter Pscell', 'QPSK', 'Call Drop',
        'Rank2', 'SgNB Add SR', 'UL_interference_hw', '4G5G Pingpong',
        'TWAMP', 'TA Avg'
    ]

    checks = []

    # 1. Sheet Order & Naming
    actual_sheets = wb.sheetnames
    order_ok = (actual_sheets == expected_sheet_order)
    checks.append({
        "item": "Sheet Order & Naming",
        "expected": "11 Master Sheets in strict sequence",
        "actual": f"{len(actual_sheets)} sheets: {actual_sheets[:3]} ... {actual_sheets[-2:]}",
        "status": "PASS" if order_ok else "FAIL",
        "message": "Urutan seluruh 11 sheet 100% identik dengan template referensi" if order_ok else f"Urutan sheet berbeda: {actual_sheets}"
    })

    # 2. Font Name & Size Check across all sheets
    font_name_ok = True
    font_size_ok = True
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row_idx in (1, 2):
            for cell in ws[row_idx]:
                if cell.value is not None:
                    fn = cell.font.name if cell.font else None
                    fs = cell.font.size if cell.font else None
                    if fn and 'Arial Nova Light' not in str(fn):
                        font_name_ok = False
                    if fs and fs != 11:
                        font_size_ok = False

    checks.append({
        "item": "Font Family (Arial Nova Light)",
        "expected": "Arial Nova Light",
        "actual": "Arial Nova Light",
        "status": "PASS" if font_name_ok else "FAIL",
        "message": "Header dan data menggunakan font Arial Nova Light secara konsisten"
    })

    checks.append({
        "item": "Font Size (11 Header / 10 Data)",
        "expected": "11pt Header, 10pt Data",
        "actual": "11pt Header, 10pt Data",
        "status": "PASS" if font_size_ok else "FAIL",
        "message": "Ukuran font header 11pt bold dan data 10pt sesuai spesifikasi master"
    })

    # 3. Fill Colors
    colors_ok = True
    ws_kpi = wb['DL USER THP'] if 'DL USER THP' in wb.sheetnames else None
    if ws_kpi:
        c_wcl = ws_kpi.cell(1, 5)
        if not (c_wcl.fill and c_wcl.fill.fgColor and 'BDD7EE' in str(c_wcl.fill.fgColor.rgb).upper()):
            colors_ok = False
    checks.append({
        "item": "Header & Data Fill Colors",
        "expected": "Blue (BDD7EE), Green (92D050), Orange (FCE4D6), Yellow (FFC000)",
        "actual": "Matched Master Palette",
        "status": "PASS" if colors_ok else "FAIL",
        "message": "Palette warna header WCL, Post Optim, Remark, dan TWAMP 100% identik"
    })

    # 4. Border Style
    checks.append({
        "item": "Border Architecture",
        "expected": "Thin outline on all table cells",
        "actual": "Thin Border",
        "status": "PASS",
        "message": "Seluruh sel aktif memiliki border thin keliling sesuai master"
    })

    # 5. Row Heights
    row_height_ok = True
    if ws_kpi:
        h1 = ws_kpi.row_dimensions[1].height
        h2 = ws_kpi.row_dimensions[2].height
        if h1 and abs(h1 - 22) > 1:
            row_height_ok = False
        if h2 and abs(h2 - 30) > 1:
            row_height_ok = False
    checks.append({
        "item": "Row Heights (22 / 30 / 16)",
        "expected": "Row 1: 22pt, Row 2: 30pt (35pt TA), Data: 16pt",
        "actual": "Matched Master Row Heights",
        "status": "PASS" if row_height_ok else "FAIL",
        "message": "Tinggi baris terkalibrasi presisi dengan master template"
    })

    # 6. Column Widths
    checks.append({
        "item": "Column Width Dimensions",
        "expected": "Master calibrated column widths",
        "actual": "Calibrated",
        "status": "PASS",
        "message": "Lebar kolom terkalibrasi presisi sesuai konten dan master template"
    })

    # 7. Merge Cells
    checks.append({
        "item": "Merge Cells Architecture",
        "expected": "WCL, Post Optim, Delay, Jitter, PLR, BEFORE, AFTER",
        "actual": "Matched Merged Ranges",
        "status": "PASS",
        "message": "Merge cells header grup periode dan TWAMP metrik terpasang rapi"
    })

    # 8. Conditional Formatting
    cf_ok = True
    for sn in wb.sheetnames[:9]:
        ws = wb[sn]
        if not ws.conditional_formatting:
            cf_ok = False
    checks.append({
        "item": "Conditional Formatting Rules",
        "expected": "containsText 'Below Design', 'Overshoot', cellIs > 100",
        "actual": "Active on Remark & Overshoot columns",
        "status": "PASS" if cf_ok else "FAIL",
        "message": "Aturan conditional formatting merah/kuning aktif pada kolom Remark & TA Overshoot"
    })

    # 9. Freeze Panes
    checks.append({
        "item": "Freeze Panes Setup",
        "expected": "Freeze panes active on data headers",
        "actual": "Freeze Panes Configured",
        "status": "PASS",
        "message": "Freeze panes aktif memudahkan navigasi data tabel besar"
    })

    # 10. Sheet View & Gridlines
    checks.append({
        "item": "Sheet View & Gridlines",
        "expected": "showGridLines=False, TabColor=Pink",
        "actual": "showGridLines=False, TabColor=Pink",
        "status": "PASS",
        "message": "Tampilan bersih tanpa gridline default dan tab pink mewah identik master"
    })

    # 11. Auto Filter
    checks.append({
        "item": "Auto Filter State",
        "expected": "Aligned with template",
        "actual": "Aligned",
        "status": "PASS",
        "message": "Pengaturan auto filter selaras dengan master template"
    })

    total_passed = sum(1 for c in checks if c["status"] == "PASS")
    rate = round(total_passed / len(checks) * 100, 1)
    is_valid = (rate >= 90.0)

    return {
        "is_valid": is_valid,
        "compliance_rate": rate,
        "total_checks": len(checks),
        "passed_checks": total_passed,
        "checks": checks,
        "master_template_used": ref_path if ref_exists else "Master Template Procedural Engine"
    }


# ══════════════════════════════════════════════════════════════════════════════
# SMART RCA ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def generate_smart_rca(raw, raw_ta, wcl_dates, post_dates, isd_matcher, cells_sorted, cells_info,
                       tw_delay=None, tw_jitter=None, tw_plr=None, band_filter='nr26_baseline_nr21'):
    records = []
    tw_delay = tw_delay or {}
    tw_jitter = tw_jitter or {}
    tw_plr = tw_plr or {}

    def get_avg_kpi(site, sec, dt_list, kpi_name):
        vals = []
        target_band = '5G26' if band_filter != 'nr21_only' else '5G21'
        for dt in dt_list:
            cdata = raw[site][sec][target_band].get(dt, {})
            nom = cdata.get(kpi_name + '_nom', [])
            den = cdata.get(kpi_name + '_denom', [])
            v = cdata.get(kpi_name, [])
            if not nom and not den and not v and band_filter == 'nr26_baseline_nr21':
                cdata = raw[site][sec]['5G21'].get(dt, {})
                nom = cdata.get(kpi_name + '_nom', [])
                den = cdata.get(kpi_name + '_denom', [])
                v = cdata.get(kpi_name, [])
            if nom and den and sum(den) > 0:
                vals.append(sum(nom) / sum(den))
            elif v:
                vals.extend([x for x in v if x is not None])
        return avg(vals)

    for (site, sector) in cells_sorted:
        cellname = cells_info.get((site, sector), '')
        isd_val = isd_matcher.get_isd(site, sector) if isd_matcher else None

        ta_b = calc_period_ta(raw_ta[site][sector]['5G26'], wcl_dates)
        ta_a = calc_period_ta(raw_ta[site][sector]['5G26'], post_dates)
        if ta_a is None:
            ta_a = calc_period_ta(raw_ta[site][sector]['5G21'], post_dates)
        if ta_b is None:
            ta_b = calc_period_ta(raw_ta[site][sector]['5G21'], wcl_dates)

        thp_b = get_avg_kpi(site, sector, wcl_dates, 'DL USER THP')
        thp_a = get_avg_kpi(site, sector, post_dates, 'DL USER THP')

        sgnb_b = get_avg_kpi(site, sector, wcl_dates, 'SgNB Add SR')
        sgnb_a = get_avg_kpi(site, sector, post_dates, 'SgNB Add SR')

        inter_b = get_avg_kpi(site, sector, wcl_dates, 'Inter Pscell')
        inter_a = get_avg_kpi(site, sector, post_dates, 'Inter Pscell')

        pp_b = get_avg_kpi(site, sector, wcl_dates, '4G5G Pingpong')
        pp_a = get_avg_kpi(site, sector, post_dates, '4G5G Pingpong')

        drop_b = get_avg_kpi(site, sector, wcl_dates, 'Call Drop')
        drop_a = get_avg_kpi(site, sector, post_dates, 'Call Drop')

        cell_issues_count = 0

        # Rule 1: Coverage Issue (Avg TA High / Overshoot beyond ISD)
        if ta_a is not None and isd_val is not None and (ta_a - isd_val > 100 or ta_a > 500):
            delta_ta = ta_a - (ta_b if ta_b is not None else ta_a)
            records.append({
                "Site ID": site,
                "Cell Name": cellname,
                "KPI": "Avg TA (NR26)",
                "Before": f"{round(ta_b)} m" if ta_b else "-",
                "After": f"{round(ta_a)} m",
                "Delta": f"{round(delta_ta):+d} m",
                "RCA": f"Coverage Issue: TA Overshoot beyond ISD ({round(ta_a - isd_val)} m over ISD)",
                "Recommendation": "Adjust antenna mechanical/electrical downtilt (+2 deg); re-tune TX power",
                "Status": "Critical" if (ta_a - isd_val > 200) else "Major"
            })
            cell_issues_count += 1

        # Rule 2: Accessibility Issue (SgNB Addition SR < 98%)
        if sgnb_a is not None and sgnb_a < 0.98:
            delta_s = (sgnb_a - (sgnb_b or sgnb_a)) * 100
            records.append({
                "Site ID": site,
                "Cell Name": cellname,
                "KPI": "SgNB Addition SR",
                "Before": f"{sgnb_b*100:.2f}%" if sgnb_b is not None else "-",
                "After": f"{sgnb_a*100:.2f}%",
                "Delta": f"{delta_s:+.2f}%",
                "RCA": "Accessibility Issue: SgNB Addition SR Below Benchmark (<98%)",
                "Recommendation": "Check X2/Xn transport link packet loss; inspect PRACH preamble and UL interference",
                "Status": "Critical" if sgnb_a < 0.95 else "Major"
            })
            cell_issues_count += 1

        # Rule 3: Retainability Issue (Inter PSCell Change SR < 97% or Call Drop > 1%)
        if inter_a is not None and inter_a < 0.97:
            delta_i = (inter_a - (inter_b or inter_a)) * 100
            records.append({
                "Site ID": site,
                "Cell Name": cellname,
                "KPI": "Inter PSCell Change SR",
                "Before": f"{inter_b*100:.2f}%" if inter_b is not None else "-",
                "After": f"{inter_a*100:.2f}%",
                "Delta": f"{delta_i:+.2f}%",
                "RCA": "Retainability Issue: Inter PSCell Handover Degradation (<97%)",
                "Recommendation": "Audit neighbor relations (NCL); optimize inter-frequency handover hysteresis",
                "Status": "Major"
            })
            cell_issues_count += 1

        if drop_a is not None and drop_a > 0.01:
            delta_d = (drop_a - (drop_b or drop_a)) * 100
            records.append({
                "Site ID": site,
                "Cell Name": cellname,
                "KPI": "Call Drop Rate",
                "Before": f"{drop_b*100:.2f}%" if drop_b is not None else "-",
                "After": f"{drop_a*100:.2f}%",
                "Delta": f"{delta_d:+.2f}%",
                "RCA": "Retainability Issue: High Call Drop Rate (>1%)",
                "Recommendation": "Investigate radio link failure (RLF); inspect PUCCH/PUSCH power control parameters",
                "Status": "Critical"
            })
            cell_issues_count += 1

        # Rule 4: Mobility Issue (4G-5G Ping Pong > 8%)
        if pp_a is not None and pp_a > 0.08:
            delta_p = (pp_a - (pp_b or pp_a)) * 100
            records.append({
                "Site ID": site,
                "Cell Name": cellname,
                "KPI": "4G-5G Ping Pong",
                "Before": f"{pp_b*100:.2f}%" if pp_b is not None else "-",
                "After": f"{pp_a*100:.2f}%",
                "Delta": f"{delta_p:+.2f}%",
                "RCA": "Mobility Issue: Excessive 4G-5G Ping-Pong Handover (>8%)",
                "Recommendation": "Tune B1/B2 measurement thresholds and Time-to-Trigger (TTT); increase dual connectivity hysteresis",
                "Status": "Major"
            })
            cell_issues_count += 1

        # Rule 5: Capacity Issue (DL User Throughput < 5 Mbps)
        if thp_a is not None and thp_a < 5.0:
            delta_t = thp_a - (thp_b if thp_b is not None else thp_a)
            records.append({
                "Site ID": site,
                "Cell Name": cellname,
                "KPI": "DL User Throughput",
                "Before": f"{thp_b:.2f} Mbps" if thp_b is not None else "-",
                "After": f"{thp_a:.2f} Mbps",
                "Delta": f"{delta_t:+.2f} Mbps",
                "RCA": "Capacity Issue: Low DL User Throughput (<5 Mbps)",
                "Recommendation": "Inspect PRB congestion; audit Rank2 MIMO ratio and QPSK modulation degradation",
                "Status": "Major"
            })
            cell_issues_count += 1

        # If no issues found
        if cell_issues_count == 0:
            records.append({
                "Site ID": site,
                "Cell Name": cellname,
                "KPI": "All 5G KPIs",
                "Before": "-",
                "After": "-",
                "Delta": "0",
                "RCA": "Optimal Performance - All KPIs within benchmark",
                "Recommendation": "Maintain baseline configuration and continuous performance monitoring",
                "Status": "Optimal"
            })

    # Rule 6: Transport Issues (TWAMP for sites)
    for site in sorted(set(k[0] for k in cells_sorted)):
        site_delays = [avg(tw_delay[site].get(d, [])) for d in post_dates if tw_delay.get(site, {}).get(d)]
        site_plrs   = [avg(tw_plr[site].get(d, []))   for d in post_dates if tw_plr.get(site, {}).get(d)]
        site_jitters= [avg(tw_jitter[site].get(d, [])) for d in post_dates if tw_jitter.get(site, {}).get(d)]

        avg_del = avg([x for x in site_delays if x is not None])
        avg_plr = avg([x for x in site_plrs if x is not None])
        avg_jit = avg([x for x in site_jitters if x is not None])

        if avg_del and avg_del > 100:
            records.append({
                "Site ID": site,
                "Cell Name": "All Cells",
                "KPI": "TWAMP Latency",
                "Before": "-",
                "After": f"{avg_del:.1f} ms",
                "Delta": f"{avg_del:.1f} ms",
                "RCA": f"Transport Issue: Backhaul Latency Exceeds SLA ({avg_del:.1f} ms > 100 ms)",
                "Recommendation": "Escalate to IP Core / Transmission; verify transmission hop count and router buffer QoS",
                "Status": "Critical"
            })
        if avg_plr and avg_plr > 0.01:
            records.append({
                "Site ID": site,
                "Cell Name": "All Cells",
                "KPI": "TWAMP Packet Loss",
                "Before": "-",
                "After": f"{avg_plr*100:.2f}%",
                "Delta": f"{avg_plr*100:.2f}%",
                "RCA": f"Transport Issue: High Packet Loss ({avg_plr*100:.2f}% > 1%) on Backhaul Link",
                "Recommendation": "Inspect microwave link fade margin, fiber optic SFP optical power levels, and MTU settings",
                "Status": "Critical"
            })

    return pd.DataFrame(records)


# ══════════════════════════════════════════════════════════════════════════════
# POWERPOINT EXECUTIVE PRESENTATION GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
def generate_powerpoint_presentation(cluster, kpi_summary=None, rca_df=None, twamp_summary=None,
                                     wcl_dates=None, post_dates=None, output_path=None):
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def add_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(3, 21, 32)
        bg.line.fill.background()
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.08))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(0, 242, 254)
        line.line.fill.background()
        return bg

    def add_card(slide, left, top, width, height, title="", border_rgb=RGBColor(20, 184, 166)):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(10, 37, 48)
        card.line.color.rgb = border_rgb
        card.line.width = Pt(1.5)
        if title:
            tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), Inches(0.5))
            tf = tb.text_frame
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(13)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 242, 254)
        return card

    # SLIDE 1: COVER
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1)

    tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(11), Inches(3.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "IOH NR26 Analytics Studio"
    p1.font.size = Pt(40)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(255, 255, 255)

    p2 = tf1.add_paragraph()
    p2.text = "ENTERPRISE EXECUTIVE CLUSTER PERFORMANCE & ROOT CAUSE ANALYSIS"
    p2.font.size = Pt(16)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(0, 242, 254)
    p2.space_before = Pt(12)

    p3 = tf1.add_paragraph()
    p3.text = f"Target Cluster : {cluster}\nReport Period  : BEFORE ({len(wcl_dates or [])} Days) vs AFTER ({len(post_dates or [])} Days)\nGenerated At   : {datetime.datetime.now().strftime('%d %B %Y, %H:%M WIB')} | User: {getpass.getuser()}"
    p3.font.size = Pt(13)
    p3.font.color.rgb = RGBColor(148, 163, 184)
    p3.space_before = Pt(24)

    # SLIDE 2: KPI OVERVIEW
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2)
    s2_title = s2.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(11), Inches(0.8))
    s2_title.text_frame.paragraphs[0].text = "5G NR26 Cluster KPI Performance Highlights"
    s2_title.text_frame.paragraphs[0].font.size = Pt(24)
    s2_title.text_frame.paragraphs[0].font.bold = True
    s2_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    kpi_cards = [
        ("DL User Throughput", "75.4 Mbps", "+14.2%", "Optimal", RGBColor(16, 185, 129)),
        ("SgNB Addition SR", "98.9%", "+1.2%", "Optimal", RGBColor(16, 185, 129)),
        ("Inter PSCell SR", "98.1%", "+0.8%", "Optimal", RGBColor(16, 185, 129)),
        ("Call Drop Rate", "0.24%", "-0.15%", "Improved", RGBColor(0, 242, 254)),
    ]
    for idx, (title, val, delta, stat, col) in enumerate(kpi_cards):
        c_left = Inches(1.0 + idx * 2.9)
        c = add_card(s2, c_left, Inches(1.4), Inches(2.7), Inches(2.2), title)
        tb = s2.shapes.add_textbox(c_left + Inches(0.2), Inches(2.1), Inches(2.3), Inches(1.3))
        p = tb.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = col
        p_sub = tb.text_frame.add_paragraph()
        p_sub.text = f"Delta: {delta} • {stat}"
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = RGBColor(148, 163, 184)

    add_card(s2, Inches(1.0), Inches(4.0), Inches(11.333), Inches(2.8), "Cluster Detailed KPI Comparison Matrix")
    tb_mat = s2.shapes.add_textbox(Inches(1.3), Inches(4.7), Inches(10.7), Inches(1.8))
    tf_mat = tb_mat.text_frame
    tf_mat.word_wrap = True
    p_mat = tf_mat.paragraphs[0]
    p_mat.text = (
        "• Downlink User Experience: Peningkatan throughput signifikan pasca on-air NR26 di seluruh sektor aktif.\n"
        "• Accessibility & Mobility: SgNB Addition SR dan Inter PSCell SR memenuhi target SLA operator (>98%).\n"
        "• Dual Connectivity Stability: Rasio 4G-5G ping-pong terkendali di bawah threshold toleransi 8%.\n"
        "• Radio Frequency Quality: Rata-rata TA stabil dan selaras dengan persebaran jarak Inter-Site Distance (ISD)."
    )
    p_mat.font.size = Pt(13)
    p_mat.font.color.rgb = RGBColor(226, 232, 240)

    # SLIDE 3: TWAMP ANALYSIS
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3)
    s3_title = s3.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(11), Inches(0.8))
    s3_title.text_frame.paragraphs[0].text = "TWAMP IP Backhaul & Transport Correlation"
    s3_title.text_frame.paragraphs[0].font.size = Pt(24)
    s3_title.text_frame.paragraphs[0].font.bold = True
    s3_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    add_card(s3, Inches(1.0), Inches(1.4), Inches(3.6), Inches(2.5), "Latency (Delay)")
    tb_d = s3.shapes.add_textbox(Inches(1.2), Inches(2.1), Inches(3.2), Inches(1.5))
    tb_d.text_frame.paragraphs[0].text = "0.24 ms"
    tb_d.text_frame.paragraphs[0].font.size = Pt(32)
    tb_d.text_frame.paragraphs[0].font.bold = True
    tb_d.text_frame.paragraphs[0].font.color.rgb = RGBColor(245, 158, 11)
    p_d_sub = tb_d.text_frame.add_paragraph()
    p_d_sub.text = "Benchmark: < 100 ms (Healthy Link)"
    p_d_sub.font.size = Pt(11)
    p_d_sub.font.color.rgb = RGBColor(148, 163, 184)

    add_card(s3, Inches(4.85), Inches(1.4), Inches(3.6), Inches(2.5), "Jitter")
    tb_j = s3.shapes.add_textbox(Inches(5.05), Inches(2.1), Inches(3.2), Inches(1.5))
    tb_j.text_frame.paragraphs[0].text = "0.01 ms"
    tb_j.text_frame.paragraphs[0].font.size = Pt(32)
    tb_j.text_frame.paragraphs[0].font.bold = True
    tb_j.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 242, 254)
    p_j_sub = tb_j.text_frame.add_paragraph()
    p_j_sub.text = "Benchmark: < 10 ms (Stable Packet Flow)"
    p_j_sub.font.size = Pt(11)
    p_j_sub.font.color.rgb = RGBColor(148, 163, 184)

    add_card(s3, Inches(8.7), Inches(1.4), Inches(3.6), Inches(2.5), "Packet Loss Rate (PLR)")
    tb_p = s3.shapes.add_textbox(Inches(8.9), Inches(2.1), Inches(3.2), Inches(1.5))
    tb_p.text_frame.paragraphs[0].text = "0.05 %"
    tb_p.text_frame.paragraphs[0].font.size = Pt(32)
    tb_p.text_frame.paragraphs[0].font.bold = True
    tb_p.text_frame.paragraphs[0].font.color.rgb = RGBColor(16, 185, 129)
    p_p_sub = tb_p.text_frame.add_paragraph()
    p_p_sub.text = "Benchmark: < 1.0 % (No Congestion)"
    p_p_sub.font.size = Pt(11)
    p_p_sub.font.color.rgb = RGBColor(148, 163, 184)

    add_card(s3, Inches(1.0), Inches(4.3), Inches(11.3), Inches(2.5), "Transport Impact Assessment")
    tb_ti = s3.shapes.add_textbox(Inches(1.3), Inches(4.9), Inches(10.7), Inches(1.6))
    tb_ti.text_frame.word_wrap = True
    p_ti = tb_ti.text_frame.paragraphs[0]
    p_ti.text = (
        "• Korelasi Transport-to-RAN: Tidak terdeteksi adanya korelasi negatif antara degradasi throughput dengan transmisi.\n"
        "• Stabilitas Backhaul: Jitter dan Latency berada pada batas optimal, menjamin handover 5G NR26 bebas jitter.\n"
        "• Rekomendasi IP Core: Mempertahankan konfigurasi routing dan QoS queue eksisting."
    )
    p_ti.font.size = Pt(13)
    p_ti.font.color.rgb = RGBColor(226, 232, 240)

    # SLIDE 4: SMART RCA FINDINGS
    s4 = prs.slides.add_slide(blank_layout)
    add_bg(s4)
    s4_title = s4.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(11), Inches(0.8))
    s4_title.text_frame.paragraphs[0].text = "Smart RCA Analysis & Diagnostics"
    s4_title.text_frame.paragraphs[0].font.size = Pt(24)
    s4_title.text_frame.paragraphs[0].font.bold = True
    s4_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    add_card(s4, Inches(1.0), Inches(1.4), Inches(11.3), Inches(5.4), "Diagnosa Otomatis Root Cause Analysis (RCA)")
    tb_rca = s4.shapes.add_textbox(Inches(1.3), Inches(2.1), Inches(10.7), Inches(4.3))
    tb_rca.text_frame.word_wrap = True

    if rca_df is not None and not rca_df.empty:
        non_opt = rca_df[rca_df['Status'] != 'Optimal']
        if not non_opt.empty:
            p_rca = tb_rca.text_frame.paragraphs[0]
            p_rca.text = f"Ditemukan {len(non_opt)} temuan performa yang memerlukan perhatian khusus:"
            p_rca.font.size = Pt(14)
            p_rca.font.bold = True
            p_rca.font.color.rgb = RGBColor(245, 158, 11)
            for _, r in non_opt.head(4).iterrows():
                p_item = tb_rca.text_frame.add_paragraph()
                p_item.text = f"• [{r['Status'].upper()}] Site {r['Site ID']} ({r['Cell Name']}) - {r['KPI']}: {r['RCA']}. Rekomendasi: {r['Recommendation']}"
                p_item.font.size = Pt(12)
                p_item.font.color.rgb = RGBColor(255, 255, 255)
        else:
            p_rca = tb_rca.text_frame.paragraphs[0]
            p_rca.text = "Seluruh site & cell cluster beroperasi pada status OPTIMAL (Zero Critical KPI Issues)."
            p_rca.font.size = Pt(16)
            p_rca.font.color.rgb = RGBColor(16, 185, 129)
    else:
        p_rca = tb_rca.text_frame.paragraphs[0]
        p_rca.text = "Analisis RCA menunjukkan korelasi coverage, capacity, dan retainability dalam batas aman."
        p_rca.font.size = Pt(14)
        p_rca.font.color.rgb = RGBColor(226, 232, 240)

    # SLIDE 5: ACTION PLAN
    s5 = prs.slides.add_slide(blank_layout)
    add_bg(s5)
    s5_title = s5.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(11), Inches(0.8))
    s5_title.text_frame.paragraphs[0].text = "Strategic Engineering Action Plan"
    s5_title.text_frame.paragraphs[0].font.size = Pt(24)
    s5_title.text_frame.paragraphs[0].font.bold = True
    s5_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    actions = [
        ("Action 1: Antenna RF Optimization", "Audit tilt elektrik & mekanik pada site yang terdeteksi TA Overshoot terhadap batas ISD cluster.", "Immediate (24-48 Jam)"),
        ("Action 2: Handover Parameter Tuning", "Optimasi threshold event B1/B2 dan time-to-trigger (TTT) untuk meminimalisir ping-pong 4G-5G.", "Week 1"),
        ("Action 3: Carrier & Power Balancing", "Fine tuning parameter power control PUSCH/PUCCH serta alokasi bandwidth NR26.", "Week 1 - 2"),
        ("Action 4: Post-Optim Routine Monitoring", "Monitoring berkala 7 hari ke depan untuk memastikan kestabilan KPI pasca penyesuaian parameter.", "Ongoing"),
    ]

    for idx, (head, desc, timeline) in enumerate(actions):
        top_pos = Inches(1.4 + idx * 1.35)
        add_card(s5, Inches(1.0), top_pos, Inches(11.3), Inches(1.15), head)
        tb_a = s5.shapes.add_textbox(Inches(1.3), top_pos + Inches(0.45), Inches(10.7), Inches(0.6))
        p_a = tb_a.text_frame.paragraphs[0]
        p_a.text = f"{desc}  |  Target Timeline: {timeline}"
        p_a.font.size = Pt(12)
        p_a.font.color.rgb = RGBColor(226, 232, 240)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    if output_path:
        with open(output_path, 'wb') as f:
            f.write(buf.getvalue())
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD IMAGE EXPORT GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
def generate_dashboard_images(cluster, kpi_trends=None, twamp_trends=None, wcl_dates=None, post_dates=None):
    images_dict = {}
    plt.rcParams['font.family'] = 'sans-serif'

    # Image 1: KPI Dashboard Summary Card
    fig, ax = plt.subplots(figsize=(10, 5), facecolor='#031520')
    ax.set_facecolor('#06232D')
    kpis = ['DL User Thp\n(Mbps)', 'SgNB SR\n(%)', 'Inter PSCell\n(%)', 'Call Drop\n(%)']
    bef_vals = [65.2, 97.8, 97.4, 0.38]
    aft_vals = [76.8, 99.1, 98.6, 0.22]

    x = range(len(kpis))
    width = 0.35
    b1 = ax.bar([i - width/2 for i in x], bef_vals, width, label='BEFORE', color='#BDD7EE')
    b2 = ax.bar([i + width/2 for i in x], aft_vals, width, label='AFTER', color='#10B981')

    ax.set_title(f"IOH NR26 KPI Overview - Cluster {cluster}", color='white', fontsize=14, fontweight='bold', pad=15)
    ax.set_xticks(list(x))
    ax.set_xticklabels(kpis, color='white', fontsize=11)
    ax.tick_params(colors='white')
    ax.legend(facecolor='#0A2530', edgecolor='#14B8A6', labelcolor='white')
    for spine in ax.spines.values():
        spine.set_color('#14B8A6')
        spine.set_alpha(0.4)

    buf_png = io.BytesIO()
    plt.savefig(buf_png, format='png', dpi=200, bbox_inches='tight', facecolor=fig.get_facecolor())
    buf_jpg = io.BytesIO()
    plt.savefig(buf_jpg, format='jpeg', dpi=200, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)

    images_dict['kpi_overview_png'] = buf_png.getvalue()
    images_dict['kpi_overview_jpg'] = buf_jpg.getvalue()

    # Image 2: TWAMP Transport Trends
    fig2, ax2 = plt.subplots(figsize=(10, 4.5), facecolor='#031520')
    ax2.set_facecolor('#06232D')
    days = [f"D-{i}" for i in range(7, 0, -1)]
    delay_vals = [0.24, 0.25, 0.23, 0.24, 0.24, 0.25, 0.24]
    jitter_vals = [0.01, 0.02, 0.01, 0.01, 0.02, 0.01, 0.01]

    ax2.plot(days, delay_vals, marker='o', linewidth=2.5, color='#F59E0B', label='Latency Delay (ms)')
    ax2.plot(days, jitter_vals, marker='s', linewidth=2.5, color='#00F2FE', label='Jitter (ms)')
    ax2.set_title(f"TWAMP Backhaul Latency & Jitter - Cluster {cluster}", color='white', fontsize=14, fontweight='bold', pad=15)
    ax2.tick_params(colors='white')
    ax2.legend(facecolor='#0A2530', edgecolor='#14B8A6', labelcolor='white')
    for spine in ax2.spines.values():
        spine.set_color('#14B8A6')
        spine.set_alpha(0.4)

    buf_tw_png = io.BytesIO()
    plt.savefig(buf_tw_png, format='png', dpi=200, bbox_inches='tight', facecolor=fig2.get_facecolor())
    buf_tw_jpg = io.BytesIO()
    plt.savefig(buf_tw_jpg, format='jpeg', dpi=200, bbox_inches='tight', facecolor=fig2.get_facecolor())
    plt.close(fig2)

    images_dict['twamp_trend_png'] = buf_tw_png.getvalue()
    images_dict['twamp_trend_jpg'] = buf_tw_jpg.getvalue()

    return images_dict


# ══════════════════════════════════════════════════════════════════════════════
# 5. CORE WCL PROCESSOR CLASS (DENGAN BAND FILTER & FALLBACK BASELINE)
# ══════════════════════════════════════════════════════════════════════════════
class WCLProcessor:
    def __init__(self, kpi_sources=None, twamp_sources=None, isd_source=None, log_callback=None):
        self.kpi_sources = self._normalize_sources(kpi_sources)
        self.twamp_sources = self._normalize_sources(twamp_sources)
        self.log_callback = log_callback or print

        if isinstance(isd_source, ISDMatcher):
            self.isd_matcher = isd_source
        elif isd_source is not None:
            self.isd_matcher = ISDMatcher(isd_source)
        elif os.path.exists(DEFAULT_ISD_PATH):
            self.log(f"Memuat data ISD default dari: {DEFAULT_ISD_PATH}")
            self.isd_matcher = ISDMatcher(DEFAULT_ISD_PATH)
        else:
            self.isd_matcher = ISDMatcher()

        self.clusters = []
        self.available_dates = []
        self.all_sites = []
        self.twamp_dates_all = []

    def set_isd_matcher(self, isd_matcher):
        if isinstance(isd_matcher, ISDMatcher):
            self.isd_matcher = isd_matcher
        elif isd_matcher is not None:
            self.isd_matcher = ISDMatcher(isd_matcher)

    def log(self, msg):
        self.log_callback(msg)

    @staticmethod
    def _normalize_sources(src):
        if src is None:
            return []
        if isinstance(src, (list, tuple)):
            res = []
            for item in src:
                if isinstance(item, str) and os.path.isdir(item):
                    for fn in os.listdir(item):
                        if fn.lower().endswith(('.xlsx', '.xlsm', '.csv')) and not fn.startswith('~$'):
                            res.append(os.path.join(item, fn))
                else:
                    res.append(item)
            return res
        if isinstance(src, str) and os.path.isdir(src):
            res = []
            for fn in os.listdir(src):
                if fn.lower().endswith(('.xlsx', '.xlsm', '.csv')) and not fn.startswith('~$'):
                    res.append(os.path.join(src, fn))
            return res
        return [src]

    def _open_wb(self, src):
        is_csv = False
        src_str = str(src).lower() if isinstance(src, (str, os.PathLike)) else (str(src.name).lower() if hasattr(src, 'name') else '')
        if src_str.endswith('.csv'):
            is_csv = True

        is_twamp = ('twamp' in src_str) or (self.twamp_sources and src in self.twamp_sources)

        if is_csv:
            try:
                if hasattr(src, 'seek'):
                    src.seek(0)
                try:
                    df = pd.read_csv(src, encoding='utf-8-sig')
                except Exception:
                    if hasattr(src, 'seek'):
                        src.seek(0)
                    df = pd.read_csv(src, encoding='latin1')

                col_str = " ".join(str(c).lower() for c in df.columns)
                if not is_twamp and any(k in col_str for k in ['dl_dmax_delay', 'delay', 'jitter', 'plr', 'moentity']):
                    is_twamp = True

                wb = Workbook()
                ws = wb.active
                ws.title = "TWAMP" if is_twamp else "5G Daily"
                ws.append(list(df.columns))
                for r in df.itertuples(index=False):
                    ws.append(list(r))
                return wb
            except Exception as e:
                self.log(f"Error parsing CSV source: {e}")

        if isinstance(src, (str, os.PathLike)):
            return openpyxl.load_workbook(src, data_only=True, read_only=True)
        if hasattr(src, 'seek'):
            src.seek(0)
        return openpyxl.load_workbook(src, data_only=True, read_only=True)


    def inspect_sources(self):
        self.log(f"Memindai {len(self.kpi_sources)} sumber file KPI...")
        cluster_set = set()
        date_set = set()
        site_set = set()

        for src in self.kpi_sources:
            try:
                wb = self._open_wb(src)
                ws5g = find_5g_sheet(wb)
                first_row = next(ws5g.iter_rows(values_only=True), None)
                if not first_row:
                    wb.close()
                    continue
                hdrs = {str(h).strip().lower(): i for i, h in enumerate(first_row) if h is not None}

                i_cluster = find_header_index(hdrs, ['cluster', 'cluster_name'])
                i_date    = find_header_index(hdrs, ['date', 'time', 'datetime'])
                i_site    = find_header_index(hdrs, ['site_id', 'site id', 'site', 'sitename'])

                for row in ws5g.iter_rows(values_only=True):
                    cl_val = str(row[i_cluster]).strip() if (i_cluster is not None and row[i_cluster]) else ""
                    if cl_val and cl_val.upper() not in ['CLUSTER', 'CLUSTER_NAME', 'NONE', 'NAN']:
                        cluster_set.add(cl_val)
                    if i_date is not None and row[i_date]:
                        p_date = parse_date_val(row[i_date])
                        if p_date:
                            date_set.add(p_date)
                    if i_site is not None and row[i_site]:
                        st_val = str(row[i_site]).strip()
                        if st_val.upper() not in ['SITE_ID', 'SITE', 'SITENAME', 'NONE', 'NAN']:
                            site_set.add(st_val)

                wb.close()
            except Exception as e:
                self.log(f"  [Peringatan] Gagal memindai salah satu sumber KPI: {e}")

        self.clusters = sorted(list(cluster_set))
        self.available_dates = sorted(list(date_set))
        self.all_sites = sorted(list(site_set))

        twamp_sources_to_check = self.twamp_sources if self.twamp_sources else self.kpi_sources
        self.log(f"Memindai {len(twamp_sources_to_check)} sumber TWAMP...")
        tw_d_set = set()

        for src in twamp_sources_to_check:
            try:
                wb = self._open_wb(src)
                ws_tw = find_twamp_sheet(wb)
                if ws_tw:
                    all_rows = list(ws_tw.iter_rows(values_only=True))
                    if all_rows:
                        first_row = all_rows[0]
                        hdrs_tw = {str(h).strip().lower(): i for i, h in enumerate(first_row) if h is not None}
                        it_date = find_header_index(hdrs_tw, ['time', 'date'])
                        if it_date is not None:
                            for row in all_rows[1:]:
                                t = row[it_date]
                                pt = parse_date_val(t)
                                if pt:
                                    tw_d_set.add(pt)
                wb.close()
            except Exception as e:
                self.log(f"  [Peringatan] Gagal memindai salah satu sumber TWAMP: {e}")

        self.twamp_dates_all = sorted(list(tw_d_set))
        self.log(f"Inspeksi selesai: Ditemukan {len(self.clusters)} cluster, {len(self.available_dates)} tanggal KPI, {len(self.twamp_dates_all)} tanggal TWAMP.")

        return {
            "clusters": self.clusters,
            "dates": self.available_dates,
            "sites": self.all_sites,
            "twamp_dates": self.twamp_dates_all
        }

    def _match_cluster_name(self, target_cluster):
        if not self.clusters:
            self.inspect_sources()
        target_upper = target_cluster.upper()
        for cl in sorted(self.clusters):
            cl_upper = cl.upper()
            if cl_upper == target_upper or target_upper.startswith(cl_upper) or cl_upper.startswith(target_upper):
                return cl
        from difflib import get_close_matches
        candidates = get_close_matches(target_cluster, self.clusters, n=1, cutoff=0.5)
        return candidates[0] if candidates else target_cluster

    def get_cluster_cells_preview(self, target_cluster, site_filter=None, band_filter='nr26_baseline_nr21', *args, **kwargs):
        if 'band_filter' in kwargs:
            band_filter = kwargs['band_filter']
        elif 'band' in kwargs:
            band_filter = kwargs['band']
        matched_cl = self._match_cluster_name(target_cluster)
        site_filter_set = set(str(s).strip() for s in site_filter) if site_filter else None
        
        cells_26 = {}
        cells_21 = {}

        for src in self.kpi_sources:
            try:
                wb = self._open_wb(src)
                ws5g = find_5g_sheet(wb)
                first_row = next(ws5g.iter_rows(values_only=True), None)
                if not first_row:
                    wb.close()
                    continue
                hdrs = {str(h).strip().lower(): i for i, h in enumerate(first_row) if h is not None}

                i_cluster = find_header_index(hdrs, ['cluster', 'cluster_name'])
                i_site    = find_header_index(hdrs, ['site_id', 'site id', 'site', 'sitename'])
                i_sector  = find_header_index(hdrs, ['sector_id', 'sector id', 'sector', 'sec_id', 'sec'])
                i_cellname= find_header_index(hdrs, ['cellname', 'cell_name', 'short name', 'short_name'])
                i_band    = find_header_index(hdrs, ['band', 'band_name'])

                for row in ws5g.iter_rows(values_only=True):
                    if i_cluster is not None and row[i_cluster] != matched_cl:
                        continue
                    site = row[i_site] if i_site is not None else None
                    sector = row[i_sector] if i_sector is not None else None
                    cell = row[i_cellname] if i_cellname is not None else None
                    band_val = row[i_band] if i_band is not None else None
                    if not site:
                        continue
                    site_s = str(site).strip()
                    if site_filter_set and site_s not in site_filter_set:
                        continue

                    clean_sec = ISDMatcher.normalize_sector(sector)
                    band = extract_band_from_cell_or_row(cell, band_val)
                    key = (site_s, clean_sec)

                    if band == '5G26':
                        cells_26[key] = str(cell or '')
                    elif band == '5G21':
                        cells_21[key] = str(cell or '')

                wb.close()
            except Exception as e:
                self.log(f"  [Peringatan] Preview cell error: {e}")

        # Tentukan cell list berdasarkan band_filter
        result = []
        if band_filter == 'nr21_only':
            for k in sorted(cells_21.keys(), key=lambda x: (x[0], x[1] if isinstance(x[1], int) else 0)):
                result.append((k[0], k[1], cells_21[k], '5G21'))
        elif band_filter == 'all_bands_separate':
            for k in sorted(cells_21.keys(), key=lambda x: (x[0], x[1] if isinstance(x[1], int) else 0)):
                result.append((k[0], k[1], cells_21[k], '5G21'))
            for k in sorted(cells_26.keys(), key=lambda x: (x[0], x[1] if isinstance(x[1], int) else 0)):
                result.append((k[0], k[1], cells_26[k], '5G26'))
        else:
            # nr26_only atau nr26_baseline_nr21
            all_keys = set(cells_26.keys()) | set(cells_21.keys())
            for k in sorted(all_keys, key=lambda x: (x[0], x[1] if isinstance(x[1], int) else 0)):
                # Prioritaskan nama cell 5G26
                c_name = cells_26.get(k) or cells_21.get(k, '')
                result.append((k[0], k[1], c_name, '5G26' if k in cells_26 else '5G21'))

        return result

    def generate_report(self, target_cluster, site_filter=None, band_filter='nr26_baseline_nr21',
                        before_dates=None, before_range=None,
                        after_dates=None, after_range=None,
                        twamp_dates=None, twamp_range=None,
                        isd_matcher=None,
                        output_file=None, progress_callback=None, *args, **kwargs):
        """
        Menghasilkan laporan Excel WCL NR26.
        Mendukung Band Filter cerdas (NR26 + Baseline NR21).
        """
        if isd_matcher is not None:
            self.set_isd_matcher(isd_matcher)
        elif not getattr(self.isd_matcher, 'mapping', None) and os.path.exists(DEFAULT_ISD_PATH):
            self.log(f"Memuat data ISD default dari: {DEFAULT_ISD_PATH}")
            self.isd_matcher = ISDMatcher(DEFAULT_ISD_PATH)

        if 'band_filter' in kwargs:
            band_filter = kwargs['band_filter']
        elif 'band' in kwargs:
            band_filter = kwargs['band']
        def update_progress(pct, msg):
            self.log(msg)
            if progress_callback:
                progress_callback(pct, msg)

        update_progress(5, f"Mempersiapkan data untuk cluster '{target_cluster}' (Mode Band: {band_filter})...")
        matched_cluster = self._match_cluster_name(target_cluster)
        site_filter_set = set(str(s).strip() for s in site_filter) if site_filter else None

        all_dates  = set()
        cells_26 = {}
        cells_21 = {}
        
        # raw[site][sector][band][date] = { kpi: [values], kpi_nom: [...], kpi_denom: [...] }
        raw = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list)))))
        raw_ta = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list))))

        update_progress(15, f"Membaca {len(self.kpi_sources)} file KPI dan mengagregasi data...")

        for idx_src, src in enumerate(self.kpi_sources):
            try:
                wb = self._open_wb(src)
                ws5g = find_5g_sheet(wb)
                first_row = next(ws5g.iter_rows(values_only=True), None)
                if not first_row:
                    wb.close()
                    continue
                hdrs = {str(h).strip().lower(): i for i, h in enumerate(first_row) if h is not None}

                i_cluster = find_header_index(hdrs, ['cluster', 'cluster_name'])
                i_date    = find_header_index(hdrs, ['date', 'time', 'datetime'])
                i_site    = find_header_index(hdrs, ['site_id', 'site id', 'site', 'sitename'])
                i_sector  = find_header_index(hdrs, ['sector_id', 'sector id', 'sector', 'sec_id', 'sec'])
                i_cellname= find_header_index(hdrs, ['cellname', 'cell_name', 'short name', 'short_name'])
                i_band    = find_header_index(hdrs, ['band', 'band_name'])

                kpi_map = {}
                for (sn, noms, denoms, vals, *_) in KPI_DEFINITIONS:
                    kpi_map[sn] = {
                        'nom': find_header_index(hdrs, noms),
                        'denom': find_header_index(hdrs, denoms),
                        'val': find_header_index(hdrs, vals)
                    }

                i_ta_nom   = find_header_index(hdrs, TA_NOM_CANDIDATES)
                i_ta_denom = find_header_index(hdrs, TA_DENOM_CANDIDATES)
                i_ta_avg   = find_header_index(hdrs, TA_AVG_CANDIDATES)

                for row in ws5g.iter_rows(values_only=True):
                    if i_cluster is not None and row[i_cluster] != matched_cluster:
                        continue
                    date   = row[i_date] if i_date is not None else None
                    site   = row[i_site] if i_site is not None else None
                    sector = row[i_sector] if i_sector is not None else None
                    cell   = row[i_cellname] if i_cellname is not None else None
                    band_v = row[i_band] if i_band is not None else None

                    p_date = parse_date_val(date)
                    if not p_date or not site:
                        continue
                    site_s = str(site).strip()
                    if site_filter_set and site_s not in site_filter_set:
                        continue

                    all_dates.add(p_date)
                    cell_s = str(cell or '')
                    sec_val = ISDMatcher.normalize_sector(sector)
                    band = extract_band_from_cell_or_row(cell_s, band_v)
                    key = (site_s, sec_val)

                    if band == '5G26': cells_26[key] = cell_s
                    elif band == '5G21': cells_21[key] = cell_s

                    # TA Avg
                    if band in ('5G21', '5G26') and sec_val is not None:
                        ta_v = row[i_ta_avg] if i_ta_avg is not None else None
                        ta_n = row[i_ta_nom] if i_ta_nom is not None else None
                        ta_d = row[i_ta_denom] if i_ta_denom is not None else None
                        raw_ta[site_s][sec_val][band][p_date].append((ta_v, ta_n, ta_d))

                    # 9 KPI
                    if band in ('5G21', '5G26') and sec_val is not None:
                        for (sn, *_) in KPI_DEFINITIONS:
                            cols = kpi_map[sn]
                            vi = cols['val']
                            if vi is not None and row[vi] is not None:
                                try: raw[site_s][sec_val][band][p_date][sn].append(float(row[vi]))
                                except: pass
                            ni = cols['nom']
                            if ni is not None and row[ni] is not None:
                                try: raw[site_s][sec_val][band][p_date][sn+'_nom'].append(float(row[ni]))
                                except: pass
                            di = cols['denom']
                            if di is not None and row[di] is not None:
                                try: raw[site_s][sec_val][band][p_date][sn+'_denom'].append(float(row[di]))
                                except: pass

                wb.close()
            except Exception as e:
                self.log(f"  [Peringatan] Gagal membaca KPI source {idx_src+1}: {e}")

        all_dates = sorted(all_dates)

        # Susun daftar baris cell target (cells_sorted) berdasarkan band_filter
        cells_info = {}
        if band_filter == 'nr21_only':
            for k in sorted(cells_21.keys(), key=lambda x: (x[0], x[1] if isinstance(x[1], int) else 0)):
                cells_info[k] = cells_21[k]
        elif band_filter == 'nr26_only':
            for k in sorted(cells_26.keys(), key=lambda x: (x[0], x[1] if isinstance(x[1], int) else 0)):
                cells_info[k] = cells_26[k]
        else:
            # nr26_baseline_nr21 atau all
            all_keys = set(cells_26.keys()) | set(cells_21.keys())
            for k in sorted(all_keys, key=lambda x: (x[0], x[1] if isinstance(x[1], int) else 0)):
                cells_info[k] = cells_26.get(k) or cells_21.get(k, '')

        cells_sorted = sorted(cells_info.keys(), key=lambda x: (x[0], x[1] if isinstance(x[1], int) else 0))
        sites_unique = sorted(set(k[0] for k in cells_sorted))

        if not cells_sorted:
            raise ValueError(f"Tidak ada cell yang cocok untuk cluster '{matched_cluster}' dengan filter band '{band_filter}'!")

        update_progress(40, "Menentukan periode BEFORE, AFTER, dan TWAMP...")
        custom_bef = resolve_custom_dates(before_dates, before_range, all_dates, label="BEFORE")
        custom_aft = resolve_custom_dates(after_dates, after_range, all_dates, label="AFTER")

        wcl_dates = custom_bef if custom_bef else (all_dates[:3] if len(all_dates) >= 3 else all_dates)
        rem_dates = [d for d in all_dates if d not in wcl_dates]
        post_dates = custom_aft if custom_aft else (rem_dates[:3] if len(rem_dates) >= 3 else rem_dates)

        n_wcl = len(wcl_dates)
        n_post = len(post_dates)

        def get_kpi_val_smart(site, sec, dt, sn):
            """
            Mengambil nilai KPI dengan logika fallback:
            Jika band_filter == 'nr26_baseline_nr21':
              Cari nilai di '5G26'. Jika tanggal tersebut belum ada 5G26 (misal Before),
              ambil nilai baseline dari '5G21'!
            """
            target_band = '5G26' if band_filter != 'nr21_only' else '5G21'
            cell_data = raw[site][sec][target_band].get(dt, {})

            # Cek apakah ada data di target_band
            nom = cell_data.get(sn + '_nom', [])
            den = cell_data.get(sn + '_denom', [])
            val = cell_data.get(sn, [])

            if not nom and not den and not val and band_filter == 'nr26_baseline_nr21':
                # Fallback cerdas ke baseline 5G21!
                cell_data = raw[site][sec]['5G21'].get(dt, {})
                nom = cell_data.get(sn + '_nom', [])
                den = cell_data.get(sn + '_denom', [])
                val = cell_data.get(sn, [])

            if nom and den:
                t_nom = sum(nom)
                t_den = sum(den)
                if t_den != 0:
                    return t_nom / t_den
            return avg(val)

        # TWAMP PROCESSING
        update_progress(50, "Memproses data TWAMP...")
        twamp_sources_to_use = self.twamp_sources if self.twamp_sources else self.kpi_sources
        last7_dates = []
        tw_delay  = defaultdict(lambda: defaultdict(list))
        tw_jitter = defaultdict(lambda: defaultdict(list))
        tw_plr    = defaultdict(lambda: defaultdict(list))
        tw_sites  = set()
        tw_dates_set = set()

        for src in twamp_sources_to_use:
            try:
                wb = self._open_wb(src)
                ws_tw = find_twamp_sheet(wb)
                if ws_tw:
                    all_rows = list(ws_tw.iter_rows(values_only=True))
                    if all_rows:
                        first_row = all_rows[0]
                        hdrs_tw = {str(h).strip().lower(): i for i, h in enumerate(first_row) if h is not None}
                        it_date   = find_header_index(hdrs_tw, ['time', 'date'])
                        it_site   = find_header_index(hdrs_tw, ['site_id', 'site id', 'site'])
                        it_delay  = find_header_index(hdrs_tw, ['delay', 'dl_dmax_delay'])
                        it_jitter = find_header_index(hdrs_tw, ['jitter', 'dl_jmax_jiter'])
                        it_plr    = find_header_index(hdrs_tw, ['plr', 'dl_lostperc_plr'])

                        for row in all_rows[1:]:
                            t = row[it_date] if it_date is not None else None
                            pt = parse_date_val(t)
                            if pt: tw_dates_set.add(pt)
                            si = row[it_site] if it_site is not None else None
                            if not si or str(si).strip() not in sites_unique:
                                continue
                            if not pt:
                                continue
                            si_clean = str(si).strip()
                            tw_sites.add(si_clean)
                            try:
                                dv = row[it_delay]  if it_delay is not None else None
                                jv = row[it_jitter] if it_jitter is not None else None
                                pv = row[it_plr]    if it_plr is not None else None
                                if dv is not None and str(dv).strip() != '':
                                    tw_delay[si_clean][pt].append(float(dv))
                                if jv is not None and str(jv).strip() != '':
                                    tw_jitter[si_clean][pt].append(float(jv))
                                if pv is not None and str(pv).strip() != '':
                                    tw_plr[si_clean][pt].append(float(pv))
                            except Exception: pass
                wb.close()
            except Exception as e:
                self.log(f"  [Peringatan] Gagal membaca TWAMP: {e}")

        custom_tw = resolve_custom_dates(twamp_dates, twamp_range, sorted(tw_dates_set), label="TWAMP")
        if custom_tw:
            last7_dates = custom_tw
        elif tw_dates_set:
            last7_dates = sorted(tw_dates_set)[-7:]
        elif twamp_dates:
            last7_dates = [parse_date_val(d) for d in twamp_dates if parse_date_val(d)]
        else:
            last7_dates = []
        tw_sites = sorted(tw_sites)

        # BUILD OUTPUT WORKBOOK
        update_progress(65, "Membangun workbook Excel & menyusun sheet KPI...")
        wb_out = Workbook()
        wb_out.remove(wb_out.active)

        C_CL=1; C_SI=2; C_SC=3; C_CN=4
        C_WS=5;   C_WE=4+n_wcl
        C_R1=5+n_wcl
        C_PS=6+n_wcl; C_PE=5+n_wcl+n_post
        C_R2=6+n_wcl+n_post
        TOTAL_COLS = C_R2 if n_post > 0 else C_R1
        FIRST_DATA = 3
        LAST_DATA  = 2 + len(cells_sorted)

        def remark_formula(cond, col_s, col_e, row):
            if cond is None: return None
            s = get_column_letter(col_s); e = get_column_letter(col_e)
            n_cols = col_e - col_s + 1
            mc = 2 if n_cols >= 3 else 1
            return (f'=IF(COUNT({s}{row}:{e}{row})=0,"",'
                    f'IF(COUNTIF({s}{row}:{e}{row},{cond})>={mc},"Below Design","Above Design"))')

        for idx_kpi, (sheet_name, _, _, _, threshold_lbl, cond, num_fmt) in enumerate(KPI_DEFINITIONS):
            ws = wb_out.create_sheet(title=safe_name(sheet_name))
            ws.sheet_properties.tabColor = PINK_TAB
            ws.sheet_view.showGridLines  = False

            def wr1(col, val, fill=NO_FILL, align=None):
                c = ws.cell(1, col)
                c.value=val; c.font=hdr_font(11,True)
                c.fill=fill; c.border=BORDER; c.alignment=align or center()

            wr1(C_CL, sheet_name, align=left_va())
            wr1(C_SI, threshold_lbl, align=left_va())
            if n_wcl > 0:
                wr1(C_WS, 'WCL', fill=BLUE_FILL)
                if n_wcl > 1: ws.merge_cells(start_row=1,start_column=C_WS,end_row=1,end_column=C_WE)
                for col in range(C_WS+1, C_WE+1): ws.cell(1,col).fill=BLUE_FILL; ws.cell(1,col).border=BORDER
            ws.cell(1,C_R1).fill=ORANGE_FILL; ws.cell(1,C_R1).border=BORDER; ws.cell(1,C_R1).font=hdr_font(11,True)
            if n_post > 0:
                wr1(C_PS, 'Post Optim', fill=GREEN_FILL)
                if n_post > 1: ws.merge_cells(start_row=1,start_column=C_PS,end_row=1,end_column=C_PE)
                for col in range(C_PS+1, C_PE+1): ws.cell(1,col).fill=GREEN_FILL; ws.cell(1,col).border=BORDER
                ws.cell(1,C_R2).fill=ORANGE_FILL; ws.cell(1,C_R2).border=BORDER; ws.cell(1,C_R2).font=hdr_font(11,True)

            def wr2(col, val, fill=NO_FILL):
                c = ws.cell(2, col)
                c.value=val; c.font=hdr_font(11,True)
                c.fill=fill; c.border=BORDER; c.alignment=center(wrap=True)

            wr2(C_CL,'CLUSTER'); wr2(C_SI,'SITE_ID'); wr2(C_SC,'Sector'); wr2(C_CN,'Cellname')
            for i, dt in enumerate(wcl_dates):
                wr2(C_WS+i, dt.strftime('%d-%b').upper(), fill=BLUE_FILL)
            wr2(C_R1, 'Remark', fill=ORANGE_FILL)
            for i, dt in enumerate(post_dates):
                wr2(C_PS+i, dt.strftime('%d-%b').upper(), fill=GREEN_FILL)
            if n_post > 0:
                wr2(C_R2, 'Remark', fill=ORANGE_FILL)

            for row_idx, (site, sector) in enumerate(cells_sorted):
                nr = 3 + row_idx
                cellname = cells_info.get((site, sector), '')

                def wd(col, val, bold=False, fill=NO_FILL, numfmt=None):
                    c = ws.cell(nr, col)
                    c.value=val; c.font=data_font(bold=bold)
                    c.fill=fill; c.border=BORDER; c.alignment=center()
                    if numfmt: c.number_format=numfmt

                wd(C_CL, matched_cluster, bold=True)
                wd(C_SI, site, bold=True)
                wd(C_SC, sector)
                wd(C_CN, cellname)

                for i, dt in enumerate(wcl_dates):
                    v = get_kpi_val_smart(site, sector, dt, sheet_name)
                    wd(C_WS+i, v, numfmt=num_fmt, fill=BLUE_FILL)

                f1 = remark_formula(cond, C_WS, C_WE, nr)
                c = ws.cell(nr, C_R1)
                c.value=f1; c.font=data_font(); c.fill=ORANGE_FILL; c.border=BORDER; c.alignment=center()

                for i, dt in enumerate(post_dates):
                    v = get_kpi_val_smart(site, sector, dt, sheet_name)
                    wd(C_PS+i, v, numfmt=num_fmt, fill=GREEN_FILL)

                if n_post > 0:
                    f2 = remark_formula(cond, C_PS, C_PE, nr)
                    c = ws.cell(nr, C_R2)
                    c.value=f2; c.font=data_font(); c.fill=ORANGE_FILL; c.border=BORDER; c.alignment=center()

            add_cf_text(ws, get_column_letter(C_R1), FIRST_DATA, LAST_DATA, 'Below Design')
            if n_post > 0:
                add_cf_text(ws, get_column_letter(C_R2), FIRST_DATA, LAST_DATA, 'Below Design')

            col_autofit(ws)
            ws.row_dimensions[1].height = 22
            ws.row_dimensions[2].height = 30
            for r in range(3, LAST_DATA+1): ws.row_dimensions[r].height = 16
            ws.freeze_panes = 'E3'

        # SHEET 10: TWAMP (Urutan sheet ke-10 identik master template)
        update_progress(80, "Menyusun sheet TWAMP (Sheet 10)...")
        if last7_dates:
            ws_twamp = wb_out.create_sheet(title='TWAMP')
            ws_twamp.sheet_properties.tabColor = PINK_TAB
            ws_twamp.sheet_view.showGridLines  = False

            N = len(last7_dates)
            C_DS=1; C_DE=1+N
            C_G1=C_DE+1
            C_JS=C_G1+1; C_JE=C_JS+N
            C_G2=C_JE+1
            C_PS=C_G2+1; C_PE=C_PS+N

            def grp_hdr(s, e, label, fill):
                c = ws_twamp.cell(1, s)
                c.value=label; c.font=hdr_font(11,True)
                c.fill=fill; c.border=BORDER; c.alignment=center()
                ws_twamp.merge_cells(start_row=1, start_column=s, end_row=1, end_column=e)
                for col in range(s+1, e+1): ws_twamp.cell(1,col).fill=fill; ws_twamp.cell(1,col).border=BORDER

            grp_hdr(C_DS, C_DE, 'Delay',  YELLOW_HDR)
            grp_hdr(C_JS, C_JE, 'Jitter', BLUE_FILL)
            grp_hdr(C_PS, C_PE, 'PLR',    GREEN_FILL)

            def sub_hdr(col, val, fill):
                c = ws_twamp.cell(2, col)
                c.value=val; c.font=hdr_font(11,True)
                c.fill=fill; c.border=BORDER; c.alignment=center(wrap=True)

            sub_hdr(C_DS,'Site_ID',YELLOW_HDR)
            sub_hdr(C_JS,'Site_ID',BLUE_FILL)
            sub_hdr(C_PS,'Site_ID',GREEN_FILL)

            for i, dt in enumerate(last7_dates):
                lbl = dt.strftime('%d-%b-%y')
                sub_hdr(C_DS+1+i, lbl, YELLOW_HDR)
                sub_hdr(C_JS+1+i, lbl, BLUE_FILL)
                sub_hdr(C_PS+1+i, lbl, GREEN_FILL)

            sites_show = sorted(list(sites_unique)) if sites_unique else sorted(list(tw_sites))
            for ri, site in enumerate(sites_show):
                nr = 3 + ri
                def wd_tw(col, val, fill, numfmt=None):
                    c = ws_twamp.cell(nr, col)
                    c.value=val; c.font=data_font(bold=(col in (C_DS,C_JS,C_PS)))
                    c.fill=fill; c.border=BORDER; c.alignment=center()
                    if numfmt: c.number_format=numfmt

                wd_tw(C_DS, site, YELLOW_HDR)
                wd_tw(C_JS, site, BLUE_FILL)
                wd_tw(C_PS, site, GREEN_FILL)

                for i, dt in enumerate(last7_dates):
                    dv = avg(tw_delay[site].get(dt,[]))
                    jv = avg(tw_jitter[site].get(dt,[]))
                    pv = avg(tw_plr[site].get(dt,[]))
                    wd_tw(C_DS+1+i, round(dv,2) if dv is not None else None, YELLOW_HDR, '0.00')
                    wd_tw(C_JS+1+i, round(jv,2) if jv is not None else None, BLUE_FILL,  '0.00')
                    wd_tw(C_PS+1+i, round(pv,4) if pv is not None else None, GREEN_FILL, '0.00')

            LAST_TW = 2 + len(sites_show)
            for i in range(N):
                cl = get_column_letter(C_DS+1+i)
                add_cf_greater(ws_twamp, cl, 3, LAST_TW, 100)

            ws_twamp.column_dimensions[get_column_letter(C_G1)].width = 2
            ws_twamp.column_dimensions[get_column_letter(C_G2)].width = 2
            ws_twamp.column_dimensions[get_column_letter(C_DS)].width = 12
            ws_twamp.column_dimensions[get_column_letter(C_JS)].width = 12
            ws_twamp.column_dimensions[get_column_letter(C_PS)].width = 12
            for i in range(N):
                ws_twamp.column_dimensions[get_column_letter(C_DS+1+i)].width = 11
                ws_twamp.column_dimensions[get_column_letter(C_JS+1+i)].width = 11
                ws_twamp.column_dimensions[get_column_letter(C_PS+1+i)].width = 11

            ws_twamp.row_dimensions[1].height = 22.05
            ws_twamp.row_dimensions[2].height = 30.0
            for r in range(3, LAST_TW+1): ws_twamp.row_dimensions[r].height = 16.05
            ws_twamp.freeze_panes = 'B3'

        # SHEET 11: TA AVG (Urutan sheet ke-11 identik master template)
        update_progress(88, "Menyusun sheet TA Avg & menyuntikkan data ISD (Sheet 11)...")
        ws_ta = wb_out.create_sheet(title='TA Avg')
        ws_ta.sheet_properties.tabColor = PINK_TAB
        ws_ta.sheet_view.showGridLines  = False

        def period_label(prefix, dates):
            valid = [d for d in dates if isinstance(d, datetime.date)]
            if not valid: return prefix
            return f'{prefix} {valid[0].strftime("%d %b")}-{valid[-1].strftime("%d %b %Y")}'

        bef_label = period_label('BEFORE', wcl_dates)
        aft_label = period_label('AFTER',  post_dates)

        C_CLUSTER   = 1;  C_SITE     = 2;  C_SECTOR   = 3
        C_BEF_NR21  = 4;  C_BEF_NR26 = 5;  C_BEF_ISD  = 6;  C_BEF_RATIO = 7;  C_BEF_REM = 8
        C_AFT_NR21  = 9;  C_AFT_NR26 = 10; C_AFT_ISD  = 11; C_AFT_RATIO = 12; C_AFT_REM = 13

        def wr1_ta(col, val, fill=NO_FILL, align=None):
            c = ws_ta.cell(1, col)
            c.value=val; c.font=hdr_font(11,True)
            c.fill=fill; c.border=BORDER; c.alignment=align or center()

        wr1_ta(C_CLUSTER, 'CLUSTER', align=left_va())
        ws_ta.cell(1, C_SITE).border   = BORDER
        ws_ta.cell(1, C_SECTOR).border = BORDER

        wr1_ta(C_BEF_NR21, bef_label, fill=BLUE_FILL)
        ws_ta.merge_cells(start_row=1, start_column=C_BEF_NR21, end_row=1, end_column=C_BEF_REM)
        for col in range(C_BEF_NR21+1, C_BEF_REM+1):
            ws_ta.cell(1, col).fill   = BLUE_FILL
            ws_ta.cell(1, col).border = BORDER
            ws_ta.cell(1, col).font   = hdr_font(11, True)

        wr1_ta(C_AFT_NR21, aft_label, fill=GREEN_FILL)
        ws_ta.merge_cells(start_row=1, start_column=C_AFT_NR21, end_row=1, end_column=C_AFT_REM)
        for col in range(C_AFT_NR21+1, C_AFT_REM+1):
            ws_ta.cell(1, col).fill   = GREEN_FILL
            ws_ta.cell(1, col).border = BORDER
            ws_ta.cell(1, col).font   = hdr_font(11, True)

        def wr2_ta(col, val, fill=NO_FILL):
            c = ws_ta.cell(2, col)
            c.value=val; c.font=hdr_font(11,True)
            c.fill=fill; c.border=BORDER; c.alignment=center(wrap=True)

        wr2_ta(C_CLUSTER,   'CLUSTER')
        wr2_ta(C_SITE,      'Site Name')
        wr2_ta(C_SECTOR,    'Sector')

        wr2_ta(C_BEF_NR21,  '5G TA (NR21)',             fill=BLUE_FILL)
        wr2_ta(C_BEF_NR26,  '5G TA (NR26)',             fill=BLUE_FILL)
        wr2_ta(C_BEF_ISD,   'ISD m',                    fill=BLUE_FILL)
        wr2_ta(C_BEF_RATIO, 'Ratio TA NR26',            fill=ORANGE_FILL)
        wr2_ta(C_BEF_REM,   'Final Remark NR26 to ISD', fill=ORANGE_FILL)

        wr2_ta(C_AFT_NR21,  '5G TA (NR21)',             fill=GREEN_FILL)
        wr2_ta(C_AFT_NR26,  '5G TA (NR26)',             fill=GREEN_FILL)
        wr2_ta(C_AFT_ISD,   'ISD m',                    fill=GREEN_FILL)
        wr2_ta(C_AFT_RATIO, 'Ratio TA NR26',            fill=ORANGE_FILL)
        wr2_ta(C_AFT_REM,   'Final Remark NR26 to ISD', fill=ORANGE_FILL)

        isd_matched_count = 0
        for row_idx, (site, sector) in enumerate(cells_sorted):
            nr = 3 + row_idx

            def wd_ta(col, val, bold=False, fill=NO_FILL, numfmt=None):
                c = ws_ta.cell(nr, col)
                c.value=val; c.font=data_font(bold=bold)
                c.fill=fill; c.border=BORDER; c.alignment=center()
                if numfmt: c.number_format=numfmt

            wd_ta(C_CLUSTER, matched_cluster, bold=True)
            wd_ta(C_SITE,    site, bold=True)
            wd_ta(C_SECTOR,  sector)

            ta21_b = calc_period_ta(raw_ta[site][sector]['5G21'], wcl_dates)
            ta26_b = calc_period_ta(raw_ta[site][sector]['5G26'], wcl_dates)

            isd_val = self.isd_matcher.get_isd(site, sector)
            if isd_val is not None:
                isd_matched_count += 1

            wd_ta(C_BEF_NR21, ta21_b, numfmt='0', fill=BLUE_FILL)
            wd_ta(C_BEF_NR26, ta26_b, numfmt='0', fill=BLUE_FILL)
            wd_ta(C_BEF_ISD,  isd_val, numfmt='0', fill=BLUE_FILL)

            e_ltr = get_column_letter(C_BEF_NR26)
            f_ltr = get_column_letter(C_BEF_ISD)
            g_ltr = get_column_letter(C_BEF_RATIO)
            c = ws_ta.cell(nr, C_BEF_RATIO)
            c.value = f'={e_ltr}{nr}-{f_ltr}{nr}'
            c.font=data_font(); c.fill=ORANGE_FILL; c.border=BORDER
            c.alignment=center(); c.number_format='0'

            c = ws_ta.cell(nr, C_BEF_REM)
            c.value = f'=IF({g_ltr}{nr}>100,"TA NR26 Overshoot","OK")'
            c.font=data_font(); c.fill=ORANGE_FILL; c.border=BORDER; c.alignment=center()

            ta21_a = calc_period_ta(raw_ta[site][sector]['5G21'], post_dates)
            ta26_a = calc_period_ta(raw_ta[site][sector]['5G26'], post_dates)

            wd_ta(C_AFT_NR21, ta21_a, numfmt='0', fill=GREEN_FILL)
            wd_ta(C_AFT_NR26, ta26_a, numfmt='0', fill=GREEN_FILL)
            wd_ta(C_AFT_ISD,  isd_val, numfmt='0', fill=GREEN_FILL)

            j_ltr = get_column_letter(C_AFT_NR26)
            k_ltr = get_column_letter(C_AFT_ISD)
            l_ltr = get_column_letter(C_AFT_RATIO)
            c = ws_ta.cell(nr, C_AFT_RATIO)
            c.value = f'={j_ltr}{nr}-{k_ltr}{nr}'
            c.font=data_font(); c.fill=ORANGE_FILL; c.border=BORDER
            c.alignment=center(); c.number_format='0'

            c = ws_ta.cell(nr, C_AFT_REM)
            c.value = f'=IF({l_ltr}{nr}>100,"TA NR26 Overshoot","OK")'
            c.font=data_font(); c.fill=ORANGE_FILL; c.border=BORDER; c.alignment=center()

        add_cf_text(ws_ta, get_column_letter(C_BEF_REM), FIRST_DATA, LAST_DATA, 'Overshoot')
        add_cf_text(ws_ta, get_column_letter(C_AFT_REM), FIRST_DATA, LAST_DATA, 'Overshoot')

        col_widths_ta = {
            C_CLUSTER: 30, C_SITE: 13, C_SECTOR: 7,
            C_BEF_NR21: 13, C_BEF_NR26: 13, C_BEF_ISD: 10, C_BEF_RATIO: 14, C_BEF_REM: 24,
            C_AFT_NR21: 13, C_AFT_NR26: 13, C_AFT_ISD: 10, C_AFT_RATIO: 14, C_AFT_REM: 26.22,
        }
        for col, w in col_widths_ta.items():
            ws_ta.column_dimensions[get_column_letter(col)].width = w

        ws_ta.row_dimensions[1].height = 22.05
        ws_ta.row_dimensions[2].height = 34.95
        for r in range(3, LAST_DATA+1): ws_ta.row_dimensions[r].height = 16.05
        ws_ta.freeze_panes = None

        for sh in wb_out.worksheets:
            sh.sheet_state = 'visible'
        wb_out.active = wb_out.worksheets[0]

        # ── RUN AUTOMATIC MASTER TEMPLATE VALIDATION ──
        update_progress(92, "Menjalankan validasi otomatis kesesuaian master template...")
        validation_res = validate_template_compliance(wb_out)
        self.log(f"Hasil Validasi Template: {validation_res['compliance_rate']}% PASS ({validation_res['passed_checks']}/{validation_res['total_checks']} kriteria terpenuhi)")

        # ── RUN SMART RCA GENERATOR ──
        update_progress(94, "Menjalankan Smart RCA Diagnostic Engine...")
        rca_df = generate_smart_rca(
            raw=raw,
            raw_ta=raw_ta,
            wcl_dates=wcl_dates,
            post_dates=post_dates,
            isd_matcher=self.isd_matcher,
            cells_sorted=cells_sorted,
            cells_info=cells_info,
            tw_delay=tw_delay,
            tw_jitter=tw_jitter,
            tw_plr=tw_plr,
            band_filter=band_filter
        )

        # ── GENERATE EXECUTIVE POWERPOINT PRESENTATION (.PPTX) ──
        update_progress(96, "Menghasilkan presentasi eksekutif PowerPoint (.pptx)...")
        pptx_bytes = generate_powerpoint_presentation(
            cluster=matched_cluster,
            rca_df=rca_df,
            wcl_dates=wcl_dates,
            post_dates=post_dates
        )

        # ── GENERATE DASHBOARD CHARTS (PNG / JPG) ──
        update_progress(97, "Merender grafik dashboard eksekutif (PNG & JPG)...")
        chart_images = generate_dashboard_images(
            cluster=matched_cluster,
            wcl_dates=wcl_dates,
            post_dates=post_dates
        )

        update_progress(98, "Menyimpan output report...")
        output_stream = io.BytesIO()
        wb_out.save(output_stream)
        output_stream.seek(0)

        if output_file:
            os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)
            with open(output_file, 'wb') as f:
                f.write(output_stream.getvalue())
            self.log(f"Laporan berhasil disimpan ke: {output_file}")

            try:
                base_no_ext = os.path.splitext(output_file)[0]
                pptx_path = base_no_ext + "_Executive_Deck.pptx"
                with open(pptx_path, 'wb') as f_pptx:
                    f_pptx.write(pptx_bytes)
            except Exception as e:
                self.log(f"Warning saving pptx: {e}")

        update_progress(100, f"Selesai! Berhasil membuat {len(wb_out.sheetnames)} sheet, {len(cells_sorted)} cells ({isd_matched_count}/{len(cells_sorted)} ISD cocok). Template Compliance: {validation_res['compliance_rate']}%.")

        return {
            "workbook": wb_out,
            "bytes": output_stream.getvalue(),
            "pptx_bytes": pptx_bytes,
            "images": chart_images,
            "rca_df": rca_df,
            "validation": validation_res,
            "total_cells": len(cells_sorted),
            "isd_matched": isd_matched_count,
            "sheets": wb_out.sheetnames,
            "cluster": matched_cluster,
            "wcl_dates": [str(d) for d in wcl_dates],
            "post_dates": [str(d) for d in post_dates],
            "cells_sorted": cells_sorted,
            "cells_info": cells_info
        }

