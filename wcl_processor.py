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
from collections import defaultdict
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl.utils import get_column_letter
import pandas as pd


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
            for candidate in ['isd', 'isd (m)', 'isd m', 'isd_m', 'isd(m)', 'distance', 'distance_m', 'jarak']:
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
            if pd_d and pd_d in avail_set:
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
    for name in ['twamp', 'Twamp', 'TWAMP']:
        if name in wb.sheetnames:
            return wb[name]
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
        else:
            self.isd_matcher = ISDMatcher()

        self.clusters = []
        self.available_dates = []
        self.all_sites = []
        self.twamp_dates_all = []

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
                        if fn.endswith(('.xlsx', '.xlsm')) and not fn.startswith('~$'):
                            res.append(os.path.join(item, fn))
                else:
                    res.append(item)
            return res
        if isinstance(src, str) and os.path.isdir(src):
            res = []
            for fn in os.listdir(src):
                if fn.endswith(('.xlsx', '.xlsm')) and not fn.startswith('~$'):
                    res.append(os.path.join(src, fn))
            return res
        return [src]

    def _open_wb(self, src):
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
                    first_row = next(ws_tw.iter_rows(values_only=True), None)
                    if first_row:
                        hdrs_tw = {str(h).strip().lower(): i for i, h in enumerate(first_row) if h is not None}
                        it_date = find_header_index(hdrs_tw, ['time', 'date'])
                        if it_date is not None:
                            for row in ws_tw.iter_rows(values_only=True):
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

    def get_cluster_cells_preview(self, target_cluster, site_filter=None, band_filter='nr26_baseline_nr21'):
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
                        output_file=None, progress_callback=None):
        """
        Menghasilkan laporan Excel WCL NR26.
        Mendukung Band Filter cerdas (NR26 + Baseline NR21).
        """
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
                    first_row = next(ws_tw.iter_rows(values_only=True), None)
                    if first_row:
                        hdrs_tw = {str(h).strip().lower(): i for i, h in enumerate(first_row) if h is not None}
                        it_date   = find_header_index(hdrs_tw, ['time', 'date'])
                        it_site   = find_header_index(hdrs_tw, ['site_id', 'site id', 'site'])
                        it_delay  = find_header_index(hdrs_tw, ['delay', 'dl_dmax_delay'])
                        it_jitter = find_header_index(hdrs_tw, ['jitter', 'dl_jmax_jiter'])
                        it_plr    = find_header_index(hdrs_tw, ['plr', 'dl_lostperc_plr'])

                        for row in ws_tw.iter_rows(values_only=True):
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
                                tw_delay[si_clean][pt].append(float(dv)) if dv is not None else None
                                tw_jitter[si_clean][pt].append(float(jv)) if jv is not None else None
                                tw_plr[si_clean][pt].append(float(pv)) if pv is not None else None
                            except: pass
                wb.close()
            except Exception as e:
                self.log(f"  [Peringatan] Gagal membaca TWAMP: {e}")

        custom_tw = resolve_custom_dates(twamp_dates, twamp_range, sorted(tw_dates_set), label="TWAMP")
        last7_dates = custom_tw if custom_tw else sorted(tw_dates_set)[-7:]
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

        # SHEET TA AVG
        update_progress(80, "Menyusun sheet TA Avg & menyuntikkan data ISD (Site ID + Sector)...")
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

        add_cf_text(ws_ta, get_column_letter(C_BEF_REM), FIRST_DATA, LAST_DATA, 'TA NR26 Overshoot')
        add_cf_text(ws_ta, get_column_letter(C_AFT_REM), FIRST_DATA, LAST_DATA, 'TA NR26 Overshoot')

        col_widths_ta = {
            C_CLUSTER: 32, C_SITE: 14, C_SECTOR: 8,
            C_BEF_NR21: 13, C_BEF_NR26: 13, C_BEF_ISD: 10, C_BEF_RATIO: 14, C_BEF_REM: 25,
            C_AFT_NR21: 13, C_AFT_NR26: 13, C_AFT_ISD: 10, C_AFT_RATIO: 14, C_AFT_REM: 25,
        }
        for col, w in col_widths_ta.items():
            ws_ta.column_dimensions[get_column_letter(col)].width = w

        ws_ta.row_dimensions[1].height = 22
        ws_ta.row_dimensions[2].height = 35
        for r in range(3, LAST_DATA+1): ws_ta.row_dimensions[r].height = 16
        ws_ta.freeze_panes = 'D3'

        # SHEET TWAMP
        update_progress(90, "Menyusun sheet TWAMP...")
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

            sites_show = tw_sites if tw_sites else sites_unique
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
                    wd_tw(C_PS+1+i, round(pv,4) if pv is not None else None, GREEN_FILL, '0.0000')

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

            ws_twamp.row_dimensions[1].height = 22
            ws_twamp.row_dimensions[2].height = 30
            for r in range(3, LAST_TW+1): ws_twamp.row_dimensions[r].height = 16
            ws_twamp.freeze_panes = 'B3'

        for sh in wb_out.worksheets:
            sh.sheet_state = 'visible'
        wb_out.active = wb_out.worksheets[0]

        update_progress(95, "Menyimpan output report...")
        output_stream = io.BytesIO()
        wb_out.save(output_stream)
        output_stream.seek(0)

        if output_file:
            os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)
            with open(output_file, 'wb') as f:
                f.write(output_stream.getvalue())
            self.log(f"Laporan berhasil disimpan ke: {output_file}")

        update_progress(100, f"Selesai! Berhasil membuat {len(wb_out.sheetnames)} sheet, {len(cells_sorted)} cells ({isd_matched_count}/{len(cells_sorted)} ISD cocok).")

        return {
            "workbook": wb_out,
            "bytes": output_stream.getvalue(),
            "total_cells": len(cells_sorted),
            "isd_matched": isd_matched_count,
            "sheets": wb_out.sheetnames,
            "cluster": matched_cluster,
            "wcl_dates": [str(d) for d in wcl_dates],
            "post_dates": [str(d) for d in post_dates]
        }
