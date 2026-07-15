#!/usr/bin/env python3
"""
BikeDekho Dashboard Auto-Rebuild Pipeline
------------------------------------------
Detects new/changed LD and LMS Excel files in the workspace,
then rebuilds the affected OEM HTML dashboards in-place.

Stored in the workspace folder so it persists across sessions.
Run as: python3 /path/to/auto_rebuild.py
"""

import os, sys, json, re, hashlib, warnings, traceback
from datetime import datetime
from pathlib import Path

warnings.filterwarnings('ignore')

# ── Locate workspace (the folder this script lives in) ──────────────────────
WORKSPACE          = Path(os.path.dirname(os.path.abspath(__file__)))
LD_DIR             = WORKSPACE / "LD Files"
LMS_DIR            = WORKSPACE / "LMS Files"
GEN_DIR            = WORKSPACE / "Generation Data from BD CRM Table"
LATEST_STATUS_DIR  = WORKSPACE / "LMS Data for Latest Lead Status"
MANIFEST           = WORKSPACE / ".rebuild_manifest.json"

# ── BigQuery mode (replaces Gen Excel files with live BQ data) ───────────────
# Set USE_BQ_FOR_GEN = True to pull Generation + LD counts from BigQuery.
# The existing Excel path is fully preserved — flip back to False to revert.
# NOTE: LD Excel files are still required even in BQ mode (for opty_id→model join).
USE_BQ_FOR_GEN = False

BQ_PROJECT   = "product-dashboard-487908"
BQ_TABLE     = f"{BQ_PROJECT}.kpi_dashboard.bikedekho_monetisation_leads"
BQ_BRAND_MAP = {
    'TVS':    'TVS',
    'Jawa':   'Jawa',
    'Ampere': 'Ampere Electric',
}

print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] BikeDekho Auto-Rebuild starting...")
print(f"  Workspace: {WORKSPACE}")

# ── Dependency check + auto-install fastexcel ────────────────────────────────
try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("ERROR: pandas/numpy not installed. Run: pip install pandas openpyxl")
    sys.exit(1)

try:
    import fastexcel  # noqa
except ImportError:
    print("  Installing fastexcel + pyarrow for fast Excel reading...")
    os.system(f"{sys.executable} -m pip install fastexcel pyarrow --break-system-packages -q")

# ── Lead Type Dictionary ──────────────────────────────────────────────────────
def _load_lt_dict() -> dict:
    """Return {lt_code_int: name_str} mapping from Lead Type Dictionary.xlsx."""
    lt_path = WORKSPACE / "Lead Type Dictionary.xlsx"
    if not lt_path.exists():
        return {}
    try:
        df = pd.read_excel(lt_path)
        # Expect columns: Lead Type, Name
        code_col = next((c for c in df.columns if 'lead type' in str(c).lower()), None)
        name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
        if code_col is None or name_col is None:
            return {}
        return {int(r[code_col]): str(r[name_col])
                for _, r in df.iterrows()
                if pd.notna(r[code_col]) and pd.notna(r[name_col])}
    except Exception:
        return {}

LT_DICT = _load_lt_dict()


# ─────────────────────────────────────────────────────────────────────────────
# MANIFEST HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def file_sig(path: Path) -> str:
    """Return a signature combining mtime and size (fast, no full hash)."""
    try:
        st = path.stat()
        return f"{st.st_mtime:.0f}:{st.st_size}"
    except Exception:
        return "missing"


def scan_files(directory: Path) -> dict:
    """Recursively scan .xlsx and .csv files; return {relative_path: sig}."""
    result = {}
    if not directory.exists():
        return result
    for pattern in ("*.xlsx", "*.csv"):
        for f in directory.rglob(pattern):
            if not f.name.startswith("~"):  # skip temp files
                key = str(f.relative_to(WORKSPACE))
                result[key] = file_sig(f)
    return result


def load_manifest() -> dict:
    if MANIFEST.exists():
        try:
            return json.loads(MANIFEST.read_text())
        except Exception:
            pass
    return {}


def save_manifest(data: dict):
    MANIFEST.write_text(json.dumps(data, indent=2))


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────

LMS_NEEDED_COLS = [
    'EnquiryId', 'CreateTime', 'Current Status',
    'Retail Date', 'Retail Attribution Date', 'Booking Date',
    'TDR Date', 'TDS Date', 'Lost Date', 'PQS Date',
    'Model', 'Model.1', 'Brand',
    'L1 Organization Name', 'Registered Company Name',
    'L1 Dealer City', 'CityName', 'State',
    'Performance Month', 'Retail Purchased Model',
    'Crm Lead Id', 'Web Id', 'Source Lead Id',
]


def _read_xlsx_fast(path: Path, use_cols: list, str_cols: list = None) -> pd.DataFrame:
    """Read xlsx using fastexcel if available, else fall back to openpyxl.

    fastexcel normalises duplicate-column suffixes: 'Model.1' -> 'Model_1'.
    We load ALL columns (no use_columns filter to avoid KeyError), then
    select and rename back so downstream code sees the original names.

    str_cols: columns to force as string to prevent float64 precision loss.
    """
    str_set = set(str_cols or [])
    use_set = set(use_cols)
    try:
        import fastexcel
        wb = fastexcel.read_excel(str(path))
        sheet = wb.load_sheet(0)          # load ALL cols — fast regardless
        df = sheet.to_pandas()
        # Build selection list: prefer original name, fall back to dot->underscore
        keep = []
        rename_back = {}
        for c in use_cols:
            if c in df.columns:
                keep.append(c)
            else:
                c_norm = c.replace('.', '_')
                if c_norm in df.columns:
                    keep.append(c_norm)
                    rename_back[c_norm] = c   # e.g. Model_1 -> Model.1
        df = df[keep].rename(columns=rename_back)
        # Vectorised str_cols fix: convert numeric columns to string
        for c in str_set:
            if c not in df.columns:
                continue
            col = df[c]
            if pd.api.types.is_float_dtype(col):
                df[c] = col.astype('Int64').astype(str).str.replace('<NA>', '', regex=False)
            elif pd.api.types.is_integer_dtype(col):
                df[c] = col.astype(str)
            else:
                df[c] = col.fillna('').astype(str)
        return df
    except ImportError:
        dtype_override = {c: str for c in str_set}
        return pd.read_excel(path, usecols=lambda c: c in use_set,
                             dtype=dtype_override)
    except Exception:
        dtype_override = {c: str for c in str_set}
        return pd.read_excel(path, usecols=lambda c: c in use_set,
                             dtype=dtype_override)


def infer_active_months(oem_folder: Path) -> set:
    """
    Determine which months to include based on LMS file names.
    Files named 'Jan\'26 - File 1.xlsx' → month Jan 2026 is active.
    Picks the most recent calendar year present, all months for that year.
    Returns a set of 'Jan 2026' style strings.
    """
    import re as _re
    month_map = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
                 'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
    found = set()
    for f in oem_folder.glob("*.xlsx"):
        m = _re.match(r"([A-Za-z]{3})'(\d{2})", f.name)
        if m:
            mon, yr = m.group(1), int('20' + m.group(2))
            if m.group(1) in month_map:
                found.add((yr, month_map[m.group(1)], m.group(1)))
    if not found:
        return set()
    # Use the most recent year
    max_yr = max(y for y, _, _ in found)
    active = {f"{mon} {yr}" for yr, _, mon in found if yr == max_yr}
    return active


def load_lms_files(oem_folder: Path) -> pd.DataFrame:
    """
    Load ALL LMS xlsx files for an OEM in parallel using ThreadPoolExecutor.
    De-duplicates by EnquiryId (keep latest row).
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    files = sorted(f for f in oem_folder.glob("*.xlsx") if not f.name.startswith("~"))
    if not files:
        return pd.DataFrame()

    _str_cols = ['Source Lead Id', 'EnquiryId', 'Crm Lead Id']

    def _load(f):
        return f.name, _read_xlsx_fast(f, LMS_NEEDED_COLS, str_cols=_str_cols)

    results = {}
    with ThreadPoolExecutor(max_workers=min(6, len(files))) as pool:
        futures = {pool.submit(_load, f): f for f in files}
        for fut in as_completed(futures):
            try:
                name, df = fut.result()
                results[name] = df
            except Exception as e:
                print(f"  WARNING: Could not read {futures[fut].name}: {e}")

    # Preserve sorted order after parallel load
    loaded = sorted(results.keys())
    dfs = [results[n] for n in loaded]
    print(f"  Loaded {len(loaded)} file(s): {', '.join(loaded)}")
    if not dfs:
        return pd.DataFrame()
    combined = pd.concat(dfs, ignore_index=True)
    # De-duplicate by EnquiryId — keep last (most recently updated status)
    if 'EnquiryId' in combined.columns:
        before = len(combined)
        combined = combined.drop_duplicates(subset=['EnquiryId'], keep='last')
        print(f"  Dedup: {before:,} rows → {len(combined):,} unique leads")
    return combined


def load_ld_files(brand_filter: str) -> pd.DataFrame:
    """Load all LD RawData files and filter to the given brand.

    LD files use actual column names: DATE(t.triggered_date), cd_lead_id,
    opty_id, model, lead_type, utm_source, utm_medium, utm_campaign.
    There is no 'brand' column — brand is inferred from the model name prefix.
    brand_filter values: 'TVS', 'Ampere Electric' (→ 'Ampere'), 'Jawa', 'Yezdi'.
    """
    dfs = []
    # Actual column names in the LD RawData files
    ld_cols = ['DATE(t.triggered_date)', 'cd_lead_id', 'opty_id', 'model',
               'lead_type', 'utm_source', 'utm_medium', 'utm_campaign']

    # Normalise brand_filter to a model-prefix for matching
    _brand_prefix_map = {
        'TVS':            'tvs',
        'Ampere Electric':'ampere',
        'Ampere':         'ampere',
        'Jawa':           'jawa',
        'Yezdi':          'yezdi',
    }
    brand_prefix = _brand_prefix_map.get(brand_filter, brand_filter.lower()) if brand_filter else None

    all_ld_files = sorted(LD_DIR.glob("LD RawData*.xlsx")) + sorted(LD_DIR.glob("LD RawData*.csv"))
    for f in all_ld_files:
        if f.name.startswith("~"):
            continue
        try:
            if f.suffix.lower() == '.csv':
                df = pd.read_csv(f, usecols=lambda c: c in ld_cols, low_memory=False)
            else:
                df = _read_xlsx_fast(f, ld_cols)
            # Rename date column for consistency
            if 'DATE(t.triggered_date)' in df.columns:
                df = df.rename(columns={'DATE(t.triggered_date)': 'Date'})
            if brand_prefix and 'model' in df.columns:
                df = df[df['model'].fillna('').str.lower().str.startswith(brand_prefix)]
            dfs.append(df)
        except Exception as e:
            print(f"  WARNING: Could not read {f.name}: {e}")
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


def load_gen_files(brand_filter: str) -> pd.DataFrame:
    """Load Generation Data from BD CRM Table xlsx files for the given brand.
    Files are named like 'TVS - Gen Dec25.xlsx'. Brand matched from filename.
    Columns: update_date, cd_lead_id, model, lead_type, utm_source, utm_medium,
             utm_campaign, source_bucket.
    """
    if not GEN_DIR.exists():
        return pd.DataFrame()
    brand_prefix = brand_filter.split()[0].lower() if brand_filter else ''
    dfs = []
    all_gen_files = sorted(GEN_DIR.glob("*.xlsx")) + sorted(GEN_DIR.glob("*.csv"))
    for f in all_gen_files:
        if f.name.startswith("~"):
            continue
        if brand_prefix and brand_prefix not in f.name.lower():
            continue
        try:
            if f.suffix.lower() == '.csv':
                df = pd.read_csv(f, low_memory=False)
            else:
                df = pd.read_excel(f)
            if 'update_date' in df.columns:
                df = df.rename(columns={'update_date': 'Date'})
            # Normalise cd_lead_id — strip trailing .0 from float-stored integers
            df['cd_lead_id'] = df['cd_lead_id'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            dfs.append(df)
            print(f"  Gen: loaded {f.name} ({len(df):,} rows)")
        except Exception as e:
            print(f"  WARNING: Could not read Gen file {f.name}: {e}")
    result = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    # Normalise: treat "Others" as "Non-MS"
    if not result.empty and 'source_bucket' in result.columns:
        result['source_bucket'] = result['source_bucket'].replace('Others', 'Non-MS')
    return result


def load_gen_from_bq(oem_name: str) -> pd.DataFrame:
    """Fetch Generation + LD counts from BigQuery (live data, no Excel needed).

    Returns DataFrame with columns: Date (month start), model, generated, triggered.
    Does NOT have cd_lead_id or source_bucket — per-lead enrichment is skipped
    gracefully; supply funnel uses aggregate counts via compute_supply_funnel_bq().
    Falls back to empty DataFrame (caller then uses Excel path) on any failure.
    """
    brand = BQ_BRAND_MAP.get(oem_name)
    if not brand:
        print(f"  BQ Gen: no brand mapping for '{oem_name}' — check BQ_BRAND_MAP")
        return pd.DataFrame()
    try:
        from google.cloud import bigquery
        client = bigquery.Client(project=BQ_PROJECT)
        query = f"""
        SELECT
            DATE_TRUNC(date, MONTH) AS Date,
            COALESCE(model, 'Unknown')  AS model,
            SUM(CASE WHEN status = 'generated' THEN lead_count ELSE 0 END) AS generated,
            SUM(CASE WHEN status = 'triggered' THEN lead_count ELSE 0 END) AS triggered
        FROM `{BQ_TABLE}`
        WHERE date >= DATE_SUB(CURRENT_DATE(), INTERVAL 9 MONTH)
          AND brand = @brand
        GROUP BY 1, 2
        ORDER BY 1
        """
        job_config = bigquery.QueryJobConfig(
            query_parameters=[bigquery.ScalarQueryParameter("brand", "STRING", brand)]
        )
        df = client.query(query, job_config=job_config).to_dataframe()
        df['Date'] = pd.to_datetime(df['Date'])
        total_gen = int(df['generated'].sum())
        total_trig = int(df['triggered'].sum())
        print(f"  BQ Gen: {total_gen:,} generated | {total_trig:,} triggered "
              f"({len(df):,} model-month rows) for {oem_name}")
        return df
    except ImportError:
        print("  BQ Gen: google-cloud-bigquery not installed — falling back to Excel Gen files")
        print("         Install with: pip install google-cloud-bigquery")
        return pd.DataFrame()
    except Exception as e:
        print(f"  BQ Gen: query failed ({e}) — falling back to Excel Gen files")
        return pd.DataFrame()


def compute_supply_funnel_bq(bq_df: pd.DataFrame, lms_df: pd.DataFrame) -> dict:
    """Supply funnel using BQ aggregate counts for Gen+LD, LMS data for bottom funnel.

    Used when USE_BQ_FOR_GEN=True. BQ provides live Gen/LD totals; LMS Excel
    provides LMS entry count and retail/booking/lost outcomes.
    Returns same dict structure as compute_supply_funnel() for dashboard compatibility.
    """
    if bq_df is None or bq_df.empty:
        return {}

    gen_total  = int(bq_df['generated'].sum())
    ld_total   = int(bq_df['triggered'].sum())

    lms_total  = len(lms_df) if not lms_df.empty else 0
    retail_total = int(lms_df['_retail_attr'].sum()) if '_retail_attr' in lms_df.columns and not lms_df.empty else 0

    def safe_pct(num, den): return round(num / den * 100, 2) if den else 0.0

    # Monthly breakdown from BQ gen/triggered + LMS monthly retail
    monthly = []
    if not bq_df.empty and 'Date' in bq_df.columns:
        for dt, grp in bq_df.groupby('Date'):
            mo_label = pd.Timestamp(dt).strftime('%b %Y')
            monthly.append({
                'month':      mo_label,
                'gen':        int(grp['generated'].sum()),
                'ld':         int(grp['triggered'].sum()),
                'lms':        None,   # not available without per-lead join
                'retail':     None,
                'gen_to_ld_pct': safe_pct(int(grp['triggered'].sum()), int(grp['generated'].sum())),
            })

    return {
        'totals': {
            'gen':            gen_total,
            'ld':             ld_total,
            'lms':            lms_total,
            'retail':         retail_total,
            'gen_to_ld_pct':  safe_pct(ld_total,   gen_total),
            'ld_to_lms_pct':  safe_pct(lms_total,  ld_total),
            'retail_rate':    safe_pct(retail_total, lms_total),
        },
        'by_source_bucket': [],   # requires per-lead cd_lead_id — not in BQ aggregate
        'monthly':          monthly,
        'source': 'bigquery',
    }


def load_latest_status_overrides() -> dict:
    """Load all xlsx files from 'LMS Data for Latest Lead Status' folder.

    Returns {opty_id_str: Status_Name_str} — the most recent status for each
    lead, keyed by opty_id. These override 'Current Status' in LMS files.
    Files have columns: Date, brand, model, opty_id, lead_type, utm_campaign,
    oem_crm_id, Medium, Status_Name.
    """
    if not LATEST_STATUS_DIR.exists():
        return {}
    dfs = []
    for f in sorted(LATEST_STATUS_DIR.glob("*.xlsx")):
        if f.name.startswith("~"):
            continue
        try:
            df = pd.read_excel(f, usecols=lambda c: c in ['opty_id', 'Status_Name', 'Date'])
            dfs.append(df)
        except Exception as e:
            print(f"  WARNING: Could not read latest-status file {f.name}: {e}")
    if not dfs:
        return {}
    combined = pd.concat(dfs, ignore_index=True)
    combined = combined.dropna(subset=['opty_id', 'Status_Name'])
    combined['opty_id'] = combined['opty_id'].astype(str).str.strip()
    # If same opty_id appears in multiple files, keep the most recent Date
    if 'Date' in combined.columns:
        combined['Date'] = pd.to_datetime(combined['Date'], errors='coerce')
        combined = combined.sort_values('Date', ascending=True)
    combined = combined.drop_duplicates(subset=['opty_id'], keep='last')
    overrides = dict(zip(combined['opty_id'], combined['Status_Name']))
    print(f"  Latest-status overrides loaded: {len(overrides):,} leads across "
          f"{len(dfs)} file(s)")
    return overrides


def _normalise_lead_id(series: pd.Series) -> pd.Series:
    """Normalise Source Lead Id to string for override matching.

    Source Lead Id is now forced to string at read time (str_cols in
    _read_xlsx_fast), so this is a simple strip. The function is kept
    for defensive safety in case a non-LMS caller passes numeric series.
    """
    if pd.api.types.is_float_dtype(series):
        # Safety net: strip .0 for any float column that bypassed str_cols
        return (series.dropna().astype('Int64').astype(str)
                .reindex(series.index, fill_value=''))
    return series.astype(str).str.strip()


def apply_status_overrides(lms_df: pd.DataFrame, overrides: dict) -> pd.DataFrame:
    """Override 'Current Status' in lms_df where Source Lead Id matches overrides."""
    if not overrides or lms_df.empty or 'Source Lead Id' not in lms_df.columns:
        return lms_df
    lms_df = lms_df.copy()
    lms_df['_opty_key'] = _normalise_lead_id(lms_df['Source Lead Id'])
    mask = lms_df['_opty_key'].isin(overrides)
    if mask.any():
        lms_df.loc[mask, 'Current Status'] = lms_df.loc[mask, '_opty_key'].map(overrides)
        print(f"  Status overrides applied: {mask.sum():,}/{len(lms_df):,} leads updated")
    lms_df = lms_df.drop(columns=['_opty_key'])
    return lms_df


def compute_supply_funnel(gen_df: pd.DataFrame, ld_df: pd.DataFrame,
                          lms_df: pd.DataFrame) -> dict:
    """Compute Generation → LD → LMS pipeline conversion metrics.

    Join chain:
      Gen.cd_lead_id = LD.cd_lead_id  →  get opty_id
      LD.opty_id     = LMS.Source Lead Id  →  get funnel status

    Returns dict with keys:
      totals          : {gen, ld, lms, retail, gen_to_ld_pct, ld_to_lms_pct, retail_rate}
      by_source_bucket: [{source_bucket, gen, ld, lms, retail, ...}, ...]
      monthly         : [{month, month_full, gen, ld, lms, retail, ...}, ...]
    """
    if gen_df is None or gen_df.empty or ld_df.empty:
        return {}

    # ── Step 1: LD lookup  cd_lead_id → opty_id ─────────────────────────────
    ld = ld_df.copy()
    # LD cd_lead_id is often float64 (e.g. 127183198.0) — strip the trailing .0
    ld['cd_lead_id'] = ld['cd_lead_id'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    ld['opty_id']    = ld['opty_id'].astype(str).str.strip()
    ld_lookup = (ld[['cd_lead_id', 'opty_id']]
                 .dropna()
                 .drop_duplicates('cd_lead_id'))

    # ── Step 2: LMS lookup  Source Lead Id → retail/booking/lost flags ───────
    lms_id_set      = set()
    lms_retail_set  = set()
    lms_booking_set = set()
    lms_lost_set    = set()
    if not lms_df.empty and 'Source Lead Id' in lms_df.columns:
        lms_sl = lms_df[['Source Lead Id', '_retail_attr', '_booking', '_lost']].copy()
        lms_sl['Source Lead Id'] = lms_sl['Source Lead Id'].astype(str).str.strip()
        lms_sl = lms_sl.drop_duplicates('Source Lead Id')
        lms_id_set      = set(lms_sl['Source Lead Id'])
        lms_retail_set  = set(lms_sl.loc[lms_sl['_retail_attr'], 'Source Lead Id'])
        lms_booking_set = set(lms_sl.loc[lms_sl['_booking'],     'Source Lead Id'])
        lms_lost_set    = set(lms_sl.loc[lms_sl['_lost'],        'Source Lead Id'])

    # ── Step 3: Enrich Gen with LD opty_id ──────────────────────────────────
    gen = gen_df.copy()
    gen = gen.merge(ld_lookup, on='cd_lead_id', how='left')
    gen['_in_ld']      = gen['opty_id'].notna() & (gen['opty_id'] != 'nan')
    gen['_in_lms']     = gen['opty_id'].isin(lms_id_set)
    gen['_retail']     = gen['opty_id'].isin(lms_retail_set)
    gen['_booking']    = gen['opty_id'].isin(lms_booking_set)
    gen['_lost']       = gen['opty_id'].isin(lms_lost_set)

    # ── Helper to build a metric row from a sub-dataframe ───────────────────
    def _metrics(grp):
        n       = len(grp)
        ld_cnt  = int(grp['_in_ld'].sum())
        lms_cnt = int(grp['_in_lms'].sum())
        ret_cnt = int(grp['_retail'].sum())
        bk_cnt  = int(grp['_booking'].sum())
        ls_cnt  = int(grp['_lost'].sum())
        return {
            'gen': n, 'ld': ld_cnt, 'lms': lms_cnt,
            'retail': ret_cnt, 'booking': bk_cnt, 'lost': ls_cnt,
            'gen_to_ld_pct':  safe_rate(ld_cnt,  n),
            'ld_to_lms_pct':  safe_rate(lms_cnt, ld_cnt),
            'lms_retail_rate':safe_rate(ret_cnt, lms_cnt),
            'gen_retail_rate':safe_rate(ret_cnt, n),
        }

    totals = _metrics(gen)

    # ── By source_bucket ─────────────────────────────────────────────────────
    by_sb = []
    if 'source_bucket' in gen.columns:
        for sb, grp in gen.groupby('source_bucket'):
            r = _metrics(grp)
            r['source_bucket'] = str(sb)
            by_sb.append(r)
        by_sb = sorted(by_sb, key=lambda x: -x['gen'])

    # ── Monthly (by Gen date) ─────────────────────────────────────────────────
    monthly = []
    date_col = 'Date' if 'Date' in gen.columns else None
    if date_col:
        gen['_mo'] = pd.to_datetime(gen[date_col], errors='coerce').dt.strftime('%b %Y')
        for mo, grp in gen.groupby('_mo'):
            if pd.isna(mo):
                continue
            r = _metrics(grp)
            r['month']      = str(mo).split()[0]
            r['month_full'] = str(mo)
            monthly.append(r)
        monthly = sorted(monthly,
                         key=lambda x: datetime.strptime(x['month_full'], '%b %Y')
                         if x.get('month_full') else datetime(2000, 1, 1))

    print(f"  Supply funnel: Gen={totals['gen']:,} → LD={totals['ld']:,} "
          f"({totals['gen_to_ld_pct']}%) → LMS={totals['lms']:,} "
          f"({totals['ld_to_lms_pct']}%) → Retail={totals['retail']:,} "
          f"({totals['lms_retail_rate']}% LMS CR%)")

    return {'totals': totals, 'by_source_bucket': by_sb, 'monthly': monthly}


# ─────────────────────────────────────────────────────────────────────────────
# SOURCE BUCKET ENRICHMENT
# ─────────────────────────────────────────────────────────────────────────────

def _enrich_lms_with_source_bucket(df: pd.DataFrame, gen_df: pd.DataFrame,
                                    ld_df: pd.DataFrame) -> pd.DataFrame:
    """Attach source_bucket from Gen data to LMS rows via Gen → LD → LMS join.

    Join path:
      Gen.cd_lead_id = LD.cd_lead_id  →  get opty_id
      LD.opty_id     = LMS.Source Lead Id  →  attach source_bucket
    """
    if gen_df is None or gen_df.empty or ld_df.empty:
        return df
    if 'source_bucket' not in gen_df.columns:
        return df
    if 'Source Lead Id' not in df.columns:
        return df

    # cd_lead_id → source_bucket from Gen (one-to-one)
    gen_sb = gen_df[['cd_lead_id', 'source_bucket']].copy()
    gen_sb['cd_lead_id'] = gen_sb['cd_lead_id'].astype(str).str.strip()
    gen_sb = gen_sb.dropna(subset=['source_bucket']).drop_duplicates('cd_lead_id')

    # cd_lead_id → opty_id from LD (strip trailing .0 from float-stored ints)
    ld = ld_df.copy()
    ld['cd_lead_id'] = (ld['cd_lead_id'].astype(str).str.strip()
                        .str.replace(r'\.0$', '', regex=True))
    ld['opty_id'] = ld['opty_id'].astype(str).str.strip()
    ld_lookup = ld[['cd_lead_id', 'opty_id']].dropna().drop_duplicates('cd_lead_id')

    # opty_id → source_bucket
    sb_map = gen_sb.merge(ld_lookup, on='cd_lead_id', how='inner')
    sb_map = sb_map[['opty_id', 'source_bucket']].drop_duplicates('opty_id')

    # Join to LMS via Source Lead Id = opty_id
    df = df.copy()
    df['_src_lid'] = df['Source Lead Id'].astype(str).str.strip()
    df = df.merge(sb_map.rename(columns={'opty_id': '_src_lid'}),
                  on='_src_lid', how='left')
    df = df.drop(columns=['_src_lid'])
    df['source_bucket'] = df['source_bucket'].fillna('Unknown')
    matched = (df['source_bucket'] != 'Unknown').sum()
    print(f"  Source bucket: {matched:,}/{len(df):,} LMS rows enriched "
          f"({safe_rate(matched, len(df))}%)")
    return df


_SB_ORDER = ['Organic', 'FB', 'Paid', 'Non-MS']


def _compute_sb_data(df: pd.DataFrame) -> tuple:
    """Compute key metrics per source bucket.

    Returns (ordered_bucket_names list, sb_data dict).
    sb_data[bucket] contains: kpis, monthly, monthly_update, models,
                               funnel_monthly, states, funnel.
    Only buckets with ≥10 LMS rows are included.
    'Unknown' is excluded from the filter list.
    """
    if 'source_bucket' not in df.columns:
        return [], {}

    present = [b for b in df['source_bucket'].dropna().unique() if b != 'Unknown']
    ordered = ([b for b in _SB_ORDER if b in present]
               + sorted([b for b in present if b not in _SB_ORDER]))

    sb_data = {}
    for sb in ordered:
        sub = df[df['source_bucket'] == sb].copy()
        if len(sub) < 10:
            continue
        sb_data[sb] = {
            'kpis':           compute_kpis(sub),
            'funnel':         compute_funnel(sub),
            'monthly':        compute_monthly(sub),
            'monthly_update': compute_monthly_update(sub),
            'models':         compute_models(sub),
            'funnel_monthly': compute_funnel_monthly(sub),
            'states':         compute_states(sub),
        }
    return ordered, sb_data


# ─────────────────────────────────────────────────────────────────────────────
# STATUS CLASSIFICATION
# ─────────────────────────────────────────────────────────────────────────────

def _date_filled(col: pd.Series) -> pd.Series:
    """Return bool mask: True where the date column has a real value (not NaN/empty/None)."""
    return col.replace('', pd.NA).notna() & (col.astype(str).str.strip() != '')


def classify_status(df: pd.DataFrame) -> pd.DataFrame:
    """Add boolean columns for each funnel stage.

    Retail is now split into three mutually exclusive groups:
      _retail_attr    : Retail Attribution (confirmed CPS retail — CR% denominator)
      _pending_retail : Pending Retail (retailed but attribution processing in progress)
      _retail_only    : Pure dealer Retail (has Retail Date or 'retail' status, NOT attr/pending)
      _pipeline_retail: _retail_only | _pending_retail  (retail not yet attributed)
      _retail         : union of all retail types (backward-compat for cohort/funnel helpers)
    """
    s = df['Current Status'].fillna('').str.lower()
    df = df.copy()

    # ── Retail Attribution: status contains 'retail attribution' OR RA date filled ─
    _has_ra_date = (_date_filled(df['Retail Attribution Date'])
                    if 'Retail Attribution Date' in df.columns
                    else pd.Series(False, index=df.index))
    df['_retail_attr'] = s.str.contains('retail attribution') | _has_ra_date

    # ── Pending Retail: status explicitly says 'pending retail' ─────────────────
    df['_pending_retail'] = s.str.contains('pending retail')

    # ── Pure Retail: has Retail Date OR 'retail' in status — but NOT attr/pending ─
    _has_retail_date = (_date_filled(df['Retail Date'])
                        if 'Retail Date' in df.columns
                        else pd.Series(False, index=df.index))
    _status_retail = s.str.contains('retail') & ~df['_retail_attr'] & ~df['_pending_retail']
    df['_retail_only'] = (_has_retail_date & ~df['_retail_attr'] & ~df['_pending_retail']) | _status_retail

    # ── Pipeline Retail = Pure Retail + Pending Retail ───────────────────────────
    df['_pipeline_retail'] = df['_retail_only'] | df['_pending_retail']

    # ── _retail = union of all retail (backward compat) ─────────────────────────
    df['_retail'] = df['_retail_attr'] | df['_retail_only'] | df['_pending_retail']

    df['_booking']   = (_date_filled(df['Booking Date'])
                         if 'Booking Date' in df.columns else pd.Series(False, index=df.index)) | \
                        s.str.contains('booking|booked')
    df['_test_ride'] = (_date_filled(df['TDR Date'])
                        if 'TDR Date' in df.columns else pd.Series(False, index=df.index)) | \
                       s.str.contains('test ride')
    df['_pq']        = s.str.contains('price quote|l1 verified|pqs')
    df['_cfv']       = s.str.contains('call for verification|cfv')
    df['_lost']      = s.str.contains('lost')
    df['_active']    = ~(df['_retail'] | df['_lost'])
    return df


def _clean_date_series(series: pd.Series) -> pd.Series:
    """Replace empty strings with NaT before date parsing."""
    return series.replace('', pd.NaT).replace('nan', pd.NaT)


def parse_month(series: pd.Series) -> pd.Series:
    """Parse a date column to 'Jan 2026' string format."""
    parsed = pd.to_datetime(_clean_date_series(series), errors='coerce')
    return parsed.dt.strftime('%b %Y')


def month_abbr(series: pd.Series) -> pd.Series:
    """Parse to short 'Jan', 'Feb', 'Mar' format."""
    parsed = pd.to_datetime(_clean_date_series(series), errors='coerce')
    return parsed.dt.strftime('%b')


# ─────────────────────────────────────────────────────────────────────────────
# METRIC COMPUTATION
# ─────────────────────────────────────────────────────────────────────────────

def safe_rate(num, denom, decimals=2):
    if not denom:
        return 0.0
    return round(num / denom * 100, decimals)


def compute_kpis(df: pd.DataFrame) -> dict:
    total = len(df)
    ra    = int(df['_retail_attr'].sum())
    ro    = int(df['_retail_only'].sum())
    pr    = int(df['_pending_retail'].sum()) if '_pending_retail' in df.columns else 0
    pipe  = int(df['_pipeline_retail'].sum()) if '_pipeline_retail' in df.columns else ro
    return {
        'total_leads':       total,
        # Retail Attribution = confirmed CPS retail; CR% is based on this
        'retail_ra':         ra,                   # backward-compat alias
        'retail_attribution':ra,
        # Pure retail (dealer upload without going through RA)
        'retail_only':       ro,
        # Pending Retail
        'pending_retail':    pr,
        # Pipeline Retail = Retail + Pending Retail
        'pipeline_retail':   pipe,
        'booking':           int(df['_booking'].sum()),
        'test_ride':         int(df['_test_ride'].sum()),
        'price_quote':       int(df['_pq'].sum()),
        'cfv':               int(df['_cfv'].sum()),
        'lost':              int(df['_lost'].sum()),
        'active':            int(df['_active'].sum()),
        # CR% = Retail Attribution / Total Leads
        'retail_rate':       safe_rate(ra, total),
        'booking_rate':      safe_rate(df['_booking'].sum(), total),
        'lost_rate':         safe_rate(df['_lost'].sum(), total),
        'pq_rate':           safe_rate(df['_pq'].sum(), total),
        'trd_rate':          safe_rate(df['_test_ride'].sum(), total),
        'pipeline_rate':     safe_rate(pipe, total),
        # Open Leads = CFV + PQ/L1 (leads in active verification/quote stages)
        'open_leads':        int(df['_cfv'].sum()) + int(df['_pq'].sum()),
        'open_pct':          safe_rate(int(df['_cfv'].sum()) + int(df['_pq'].sum()), total),
    }


def compute_funnel(df: pd.DataFrame) -> list:
    total = len(df)
    pipe = (df['_pipeline_retail'].sum() if '_pipeline_retail' in df.columns
            else df['_retail_only'].sum())
    stages = [
        ('Call for Verification',      df['_cfv'].sum()),
        ('Price Quote / L1 Verified',  df['_pq'].sum()),
        ('Test Ride (All)',             df['_test_ride'].sum()),
        ('Booking / Booked',           df['_booking'].sum()),
        # Retail Attribution = confirmed CPS retail (CR% basis)
        ('Retail Attribution',         df['_retail_attr'].sum()),
        # Pipeline Retail = Pure Retail + Pending Retail (not yet attributed)
        ('Pipeline Retail',            pipe),
        ('Lost (All)',                 df['_lost'].sum()),
        ('Active / Open',              df['_active'].sum()),
    ]
    return [{'stage': s, 'count': int(c), 'pct': safe_rate(c, total)} for s, c in stages]


def compute_monthly(df: pd.DataFrame, create_col='CreateTime') -> list:
    """Monthly lead creation + funnel breakdown, sorted chronologically."""
    df = df.copy()
    if create_col not in df.columns:
        return []
    df['_month_abbr'] = month_abbr(df[create_col])
    df['_month_full'] = parse_month(df[create_col])

    # Use full 'Mon YYYY' strings for chronological sort (handles cross-year: Dec 2025 < Jan 2026)
    month_fulls = sorted(
        df['_month_full'].dropna().unique(),
        key=lambda x: datetime.strptime(x, '%b %Y')
    )

    result = []
    for mf in month_fulls:
        sub = df[df['_month_full'] == mf]
        if len(sub) == 0:
            continue
        m   = mf.split()[0]   # abbreviated: 'Jan', 'Dec', etc.
        ra  = int(sub['_retail_attr'].sum()) if '_retail_attr' in sub.columns else 0
        pipe = int(sub['_pipeline_retail'].sum()) if '_pipeline_retail' in sub.columns else 0
        pq_m  = int(sub['_pq'].sum())
        cfv_m = int(sub['_cfv'].sum()) if '_cfv' in sub.columns else 0
        n     = len(sub)
        result.append({
            'month':           m,
            'month_full':      mf,
            'leads':           n,
            'retail_ra':       ra,
            'retail_attribution': ra,
            'pipeline_retail': pipe,
            'booking':         int(sub['_booking'].sum()),
            'test_ride':       int(sub['_test_ride'].sum()),
            'price_quote':     pq_m,
            'cfv':             cfv_m,
            'open_leads':      cfv_m + pq_m,
            'lost':            int(sub['_lost'].sum()),
            'retail_rate':     safe_rate(ra, n),
        })
    return result


def compute_monthly_update(df: pd.DataFrame) -> list:
    """
    Update-level monthly breakdown: each metric is bucketed by its OWN event date,
    not by lead creation date.
      - retail_attr  → Retail Attribution Date
      - retail_only  → Retail Date  (pure dealer retail, no RA)
      - booking      → Booking Date
      - lost         → Lost Date
      - leads        → CreateTime  (lead inflow, for context)
    This gives a true 'what happened in this month' view independent of when
    the lead was originally created.
    """
    df = df.copy()

    def _month_full(series, col):
        """Return 'Mon YYYY' Series for a datetime column, NaT-safe."""
        if col not in df.columns:
            return pd.Series([''] * len(df), index=df.index)
        return pd.to_datetime(df[col], errors='coerce').dt.strftime('%b %Y').fillna('')

    # Each event stream: (flag column, date column)
    streams = {
        'retail_attr':    ('_retail_attr',    'Retail Attribution Date'),
        'retail_only':    ('_retail_only',     'Retail Date'),
        'booking':        ('_booking',         'Booking Date'),
        'lost':           ('_lost',            'Lost Date'),
    }

    # Build month-keyed accumulators
    from collections import defaultdict
    buckets = defaultdict(lambda: {
        'retail_attr': 0, 'retail_only': 0, 'booking': 0, 'lost': 0, 'leads': 0
    })

    # Leads by CreateTime
    if 'CreateTime' in df.columns:
        df['_create_mfull'] = pd.to_datetime(df['CreateTime'], errors='coerce').dt.strftime('%b %Y').fillna('')
        for mo, cnt in df[df['_create_mfull'] != ''].groupby('_create_mfull').size().items():
            buckets[mo]['leads'] += int(cnt)

    # Event streams
    for key, (flag_col, date_col) in streams.items():
        if flag_col not in df.columns:
            continue
        sub = df[df[flag_col].astype(bool)].copy()
        if sub.empty:
            continue
        mo_series = _month_full(sub, date_col)
        if mo_series.eq('').all():
            # Date column missing or all blank — fall back to CreateTime
            mo_series = pd.to_datetime(sub.get('CreateTime', pd.Series()), errors='coerce').dt.strftime('%b %Y').fillna('')
        sub['_upd_mo'] = mo_series
        for mo, cnt in sub[sub['_upd_mo'] != ''].groupby('_upd_mo').size().items():
            buckets[mo][key] += int(cnt)

    if not buckets:
        return []

    # Sort all observed months chronologically
    def _sort_key(mo_str):
        try:
            return pd.to_datetime(mo_str, format='%b %Y')
        except Exception:
            return pd.Timestamp.max

    # Only keep months where at least some leads were created — stray event dates
    # on months outside the data period are almost always data-entry noise.
    valid_lead_months = {mo for mo, b in buckets.items() if b['leads'] > 0}

    result = []
    for mo_full in sorted(buckets.keys(), key=_sort_key):
        b = buckets[mo_full]
        if mo_full not in valid_lead_months:
            continue   # skip phantom months with no lead inflow
        ra    = b['retail_attr']
        leads = b['leads']
        result.append({
            'month':           mo_full.split(' ')[0],   # "Jan 2026" → "Jan"
            'month_full':      mo_full,
            'leads':           leads,
            'retail_ra':       ra,
            'retail_attr':     ra,
            'retail_only':     b['retail_only'],
            'pipeline_retail': 0,
            'retail_rate':     safe_rate(ra, leads),
            'booking':         b['booking'],
            'lost':            b['lost'],
            'price_quote':     0,
        })
    return result


def compute_models(df: pd.DataFrame, brand_prefix: str = '') -> list:
    if 'Model' not in df.columns:
        return []
    valid = df[df['Model'].notna() & (df['Model'].astype(str).str.strip() != '')]
    if valid.empty:
        return []
    def _model_row(g):
        n   = len(g)
        ra  = int(g['_retail_attr'].sum()) if '_retail_attr' in g.columns else 0
        pq  = int(g['_pq'].sum())
        cfv = int(g['_cfv'].sum()) if '_cfv' in g.columns else 0
        return {
            'model':           g.name,
            'leads':           n,
            'retail':          ra,
            'retail_attr':     ra,
            'pipeline_retail': int(g['_pipeline_retail'].sum()) if '_pipeline_retail' in g.columns else 0,
            'booking':         int(g['_booking'].sum()),
            'test_ride':       int(g['_test_ride'].sum()),
            'price_quote':     pq,
            'cfv':             cfv,
            'open_leads':      cfv + pq,
            'open_pct':        safe_rate(cfv + pq, n),
            'lost':            int(g['_lost'].sum()),
            'lost_rate':       safe_rate(int(g['_lost'].sum()), n),
            'retail_rate':     safe_rate(ra, n),
        }
    mc = valid.groupby('Model').apply(_model_row).tolist()
    return sorted(mc, key=lambda x: -x['leads'])


def compute_models_from_ld(ld_df: pd.DataFrame) -> list:
    """
    Compute model-level funnel breakdown from LD RawData file.
    The LD file has 'model' and 'Status_Name' columns.
    'Status_Name' uses the same status strings as LMS 'Current Status',
    so we can reuse the classify_status classifier.
    Also computes per-month breakdown for the model_monthly structure.
    """
    if ld_df.empty or 'model' not in ld_df.columns or 'Status_Name' not in ld_df.columns:
        return [], []

    # Classify funnel stages from Status_Name using text matching
    # (LD has no Retail Date column, so we can't use classify_status directly)
    ld = ld_df.copy()
    s = ld['Status_Name'].fillna('').str.lower()
    ld['_retail']    = s.str.contains('retail attribution|retail')
    ld['_booking']   = s.str.contains('booking|booked')
    ld['_test_ride'] = s.str.contains('test ride')
    ld['_pq']        = s.str.contains('price quote|l1 verified|pqs')
    ld['_cfv']       = s.str.contains('call for verification|cfv')
    ld['_lost']      = s.str.contains('lost')

    # Normalise model names (strip whitespace, keep raw OEM names)
    ld['model'] = ld['model'].fillna('').astype(str).str.strip()
    ld = ld[ld['model'] != '']

    models_list = []
    for mdl, g in ld.groupby('model'):
        n = len(g)
        models_list.append({
            'model':       mdl,
            'leads':       n,
            'retail':      int(g['_retail'].sum()),
            'booking':     int(g['_booking'].sum()),
            'test_ride':   int(g['_test_ride'].sum()),
            'price_quote': int(g['_pq'].sum()),
            'lost':        int(g['_lost'].sum()),
            'retail_rate': safe_rate(g['_retail'].sum(), n),
        })
    models_list = sorted(models_list, key=lambda x: -x['leads'])

    # Per-model per-month breakdown — flat format: [{model, month, leads, ...}, ...]
    # 'month' uses 'Jan 2026' style to match what the dashboard JS expects
    model_monthly = []
    if 'Date' in ld.columns:
        ld['_mo_full'] = pd.to_datetime(ld['Date'], errors='coerce').dt.strftime('%b %Y')
        top_models = [m['model'] for m in models_list[:15]]
        for mdl in top_models:
            mdf = ld[ld['model'] == mdl]
            for mo_full, sub in mdf.groupby('_mo_full'):
                if pd.isna(mo_full):
                    continue
                n = len(sub)
                model_monthly.append({
                    'model':          mdl,
                    'month':          str(mo_full),
                    'leads':          n,
                    'retail':         int(sub['_retail_attr'].sum()),
                    'retail_attr':    int(sub['_retail_attr'].sum()),
                    'pipeline_retail':int(sub['_pipeline_retail'].sum()) if '_pipeline_retail' in sub.columns else 0,
                    'booking':        int(sub['_booking'].sum()),
                    'price_quote':    int(sub['_pq'].sum()),
                    'lost':           int(sub['_lost'].sum()),
                    'retail_rate':    safe_rate(sub['_retail_attr'].sum(), n),
                })

    return models_list, model_monthly


def compute_dealers(df: pd.DataFrame) -> list:
    name_col  = 'L1 Organization Name' if 'L1 Organization Name' in df.columns else 'Registered Company Name'
    city_col  = 'L1 Dealer City' if 'L1 Dealer City' in df.columns else 'CityName'
    state_col = 'State'
    if name_col not in df.columns:
        return []

    def _dlr_row(g):
        n   = len(g)
        ra  = int(g['_retail_attr'].sum()) if '_retail_attr' in g.columns else 0
        pq  = int(g['_pq'].sum())
        cfv = int(g['_cfv'].sum()) if '_cfv' in g.columns else 0
        return {
            'dealer':          g.name,
            'city':            str(g[city_col].mode().iloc[0]) if city_col in g.columns and len(g[city_col].dropna()) else '',
            'state':           str(g[state_col].mode().iloc[0]) if state_col in g.columns and len(g[state_col].dropna()) else '',
            'leads':           n,
            'retail':          ra,
            'retail_attr':     ra,
            'pipeline_retail': int(g['_pipeline_retail'].sum()) if '_pipeline_retail' in g.columns else 0,
            'retail_rate':     safe_rate(ra, n),
            'booking':         int(g['_booking'].sum()),
            'test_ride':       int(g['_test_ride'].sum()),
            'price_quote':     pq,
            'cfv':             cfv,
            'open_leads':      cfv + pq,
            'open_pct':        safe_rate(cfv + pq, n),
            'lost':            int(g['_lost'].sum()),
        }
    result = df.groupby(name_col).apply(_dlr_row).tolist()
    return sorted(result, key=lambda x: -x['leads'])[:200]


def compute_dealers_for_names(df: pd.DataFrame, names: list) -> list:
    """Compute dealer rows for a specific list of dealer names (used to ensure flagged dealers are included)."""
    name_col  = 'L1 Organization Name' if 'L1 Organization Name' in df.columns else 'Registered Company Name'
    city_col  = 'L1 Dealer City' if 'L1 Dealer City' in df.columns else 'CityName'
    state_col = 'State'
    if name_col not in df.columns or not names:
        return []
    sub = df[df[name_col].isin(names)]
    if sub.empty:
        return []
    def _dlr_row(g):
        n   = len(g)
        ra  = int(g['_retail_attr'].sum()) if '_retail_attr' in g.columns else 0
        pq  = int(g['_pq'].sum()) if '_pq' in g.columns else 0
        cfv = int(g['_cfv'].sum()) if '_cfv' in g.columns else 0
        return {
            'dealer':          g.name,
            'city':            str(g[city_col].mode().iloc[0]) if city_col in g.columns and len(g[city_col].dropna()) else '',
            'state':           str(g[state_col].mode().iloc[0]) if state_col in g.columns and len(g[state_col].dropna()) else '',
            'leads':           n,
            'retail':          ra,
            'retail_attr':     ra,
            'pipeline_retail': int(g['_pipeline_retail'].sum()) if '_pipeline_retail' in g.columns else 0,
            'retail_rate':     safe_rate(ra, n),
            'booking':         int(g['_booking'].sum()) if '_booking' in g.columns else 0,
            'test_ride':       int(g['_test_ride'].sum()) if '_test_ride' in g.columns else 0,
            'price_quote':     pq,
            'cfv':             cfv,
            'open_leads':      cfv + pq,
            'open_pct':        safe_rate(cfv + pq, n),
            'lost':            int(g['_lost'].sum()) if '_lost' in g.columns else 0,
        }
    return sub.groupby(name_col).apply(_dlr_row).tolist()


def compute_states(df: pd.DataFrame) -> list:
    col = 'State'
    if col not in df.columns:
        return []
    def _state_row(g):
        n   = len(g)
        ra  = int(g['_retail_attr'].sum()) if '_retail_attr' in g.columns else 0
        pq  = int(g['_pq'].sum()) if '_pq' in g.columns else 0
        cfv = int(g['_cfv'].sum()) if '_cfv' in g.columns else 0
        bk  = int(g['_booking'].sum()) if '_booking' in g.columns else 0
        ls  = int(g['_lost'].sum()) if '_lost' in g.columns else 0
        return {
            'state':           g.name,
            'leads':           n,
            'retail':          ra,
            'retail_attr':     ra,
            'pipeline_retail': int(g['_pipeline_retail'].sum()) if '_pipeline_retail' in g.columns else 0,
            'retail_rate':     safe_rate(ra, n),
            'booking':         bk,
            'booking_rate':    safe_rate(bk, n),
            'lost':            ls,
            'lost_rate':       safe_rate(ls, n),
            'price_quote':     pq,
            'pq_rate':         safe_rate(pq, n),
            'open_leads':      cfv + pq,
            'open_pct':        safe_rate(cfv + pq, n),
        }
    result = df.groupby(col).apply(_state_row).tolist()
    return sorted(result, key=lambda x: -x['leads'])


def compute_state_monthly(df: pd.DataFrame, create_col: str = 'CreateTime') -> list:
    """Flat list: {state, month, month_full, leads, retail, booking, lost, open_leads, ...} for Geo period filter."""
    col = 'State'
    if col not in df.columns or create_col not in df.columns:
        return []
    df = df.copy()
    df['_mfull'] = df[create_col].dt.strftime('%b %Y')
    result = []
    for (state, mo_full), g in df.groupby([col, '_mfull']):
        if pd.isna(mo_full):
            continue
        n   = len(g)
        ra  = int(g['_retail_attr'].sum()) if '_retail_attr' in g.columns else 0
        pq  = int(g['_pq'].sum()) if '_pq' in g.columns else 0
        cfv = int(g['_cfv'].sum()) if '_cfv' in g.columns else 0
        bk  = int(g['_booking'].sum()) if '_booking' in g.columns else 0
        ls  = int(g['_lost'].sum()) if '_lost' in g.columns else 0
        result.append({
            'state':        state,
            'month':        str(mo_full),
            'month_full':   str(mo_full),
            'leads':        n,
            'retail':       ra, 'retail_attr': ra,
            'retail_rate':  safe_rate(ra, n),
            'booking':      bk, 'booking_rate': safe_rate(bk, n),
            'lost':         ls, 'lost_rate':    safe_rate(ls, n),
            'price_quote':  pq, 'pq_rate':      safe_rate(pq, n),
            'cfv':          cfv,
            'open_leads':   cfv + pq,
            'open_pct':     safe_rate(cfv + pq, n),
        })
    return sorted(result, key=lambda x: (x['month'], -x['leads']))


def compute_dealer_monthly(df: pd.DataFrame, create_col: str = 'CreateTime') -> list:
    """Flat list: {dealer, city, state, month, month_full, leads, retail, ...} for Geo period filter."""
    name_col  = 'L1 Organization Name' if 'L1 Organization Name' in df.columns else 'Registered Company Name'
    city_col  = 'L1 Dealer City' if 'L1 Dealer City' in df.columns else 'CityName'
    state_col = 'State'
    if name_col not in df.columns or create_col not in df.columns:
        return []
    df = df.copy()
    df['_mfull'] = df[create_col].dt.strftime('%b %Y')
    # Top 200 dealers by total leads to keep payload manageable
    top_dealers = df.groupby(name_col)[name_col].count().nlargest(200).index.tolist()
    result = []
    for (dlr, mo_full), g in df[df[name_col].isin(top_dealers)].groupby([name_col, '_mfull']):
        if pd.isna(mo_full):
            continue
        n   = len(g)
        ra  = int(g['_retail_attr'].sum()) if '_retail_attr' in g.columns else 0
        pq  = int(g['_pq'].sum()) if '_pq' in g.columns else 0
        cfv = int(g['_cfv'].sum()) if '_cfv' in g.columns else 0
        result.append({
            'dealer':      str(dlr),
            'city':        str(g[city_col].mode().iloc[0]) if city_col in g.columns and len(g[city_col].dropna()) else '',
            'state':       str(g[state_col].mode().iloc[0]) if state_col in g.columns and len(g[state_col].dropna()) else '',
            'month':       str(mo_full),
            'month_full':  str(mo_full),
            'leads':       n,
            'retail':      ra, 'retail_attr': ra,
            'retail_rate': safe_rate(ra, n),
            'booking':     int(g['_booking'].sum()) if '_booking' in g.columns else 0,
            'test_ride':   int(g['_test_ride'].sum()) if '_test_ride' in g.columns else 0,
            'price_quote': pq,
            'cfv':         cfv,
            'open_leads':  cfv + pq,
            'open_pct':    safe_rate(cfv + pq, n),
            'lost':        int(g['_lost'].sum()) if '_lost' in g.columns else 0,
        })
    return sorted(result, key=lambda x: (x['month'], -x['leads']))


def compute_lead_types(df: pd.DataFrame, ld_df: pd.DataFrame, oem_crm_col='Crm Lead Id') -> list:
    """Compute lead type breakdown.
    Prefers '_ld_lead_type' column on df (set when LD join ran), falls back to ld_df directly.
    Includes retail/booking/lost from the matched LMS rows.
    """
    # Prefer lead_type already merged onto df rows
    lt_col = '_ld_lead_type' if '_ld_lead_type' in df.columns else None
    if lt_col is None and (ld_df.empty or 'lead_type' not in ld_df.columns):
        return []

    def _lt_name(lt_int):
        return LT_DICT.get(lt_int, f'LT {lt_int}')

    if lt_col:
        sub = df[df[lt_col].notna()].copy()
        sub[lt_col] = pd.to_numeric(sub[lt_col], errors='coerce')
        sub = sub[sub[lt_col].notna()]
        result = []
        for lt_val, g in sub.groupby(lt_col):
            try:
                lt = int(lt_val)
            except Exception:
                continue
            n    = len(g)
            ra   = int(g['_retail'].sum()) if '_retail' in g.columns else 0
            pq   = int(g['_pq'].sum()) if '_pq' in g.columns else 0
            cfv  = int(g['_cfv'].sum()) if '_cfv' in g.columns else 0
            result.append({
                'lt_code':     lt,
                'name':        _lt_name(lt),
                'leads':       n,
                'retail':      ra,
                'booking':     int(g['_booking'].sum()) if '_booking' in g.columns else 0,
                'lost':        int(g['_lost'].sum()) if '_lost' in g.columns else 0,
                'price_quote': pq,
                'open_leads':  cfv + pq,
                'open_pct':    safe_rate(cfv + pq, n),
                'retail_rate': safe_rate(ra, n),
            })
        return sorted(result, key=lambda x: -x['leads'])

    # Fallback: use ld_df directly (no funnel breakdown available)
    lt_counts = ld_df.groupby('lead_type').size().reset_index(name='leads')
    result = []
    for _, row in lt_counts.iterrows():
        try:
            lt = int(row['lead_type'])
        except Exception:
            continue
        result.append({
            'lt_code':     lt,
            'name':        _lt_name(lt),
            'leads':       int(row['leads']),
            'retail':      0,
            'booking':     0,
            'lost':        0,
            'price_quote': 0,
            'open_leads':  0,
            'open_pct':    0.0,
            'retail_rate': 0.0,
        })
    return sorted(result, key=lambda x: -x['leads'])


def compute_cohort_flow(df: pd.DataFrame,
                         create_col='CreateTime',
                         retail_col='Retail Attribution Date') -> list:
    """Build cohort matrix: leads created in month X, retailed in month Y."""
    df = df.copy()
    if create_col not in df.columns:
        return []

    df['_create_month'] = parse_month(df[create_col])
    retailed = df[df['_retail']].copy()

    # Use Retail Attribution Date if available, else Retail Date, else CreateTime
    if retail_col in df.columns and df[retail_col].notna().sum() > 0:
        retailed['_retail_month'] = parse_month(retailed[retail_col])
    elif 'Retail Date' in df.columns:
        retailed['_retail_month'] = parse_month(retailed['Retail Date'])
    else:
        retailed['_retail_month'] = retailed['_create_month']

    # Get distinct months in order
    all_create = sorted(df['_create_month'].dropna().unique(),
                        key=lambda x: datetime.strptime(x, '%b %Y'))
    all_retail  = sorted(retailed['_retail_month'].dropna().unique(),
                         key=lambda x: datetime.strptime(x, '%b %Y'))

    result = []
    for cm in all_create:
        cohort_retail = retailed[retailed['_create_month'] == cm]
        row = {'create_month': cm, 'total': len(cohort_retail)}
        for rm in all_retail:
            cnt = int((cohort_retail['_retail_month'] == rm).sum())
            if cnt > 0:
                row[rm] = cnt
        result.append(row)

    return result


def compute_funnel_monthly(df: pd.DataFrame, create_col='CreateTime') -> list:
    """Per-month funnel breakdown."""
    df = df.copy()
    if create_col not in df.columns:
        return []
    df['_month_full'] = parse_month(df[create_col])

    stage_defs = [
        ('Call for Verification',     '_cfv'),
        ('Price Quote / L1 Verified', '_pq'),
        ('Test Ride (All)',            '_test_ride'),
        ('Booking / Booked',          '_booking'),
        ('Retail Attribution',        '_retail_attr'),
        ('Pipeline Retail',           '_pipeline_retail'),
        ('Lost (All)',                '_lost'),
        ('Active / Open',             '_active'),
    ]

    months = sorted(df['_month_full'].dropna().unique(),
                    key=lambda x: datetime.strptime(x, '%b %Y'))
    result = []
    for m in months:
        sub = df[df['_month_full'] == m]
        total = len(sub)
        stages = [{'stage': s, 'count': int(sub[col].sum()) if col in sub.columns else 0,
                   'pct': safe_rate(sub[col].sum() if col in sub.columns else 0, total)}
                  for s, col in stage_defs]
        result.append({'month': m, 'stages': stages})
    return result


STAGE_DEFS = [
    ('Call for Verification',     '_cfv'),
    ('Price Quote / L1 Verified', '_pq'),
    ('Test Ride (All)',            '_test_ride'),
    ('Booking / Booked',          '_booking'),
    ('Retail Attribution',        '_retail_attr'),
    ('Pipeline Retail',           '_pipeline_retail'),
    ('Lost (All)',                '_lost'),
    ('Active / Open',             '_active'),
]

def _funnel_stages(sub: 'pd.DataFrame') -> list:
    """Build funnel stage list from a (sub)dataframe that has classify_status cols."""
    total = len(sub)
    return [{'stage': s,
             'count': int(sub[col].sum()) if col in sub.columns else 0,
             'pct':   safe_rate(sub[col].sum() if col in sub.columns else 0, total)}
            for s, col in STAGE_DEFS]


def compute_funnel_by_model(df: pd.DataFrame, top_n: int = 20) -> dict:
    """Funnel stages keyed by model name (all-time)."""
    col = 'Model'
    if col not in df.columns:
        return {}
    top = df.groupby(col)[col].count().nlargest(top_n).index.tolist()
    return {mdl: _funnel_stages(df[df[col] == mdl]) for mdl in top}


def compute_funnel_by_model_monthly(df: pd.DataFrame, top_n: int = 20,
                                    create_col: str = 'CreateTime') -> dict:
    """Funnel stages keyed by model → month."""
    col = 'Model'
    if col not in df.columns or create_col not in df.columns:
        return {}
    df = df.copy()
    df['_mfull'] = df[create_col].dt.strftime('%b %Y')
    top = df.groupby(col)[col].count().nlargest(top_n).index.tolist()
    result = {}
    for mdl in top:
        sub = df[df[col] == mdl]
        months = sub['_mfull'].dropna().unique()
        result[mdl] = {m: _funnel_stages(sub[sub['_mfull'] == m]) for m in months}
    return result


def compute_funnel_by_state(df: pd.DataFrame, top_n: int = 20) -> dict:
    """Funnel stages keyed by state name (all-time)."""
    col = 'State'
    if col not in df.columns:
        return {}
    top = df.groupby(col)[col].count().nlargest(top_n).index.tolist()
    return {st: _funnel_stages(df[df[col] == st]) for st in top}


def compute_funnel_by_state_monthly(df: pd.DataFrame, top_n: int = 20,
                                    create_col: str = 'CreateTime') -> dict:
    """Funnel stages keyed by state → month."""
    col = 'State'
    if col not in df.columns or create_col not in df.columns:
        return {}
    df = df.copy()
    df['_mfull'] = df[create_col].dt.strftime('%b %Y')
    top = df.groupby(col)[col].count().nlargest(top_n).index.tolist()
    result = {}
    for st in top:
        sub = df[df[col] == st]
        months = sub['_mfull'].dropna().unique()
        result[st] = {m: _funnel_stages(sub[sub['_mfull'] == m]) for m in months}
    return result


def compute_model_state(df: pd.DataFrame, min_leads: int = 30) -> list:
    """Model × State cross-tab with funnel summary (for combined filter)."""
    mcol, scol = 'Model', 'State'
    if mcol not in df.columns or scol not in df.columns:
        return []
    result = []
    for (mdl, st), g in df.groupby([mcol, scol]):
        n = len(g)
        if n < min_leads:
            continue
        ra = int(g['_retail_attr'].sum()) if '_retail_attr' in g.columns else 0
        result.append({
            'model': mdl, 'state': st, 'leads': n,
            'cfv':          int(g['_cfv'].sum()) if '_cfv' in g.columns else 0,
            'price_quote':  int(g['_pq'].sum()) if '_pq' in g.columns else 0,
            'test_ride':    int(g['_test_ride'].sum()) if '_test_ride' in g.columns else 0,
            'booking':      int(g['_booking'].sum()) if '_booking' in g.columns else 0,
            'retail_attr':  ra,
            'pipeline_retail': int(g['_pipeline_retail'].sum()) if '_pipeline_retail' in g.columns else 0,
            'lost':         int(g['_lost'].sum()) if '_lost' in g.columns else 0,
            'active':       int(g['_active'].sum()) if '_active' in g.columns else 0,
            'retail_rate':  safe_rate(ra, n),
        })
    return sorted(result, key=lambda x: -x['leads'])


def compute_lt_by_dim(df: pd.DataFrame, dim_col: str, top_n: int = 20) -> list:
    """Lead-type breakdown grouped by a dimension (model or state)."""
    lt_col = '_ld_lead_type'
    if dim_col not in df.columns or lt_col not in df.columns:
        return []
    top_dims = df.groupby(dim_col)[dim_col].count().nlargest(top_n).index.tolist()
    result = []
    for dim_val in top_dims:
        sub = df[df[dim_col] == dim_val].copy()
        sub[lt_col] = pd.to_numeric(sub[lt_col], errors='coerce')
        sub = sub[sub[lt_col].notna()]
        for lt_val, g in sub.groupby(lt_col):
            try:
                lt = int(lt_val)
            except Exception:
                continue
            n = len(g)
            ra = int(g['_retail_attr'].sum()) if '_retail_attr' in g.columns else 0
            row = {dim_col.lower(): dim_val, 'lt_code': lt, 'name': LT_DICT.get(lt, f'LT {lt}'), 'leads': n,
                   'price_quote': int(g['_pq'].sum()) if '_pq' in g.columns else 0,
                   'test_ride':   int(g['_test_ride'].sum()) if '_test_ride' in g.columns else 0,
                   'booking':     int(g['_booking'].sum()) if '_booking' in g.columns else 0,
                   'retail_attr': ra, 'retail': ra,
                   'lost':        int(g['_lost'].sum()) if '_lost' in g.columns else 0,
                   'retail_rate': safe_rate(ra, n)}
            result.append(row)
    return sorted(result, key=lambda x: -x['leads'])


def compute_lt_by_city(df: pd.DataFrame, top_n: int = 50) -> list:
    """Lead-type breakdown grouped by city (top N cities by lead volume)."""
    lt_col = '_ld_lead_type'
    city_col = 'L1 Dealer City' if 'L1 Dealer City' in df.columns else ('CityName' if 'CityName' in df.columns else None)
    if city_col is None or lt_col not in df.columns:
        return []
    top_cities = df.groupby(city_col)[city_col].count().nlargest(top_n).index.tolist()
    result = []
    for city in top_cities:
        sub = df[df[city_col] == city].copy()
        sub[lt_col] = pd.to_numeric(sub[lt_col], errors='coerce')
        sub = sub[sub[lt_col].notna()]
        for lt_val, g in sub.groupby(lt_col):
            try:
                lt = int(lt_val)
            except Exception:
                continue
            n  = len(g)
            ra = int(g['_retail_attr'].sum()) if '_retail_attr' in g.columns else 0
            result.append({
                'city':        str(city),
                'lt_code':     lt,
                'name':        LT_DICT.get(lt, f'LT {lt}'),
                'leads':       n,
                'retail':      ra,
                'retail_attr': ra,
                'booking':     int(g['_booking'].sum()) if '_booking' in g.columns else 0,
                'lost':        int(g['_lost'].sum()) if '_lost' in g.columns else 0,
                'retail_rate': safe_rate(ra, n),
            })
    return sorted(result, key=lambda x: -x['leads'])


def compute_lt_by_model_monthly(df: pd.DataFrame, top_n: int = 15,
                                create_col: str = 'CreateTime') -> list:
    """Lead-type by model by month for period+model combined filter on LT tab."""
    lt_col = '_ld_lead_type'
    if 'Model' not in df.columns or lt_col not in df.columns or create_col not in df.columns:
        return []
    df = df.copy()
    df['_mfull'] = df[create_col].dt.strftime('%b %Y')
    df[lt_col] = pd.to_numeric(df[lt_col], errors='coerce')
    df = df[df[lt_col].notna()]
    top_models = df.groupby('Model')['Model'].count().nlargest(top_n).index.tolist()
    result = []
    for (mdl, mo, lt_val), g in df[df['Model'].isin(top_models)].groupby(['Model','_mfull', lt_col]):
        try:
            lt = int(lt_val)
        except Exception:
            continue
        n = len(g)
        ra = int(g['_retail_attr'].sum()) if '_retail_attr' in g.columns else 0
        result.append({'model': mdl, 'month': mo, 'lt_code': lt, 'name': f'LT {lt}',
                       'leads': n,
                       'price_quote': int(g['_pq'].sum()) if '_pq' in g.columns else 0,
                       'test_ride':   int(g['_test_ride'].sum()) if '_test_ride' in g.columns else 0,
                       'booking':     int(g['_booking'].sum()) if '_booking' in g.columns else 0,
                       'retail_attr': ra, 'retail': ra,
                       'lost':        int(g['_lost'].sum()) if '_lost' in g.columns else 0,
                       'retail_rate': safe_rate(ra, n)})
    return result


def compute_retail_dispersion(df: pd.DataFrame) -> list:
    """
    Retail dispersion: for retailed leads, did the customer buy the enquired model
    or a different one? Returns list of {enquired, purchased, count, month} for JS.
    """
    if 'Model' not in df.columns or 'Retail Purchased Model' not in df.columns:
        return []

    retailed = df[df['_retail'] & df['Retail Purchased Model'].notna()].copy()
    if retailed.empty:
        return []

    # Determine month column (prefer RA date, fall back to Retail Date)
    if 'Retail Attribution Date' in retailed.columns:
        retailed['_disp_mo'] = pd.to_datetime(retailed['Retail Attribution Date'], errors='coerce').dt.strftime('%b %Y')
    if '_disp_mo' not in retailed.columns or retailed['_disp_mo'].isna().all():
        if 'Retail Date' in retailed.columns:
            retailed['_disp_mo'] = pd.to_datetime(retailed['Retail Date'], errors='coerce').dt.strftime('%b %Y')
    if '_disp_mo' not in retailed.columns:
        retailed['_disp_mo'] = 'All'
    retailed['_disp_mo'] = retailed['_disp_mo'].fillna('All')

    result = []
    grp_cols = ['Model', 'Retail Purchased Model', '_disp_mo']
    for (enq, pur, mo), sub in retailed.groupby(grp_cols):
        result.append({
            'enquired':  str(enq),
            'purchased': str(pur),
            'month':     str(mo),
            'count':     len(sub),
        })
    return sorted(result, key=lambda x: -x['count'])


def generate_flags(df: pd.DataFrame) -> list:
    """Auto-generate RED/AMBER flags based on key metrics."""
    flags = []
    name_col = 'L1 Organization Name' if 'L1 Organization Name' in df.columns else None
    city_col = 'L1 Dealer City' if 'L1 Dealer City' in df.columns else 'CityName'

    if name_col and city_col in df.columns:
        dealer_city = df.groupby([name_col, city_col]).size().reset_index(name='leads')
        city_total  = df.groupby(city_col).size().reset_index(name='city_total')
        dealer_city = dealer_city.merge(city_total, on=city_col)
        dealer_city['share'] = dealer_city['leads'] / dealer_city['city_total'] * 100

        # Cities with 2 dealers where one has <30% share
        city_dealer_count = dealer_city.groupby(city_col)[name_col].count().reset_index(name='dealer_count')
        multi_city = city_dealer_count[city_dealer_count['dealer_count'] >= 2][city_col].tolist()
        for _, row in dealer_city[dealer_city[city_col].isin(multi_city)].iterrows():
            if row['share'] < 30:
                flags.append({
                    'severity': 'RED',
                    'category': 'Supply Concentration',
                    'entity':   str(row[name_col]),
                    'metric':   'City Lead Share',
                    'value':    f"{row['share']:.0f}%",
                    'threshold': '<50% (under-supplied)',
                    'description': f"{row[name_col]} receives only {row['share']:.0f}% of {row[city_col]} leads",
                    'dealers_in_city': int(city_dealer_count[city_dealer_count[city_col]==row[city_col]]['dealer_count'].iloc[0]),
                })

    # Dealers with 0% retail rate (min 50 leads)
    if name_col:
        dlr = df.groupby(name_col).agg(leads=('_retail', 'count'), retail=('_retail', 'sum')).reset_index()
        dlr['retail_rate'] = dlr['retail'] / dlr['leads'] * 100
        zero_retail = dlr[(dlr['leads'] >= 50) & (dlr['retail'] == 0)]
        for _, row in zero_retail.head(10).iterrows():
            flags.append({
                'severity': 'RED',
                'category': 'Zero Conversion',
                'entity':   str(row[name_col]),
                'metric':   'Retail Rate',
                'value':    '0%',
                'threshold': '>0%',
                'description': f"{row[name_col]} has 0 retails out of {row['leads']} leads",
            })

    return flags[:50]


# ─────────────────────────────────────────────────────────────────────────────
# FULL OEM DATA BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def build_oem_data(oem_name: str, lms_df: pd.DataFrame, ld_df: pd.DataFrame,
                   period: str, gen_df: pd.DataFrame = None,
                   bq_gen_df: pd.DataFrame = None) -> dict:
    """Build the complete DATA dict for an OEM dashboard."""
    print(f"  Computing metrics for {oem_name} ({len(lms_df):,} LMS rows)...")

    # ── Enrich LMS with model + lead_type from LD via Source Lead Id = opty_id ──
    # The LMS 'Model' column is typically empty; LD carries the correct model name.
    if not ld_df.empty and 'opty_id' in ld_df.columns and 'Source Lead Id' in lms_df.columns:
        ld_lookup = ld_df[['opty_id', 'model', 'lead_type']].copy()
        ld_lookup['opty_id'] = ld_lookup['opty_id'].astype(str).str.strip()
        lms_df = lms_df.copy()
        lms_df['Source Lead Id'] = lms_df['Source Lead Id'].astype(str).str.strip()
        lms_df = lms_df.merge(ld_lookup, left_on='Source Lead Id', right_on='opty_id', how='left')
        # Use LD model (fills blank LMS Model column)
        if 'model' in lms_df.columns:
            lms_df['Model'] = lms_df['model'].fillna(lms_df.get('Model', '')).fillna('')
        if 'lead_type' in lms_df.columns:
            lms_df['_ld_lead_type'] = lms_df['lead_type']
        matched = lms_df['Model'].notna() & (lms_df['Model'].astype(str).str.strip() != '')
        print(f"  LD join: {matched.sum():,}/{len(lms_df):,} rows enriched with model")

    # ── Fallback: if Model still empty, use 'Model.1' column (present in Jawa/Ampere LMS) ──
    # Jawa/Ampere LMS exports 'Model' as blank but populate 'Model.1' with the actual model name.
    if 'Model.1' in lms_df.columns:
        lms_df = lms_df.copy()
        # Ensure Model column is string type before filling
        lms_df['Model'] = lms_df['Model'].astype(str).replace({'nan': '', 'None': ''})
        blank_model = lms_df['Model'].str.strip() == ''
        if blank_model.any():
            lms_df.loc[blank_model, 'Model'] = lms_df.loc[blank_model, 'Model.1'].fillna('').astype(str)
            filled = (lms_df['Model'].str.strip() != '').sum()
            print(f"  Model.1 fallback: {blank_model.sum():,} blank rows filled → {filled:,} total with model")

    # ── Normalise model names: club variants before any groupby ──────────────
    def _normalise_model(m):
        m = str(m).strip()
        ml = m.lower()
        if ml.startswith('tvs iqube') and ml != 'tvs iqube':
            return 'TVS iQube'
        if ml == 'tvs apache rtr 160 4v':
            return 'TVS Apache RTR 160'
        return m
    lms_df = lms_df.copy()
    lms_df['Model'] = lms_df['Model'].apply(_normalise_model)

    df = classify_status(lms_df)

    # ── Attach source_bucket from Gen → LD → LMS join ────────────────────────
    df = _enrich_lms_with_source_bucket(df, gen_df, ld_df)

    # Parse date columns
    for col in ['CreateTime', 'Retail Date', 'Retail Attribution Date',
                'Booking Date', 'Lost Date', 'TDR Date']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    # ── Source bucket per-slice metrics ─────────────────────────────────────
    source_buckets, sb_data = _compute_sb_data(df)

    kpis      = compute_kpis(df)
    funnel    = compute_funnel(df)
    monthly   = compute_monthly(df)
    monthly_u = compute_monthly_update(df)
    models    = compute_models(df)   # now uses LD-enriched 'Model' column
    dealers   = compute_dealers(df)
    states    = compute_states(df)
    lead_types = compute_lead_types(df, ld_df)
    cohort_flow = compute_cohort_flow(df)
    funnel_monthly = compute_funnel_monthly(df)
    dispersion = compute_retail_dispersion(df)
    flags     = generate_flags(df)

    # ── Ensure all flagged dealers appear in dealer list ─────────────────────
    existing_dealer_names = {d['dealer'].lower() for d in dealers}
    flagged_names = [f['entity'] for f in flags if f.get('entity')
                     and f['entity'].lower() not in existing_dealer_names]
    if flagged_names:
        extra = compute_dealers_for_names(df, flagged_names)
        dealers = dealers + extra  # flagged dealers appended after top-200

    # ── Cross-tab data for model/state filters on Funnel & LT tabs ───────────
    funnel_by_model         = compute_funnel_by_model(df)
    funnel_by_model_monthly = compute_funnel_by_model_monthly(df)
    funnel_by_state         = compute_funnel_by_state(df)
    funnel_by_state_monthly = compute_funnel_by_state_monthly(df)
    model_state_xt          = compute_model_state(df)
    lt_by_model             = compute_lt_by_dim(df, 'Model')
    lt_by_state             = compute_lt_by_dim(df, 'State')
    lt_by_model_monthly     = compute_lt_by_model_monthly(df)
    lt_by_city              = compute_lt_by_city(df)
    state_monthly  = compute_state_monthly(df)
    dealer_monthly = compute_dealer_monthly(df)
    print(f"  Cross-tabs: {len(funnel_by_model)} models, {len(funnel_by_state)} states, "
          f"{len(model_state_xt)} model×state, {len(lt_by_model)} lt×model, {len(lt_by_state)} lt×state, "
          f"{len(lt_by_city)} lt×city, {len(state_monthly)} state×month, {len(dealer_monthly)} dealer×month")

    print(f"  Models found: {len(models)}")

    # ── model_monthly (INSERT level): bucket by CreateTime ────────────────────
    model_monthly = []
    valid_model = df[df['Model'].notna() & (df['Model'].astype(str).str.strip() != '')].copy()
    if not valid_model.empty and 'CreateTime' in valid_model.columns:
        valid_model['_mo_full'] = valid_model['CreateTime'].dt.strftime('%b %Y')
        top_models = [m['model'] for m in models[:15]]
        for mdl in top_models:
            mdf = valid_model[valid_model['Model'] == mdl]
            for mo_full, sub in mdf.groupby('_mo_full'):
                if pd.isna(mo_full):
                    continue
                n = len(sub)
                model_monthly.append({
                    'model':          mdl,
                    'month':          str(mo_full),
                    'leads':          n,
                    'retail':         int(sub['_retail_attr'].sum()),
                    'retail_attr':    int(sub['_retail_attr'].sum()),
                    'pipeline_retail':int(sub['_pipeline_retail'].sum()) if '_pipeline_retail' in sub.columns else 0,
                    'booking':        int(sub['_booking'].sum()),
                    'price_quote':    int(sub['_pq'].sum()),
                    'lost':           int(sub['_lost'].sum()),
                    'retail_rate':    safe_rate(sub['_retail_attr'].sum(), n),
                })

    # ── model_monthly_update (UPDATE level): bucket by Retail Attribution Date ─
    # Shows when retails ACTUALLY HAPPENED per model, regardless of lead creation date.
    model_monthly_update = []
    ra_date_col = 'Retail Attribution Date'
    if (ra_date_col in df.columns
            and '_retail_attr' in df.columns
            and 'Model' in df.columns):
        retailed = df[df['_retail_attr'].astype(bool) & df['Model'].notna()].copy()
        retailed['_ra_mo'] = pd.to_datetime(
            retailed[ra_date_col], errors='coerce'
        ).dt.strftime('%b %Y')
        # Fall back to Retail Date if RA date is blank
        if 'Retail Date' in retailed.columns:
            mask_blank = retailed['_ra_mo'].isna() | (retailed['_ra_mo'] == 'NaT')
            retailed.loc[mask_blank, '_ra_mo'] = pd.to_datetime(
                retailed.loc[mask_blank, 'Retail Date'], errors='coerce'
            ).dt.strftime('%b %Y')
        retailed = retailed[retailed['_ra_mo'].notna() & (retailed['_ra_mo'] != 'NaT')]
        top_models = [m['model'] for m in models[:15]]
        for (mdl, mo_full), sub in retailed[retailed['Model'].isin(top_models)].groupby(['Model', '_ra_mo']):
            n_retail = len(sub)
            # For context: how many leads of this model were created overall
            total_leads = int((valid_model['Model'] == mdl).sum()) if not valid_model.empty else 0
            model_monthly_update.append({
                'model':       mdl,
                'month':       str(mo_full),
                'leads':       total_leads,     # total model leads (denominator context)
                'retail':      n_retail,
                'retail_attr': n_retail,
                'pipeline_retail': 0,
                'booking':     0,
                'price_quote': 0,
                'lost':        0,
                'retail_rate': safe_rate(n_retail, total_leads),
            })
    else:
        # No RA date column — fall back to insert-level data
        model_monthly_update = model_monthly

    # ── Sort monthly arrays chronologically (Dec 2025 → Jan 2026 → ...) ──────
    def _mo_sort_key(row):
        try:    return datetime.strptime(row.get('month',''), '%b %Y')
        except: return datetime(2000, 1, 1)
    model_monthly        = sorted(model_monthly,        key=_mo_sort_key)
    model_monthly_update = sorted(model_monthly_update, key=_mo_sort_key)

    ra_date_avail = ra_date_col in df.columns and df[ra_date_col].notna().sum() > 0
    booking_date_avail = 'Booking Date' in df.columns and df['Booking Date'].notna().sum() > 0
    lost_date_avail = 'Lost Date' in df.columns and df['Lost Date'].notna().sum() > 0
    print(f"  Event dates available — RA: {ra_date_avail}, Booking: {booking_date_avail}, Lost: {lost_date_avail}")
    print(f"  model_monthly rows: {len(model_monthly)} (insert) | {len(model_monthly_update)} (update)")

    return {
        'oem':                 oem_name,
        'period':              period,
        'kpis':                kpis,
        'funnel':              funnel,
        'monthly':             monthly,
        'models':              models,
        'lead_types':          lead_types,
        'states':              states,
        'dealers':             dealers,
        'model_state':              model_state_xt,
        'model_monthly':            model_monthly,
        'model_monthly_update':     model_monthly_update,
        'retail_dispersion':        dispersion,
        'monthly_update':           monthly_u,
        'cohort_flow':              cohort_flow,
        'cohort_flow_by_model':     [],
        'model_update':             [],
        'funnel_monthly':           funnel_monthly,
        'funnel_by_model':          funnel_by_model,
        'funnel_by_model_monthly':  funnel_by_model_monthly,
        'funnel_by_state':          funnel_by_state,
        'funnel_by_state_monthly':  funnel_by_state_monthly,
        'lt_monthly':               [],
        'lt_by_model':              lt_by_model,
        'lt_by_model_monthly':      lt_by_model_monthly,
        'lt_by_state':              lt_by_state,
        'lt_by_city':               lt_by_city,
        'state_monthly':            state_monthly,
        'dealer_monthly':           dealer_monthly,
        'flags':                    flags,
        'supply_funnel':            (compute_supply_funnel_bq(bq_gen_df, df)
                                     if (bq_gen_df is not None and not bq_gen_df.empty)
                                     else compute_supply_funnel(gen_df, ld_df, df)),
        'source_buckets':           source_buckets,
        'sb_data':                  sb_data,
        'data_as_of':               datetime.now().strftime('%-d %b %Y'),
    }


# ─────────────────────────────────────────────────────────────────────────────
# HTML UPDATER
# ─────────────────────────────────────────────────────────────────────────────

def update_dashboard_html(html_path: Path, data: dict, flags: list = None,
                          lt_n2c: dict = None):
    """Replace the DATA, FLAGS, and LT_N2C constants in an HTML dashboard."""
    if not html_path.exists():
        print(f"  SKIP (file not found): {html_path.name}")
        return False

    content = html_path.read_text(encoding='utf-8')
    oem_key = data['oem'].replace(' ', '_').upper()

    # Sanitize: remove control characters that break JSON parsing in browsers
    def _sanitize(obj):
        if isinstance(obj, str):
            import re as _re
            # Remove ALL ASCII control chars (0x00–0x1f) and DEL (0x7f)
            # json.dumps *would* escape \t/\n/\r, but some environments fail on literal
            # control bytes, so we strip them proactively.
            return _re.sub(r'[\x00-\x1f\x7f]', ' ', obj).strip()
        if isinstance(obj, dict):
            return {k: _sanitize(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [_sanitize(i) for i in obj]
        return obj

    data_clean  = _sanitize(data)
    flags_clean = _sanitize(flags or data.get('flags', []))
    lt_n2c_clean = _sanitize(lt_n2c or {})

    data_json   = json.dumps(data_clean, ensure_ascii=False, separators=(',', ':'))
    flags_json  = json.dumps(flags_clean, ensure_ascii=False, separators=(',', ':'))
    lt_n2c_json = json.dumps(lt_n2c_clean, ensure_ascii=False, separators=(',', ':'))

    oem_lower = data['oem'].lower()
    flags_var  = f"{data['oem'].upper()}_FLAGS" if data['oem'] == 'TVS' else \
                 f"{data['oem'].replace(' ', '_').upper()}_FLAGS"

    # Replace DATA line
    content = re.sub(r'const DATA = \{[^\n]+\};?',
                     f'const DATA = {data_json};', content, count=1)

    # Replace FLAGS line (handles TVS_FLAGS, AMPERE_FLAGS, JAWA_FLAGS etc.)
    content = re.sub(r'const \w+FLAGS[^\n]+;?',
                     f'const {flags_var} = window.FLAGS_DATA = {flags_json};',
                     content, count=1)

    # Replace _LT_N2C line if present
    if lt_n2c is not None:
        content = re.sub(r'const _LT_N2C = \{[^\n]+\};?',
                         f'const _LT_N2C = {lt_n2c_json};', content, count=1)

    # ── Patch FUNNEL_COLORS (TVS style): remove legacy 'Retail', add 'Pipeline Retail' ─
    new_fc = ("const FUNNEL_COLORS={"
              "'Call for Verification':'#3b82f6','Price Quote / L1 Verified':'#60a5fa',"
              "'Test Ride (All)':'#a78bfa','Booking / Booked':'#f59e0b',"
              "'Retail Attribution':'#10b981','Pipeline Retail':'#f97316',"
              "'Lost (All)':'#ef4444','Active / Open':'#94a3b8'"
              "};")
    content = re.sub(r'const FUNNEL_COLORS=\{[^}]+\};', new_fc, content, count=1)

    # ── Patch SC (Jawa/Ampere style): add 'Pipeline Retail', remove legacy 'Retail' ─
    new_sc = ("const SC={"
              "'Call for Verification':'#3b82f6','Price Quote':'#0ea5e9','L1 Verified':'#8b5cf6',"
              "'Price Quote / L1 Verified':'#60a5fa',"
              "'Test Ride Requested':'#10b981','Test Ride Scheduled':'#34d399','Test Ride Completed':'#059669',"
              "'Test Ride (All)':'#a78bfa',"
              "'Booking':'#f59e0b','Booking / Booked':'#f59e0b',"
              "'Retail Attribution':'#10b981','Pipeline Retail':'#f97316',"
              "'Lost Not Purchased':'#ef4444','Lost Purchased':'#f97316','Lost Not Contactable':'#dc2626',"
              "'Lost (All)':'#ef4444','Active / Open':'#94a3b8'"
              "};")
    content = re.sub(r'const SC=\{[^}]+\};', new_sc, content, count=1)

    # ── Fix donut chart rt/bk to use new stage names ─────────────────────────────
    content = content.replace(
        "const rt=get('Retail')+get('Retail Attribution'),bk=get('Booking');",
        "const rt=get('Retail Attribution')+get('Pipeline Retail'),bk=get('Booking / Booked')||get('Booking');"
    )

    # ── Patch static KPI card block ───────────────────────────────────────────
    kpis       = data.get('kpis', {})
    period_str = data.get('period', '')

    def _fv(n):
        try:
            return f"{int(n):,}"
        except Exception:
            return str(n)

    def _fp(r):
        try:
            return f"{float(r):.2f}%"
        except Exception:
            return "0.00%"

    kpi_html = (
        f'<div class="filter-row" style="margin-bottom:10px;" id="ov-filter-row">\n'
        f'  <span style="font-size:11px;color:#64748b;font-weight:600;">Month:</span>\n'
        f'  <div style="display:flex;gap:4px;" id="ov-period-bar">\n'
        f'    <button class="mdl-pb active" onclick="setOvPeriod(\'all\',this)">All</button>\n'
        f'  </div>\n'
        f'</div>\n'
        f'<div class="kpi-row">\n'
        f'  <div class="kpi c1"><div class="kpi-lbl">Total Leads</div>'
        f'<div class="kpi-val" id="ov-k-leads">{_fv(kpis.get("total_leads",0))}</div>'
        f'<div class="kpi-sub" id="ov-k-leads-sub">LD+LMS merged · {period_str}</div></div>\n'
        f'  <div class="kpi c2"><div class="kpi-lbl">Retail Attribution</div>'
        f'<div class="kpi-val" id="ov-k-retail">{_fv(kpis.get("retail_attribution",0))}</div>'
        f'<div class="kpi-sub" id="ov-k-retail-sub">{_fp(kpis.get("retail_rate",0))} CR%</div></div>\n'
        f'  <div class="kpi" style="border-left-color:#f97316">'
        f'<div class="kpi-lbl">Pipeline Retail</div>'
        f'<div class="kpi-val" id="ov-k-pipeline">{_fv(kpis.get("pipeline_retail",0))}</div>'
        f'<div class="kpi-sub" id="ov-k-pipeline-sub">Retail + Pending · {_fp(kpis.get("pipeline_rate",0))}</div></div>\n'
        f'  <div class="kpi c3"><div class="kpi-lbl">Booking</div>'
        f'<div class="kpi-val" id="ov-k-booking">{_fv(kpis.get("booking",0))}</div>'
        f'<div class="kpi-sub" id="ov-k-booking-sub">{_fp(kpis.get("booking_rate",0))} booking rate</div></div>\n'
        f'  <div class="kpi c6"><div class="kpi-lbl">Test Ride Sched.</div>'
        f'<div class="kpi-val" id="ov-k-tr">{_fv(kpis.get("test_ride",0))}</div>'
        f'<div class="kpi-sub" id="ov-k-tr-sub">{_fp(kpis.get("trd_rate",0))} TR rate</div></div>\n'
        f'  <div class="kpi c4"><div class="kpi-lbl">Price Quote / L1</div>'
        f'<div class="kpi-val" id="ov-k-pq">{_fv(kpis.get("price_quote",0))}</div>'
        f'<div class="kpi-sub" id="ov-k-pq-sub">{_fp(kpis.get("pq_rate",0))} PQ rate</div></div>\n'
        f'  <div class="kpi" style="border-left-color:#e84545">'
        f'<div class="kpi-lbl">Lost (Total)</div>'
        f'<div class="kpi-val" id="ov-k-lost">{_fv(kpis.get("lost",0))}</div>'
        f'<div class="kpi-sub" id="ov-k-lost-sub">{_fp(kpis.get("lost_rate",0))} lost rate</div></div>\n'
        f'  <div class="kpi" style="border-left-color:#7c3aed">'
        f'<div class="kpi-lbl">Open Leads</div>'
        f'<div class="kpi-val" id="ov-k-open">{_fv(kpis.get("open_leads", kpis.get("cfv",0)+kpis.get("price_quote",0)))}</div>'
        f'<div class="kpi-sub" id="ov-k-open-sub">{_fp(kpis.get("open_pct",0))} open · CFV+PQ</div></div>\n'
        f'</div>'
    )
    # Depth-aware replacement: anchor on ov-filter-row if present (prevents duplicate
    # filter bars), otherwise fall back to kpi-row.  kpi_html always includes the
    # filter bar + kpi-row, so we must replace from the filter bar's opening tag.
    _ov_anchor = '<div class="filter-row" style="margin-bottom:10px;" id="ov-filter-row">'
    _kpi_anchor = '<div class="kpi-row">'
    replace_from = content.find(_ov_anchor)
    if replace_from == -1:
        replace_from = content.find(_kpi_anchor)
    kpi_row_start = content.find(_kpi_anchor, replace_from) if replace_from != -1 else -1
    if replace_from != -1 and kpi_row_start != -1:
        depth, i = 0, kpi_row_start
        while i < len(content):
            if content[i:i+4] == '<div':
                depth += 1; i += 4
            elif content[i:i+6] == '</div>':
                depth -= 1
                if depth == 0:
                    content = content[:replace_from] + kpi_html + content[i+6:]
                    break
                i += 6
            else:
                i += 1

    # ── Patch table column headers: rename legacy "Retail+RA" → "Retail Attribution" ─
    for old_hdr, new_hdr in [
        ("sortMdl('retail',DATA)\">Retail+RA ↕",  "sortMdl('retail_attr',DATA)\">Retail Attribution ↕"),
        ("sortSt('retail',DATA)\">Retail+RA ↕",   "sortSt('retail',DATA)\">Retail Attribution ↕"),
        ("sortDlr('retail',DATA)\">Retail+RA ↕",  "sortDlr('retail',DATA)\">Retail Attribution ↕"),
        ("sortDlr('retail',DATA)\">Retail+RA",     "sortDlr('retail',DATA)\">Retail Attribution"),
        ("sortLT('retail',DATA)\">Retail+RA ↕",   "sortLT('retail',DATA)\">Retail Attribution ↕"),
        ("sortLT('retail',DATA)\">Retail+RA",      "sortLT('retail',DATA)\">Retail Attribution"),
        (">Retail % ↕<",                           ">Conv % ↕<"),
        (">Rate ↕<",                               ">Conv % ↕<"),
    ]:
        content = content.replace(old_hdr, new_hdr)

    # ── Ensure old static Cohort Maturity Note is replaced with dynamic panel ──
    if 'id="insert-note-box"' in content:
        import re as _re
        content = _re.sub(
            r'<div class="box"[^>]*id="insert-note-box"[^>]*>.*?</div>',
            '<div class="box" id="insight-panel" style="border-radius:10px;padding:14px;"></div>',
            content, flags=_re.DOTALL
        )

    html_path.write_text(content, encoding='utf-8')
    print(f"  ✓ Updated: {html_path.name}")
    return True


# ─────────────────────────────────────────────────────────────────────────────
# MASTER DASHBOARD UPDATER
# ─────────────────────────────────────────────────────────────────────────────

def _sanitize_deep(obj):
    """Recursively strip ASCII control characters from all string values."""
    if isinstance(obj, str):
        return re.sub(r'[\x00-\x1f\x7f]', ' ', obj).strip()
    if isinstance(obj, dict):
        return {k: _sanitize_deep(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_deep(i) for i in obj]
    return obj


def _build_master_oem_block(oem: str, data: dict) -> dict:
    """Build per-OEM sub-object for const MD in the Master dashboard."""
    kpis = data.get('kpis', {})

    # Use data.monthly (compute_monthly output) — already sorted chronologically,
    # already has correct total leads per month (not stage sub-counts).
    # Fall back to funnel_monthly-based approach only if monthly is absent.
    raw_monthly = data.get('monthly', [])

    monthly = []
    if raw_monthly:
        for row in raw_monthly:
            mo_full = row.get('month_full', '')
            mo_short = row.get('month', mo_full.split()[0] if mo_full else '')
            total   = row.get('leads', 0)
            retail  = row.get('retail_ra', row.get('retail_attribution', 0))
            monthly.append({
                'month':       mo_short,
                'month_full':  mo_full,
                'leads':       total,
                'retail_ra':   retail,
                'retail':      retail,
                'booking':     row.get('booking', 0),
                'test_ride':   row.get('test_ride', 0),
                'price_quote': row.get('price_quote', 0),
                'lost':        row.get('lost', 0),
                'retail_rate': round(retail / total * 100, 2) if total else 0,
            })
    else:
        # Fallback: derive from funnel_monthly stages (less accurate for total leads)
        for fm_row in sorted(data.get('funnel_monthly', []),
                             key=lambda r: datetime.strptime(r['month'], '%b %Y') if r.get('month') else datetime(2000,1,1)):
            mo = fm_row.get('month', '')
            stages = {s['stage']: s['count'] for s in fm_row.get('stages', [])}
            total = max((s['count'] for s in fm_row.get('stages', [])), default=0)
            retail = stages.get('Retail Attribution', 0)
            parts = mo.split()
            mo_short = parts[0]
            monthly.append({
                'month':       mo_short,
                'month_full':  mo,
                'leads':       total,
                'retail_ra':   retail,
                'retail':      retail,
                'booking':     stages.get('Booking / Booked', 0),
                'test_ride':   stages.get('Test Ride (All)', 0),
                'price_quote': stages.get('Price Quote / L1 Verified', 0),
                'lost':        stages.get('Lost (All)', 0),
                'retail_rate': round(retail / total * 100, 2) if total else 0,
            })

    return {
        'oem':            oem,
        'period':         data.get('period', ''),
        'kpis':           kpis,
        'funnel':         data.get('funnel', []),
        'monthly':        monthly,
        'monthly_update': data.get('monthly_update', []),
        'models':         data.get('models', []),
        'lead_types':     data.get('lead_types', []),
        'states':         data.get('states', []),
        'dealers':        data.get('dealers', []),
        'source_buckets': data.get('source_buckets', []),
        'sb_data':        data.get('sb_data', {}),
    }


def update_master_dashboard(oem_data_full: dict):
    """Update the Master OEM Dashboard — patches const MD in the HTML."""
    html_path = WORKSPACE / "Master_OEM_Dashboard.html"
    if not html_path.exists():
        print("  SKIP: Master_OEM_Dashboard.html not found")
        return

    content = html_path.read_text(encoding='utf-8')

    # Build per-OEM blocks
    oem_blocks = {}
    for oem, data in oem_data_full.items():
        oem_blocks[oem.lower()] = _build_master_oem_block(oem, data)

    # Build all_kpis (portfolio totals across all OEMs)
    all_kpis = {}
    for key in ['total_leads','retail_ra','retail_attribution','booking','test_ride',
                'price_quote','cfv','lost','active']:
        all_kpis[key] = sum(d.get('kpis', {}).get(key, 0) for d in oem_data_full.values())
    total = all_kpis.get('total_leads', 1) or 1
    all_kpis['retail_rate']   = round(all_kpis.get('retail_ra', 0) / total * 100, 2)
    all_kpis['booking_rate']  = round(all_kpis.get('booking', 0) / total * 100, 2)
    all_kpis['lost_rate']     = round(all_kpis.get('lost', 0) / total * 100, 2)
    all_kpis['pq_rate']       = round(all_kpis.get('price_quote', 0) / total * 100, 2)

    # Build oem_compare — include all fields the JS uses in renderOEMTable
    oem_compare = []
    for oem, data in oem_data_full.items():
        k = data.get('kpis', {})
        tot = k.get('total_leads', 0) or 1
        oem_compare.append({
            'oem':          oem,
            'leads':        k.get('total_leads', 0),   # JS uses o.leads
            'total_leads':  k.get('total_leads', 0),
            'price_quote':  k.get('price_quote', 0),   # JS uses o.price_quote
            'pq_rate':      k.get('pq_rate', 0),       # JS uses o.pq_rate
            'retail':       k.get('retail_ra', 0),
            'retail_ra':    k.get('retail_ra', 0),
            'retail_rate':  k.get('retail_rate', 0),
            'booking':      k.get('booking', 0),
            'booking_rate': k.get('booking_rate', 0),
            'lost':         k.get('lost', 0),
            'lost_rate':    k.get('lost_rate', 0),
            'l1':           k.get('cfv', 0),
            'trd_rate':     k.get('trd_rate', 0),
            'period':       data.get('period', ''),
        })

    # Build cross_monthly — combined monthly totals across all OEMs
    # Use data.monthly (from compute_monthly) for accurate lead counts
    month_acc: dict = {}
    for data in oem_data_full.values():
        for row in data.get('monthly', []):
            mo_full = row.get('month_full', '')
            if not mo_full: continue
            if mo_full not in month_acc:
                month_acc[mo_full] = {'month_full': mo_full, 'leads': 0, 'retail_ra': 0, 'booking': 0, 'lost': 0}
            month_acc[mo_full]['leads']     += row.get('leads', 0)
            month_acc[mo_full]['retail_ra'] += row.get('retail_ra', row.get('retail_attribution', 0))
            month_acc[mo_full]['booking']   += row.get('booking', 0)
            month_acc[mo_full]['lost']      += row.get('lost', 0)
    cross_monthly = []
    for mo in sorted(month_acc.keys(), key=lambda m: datetime.strptime(m, '%b %Y')):
        parts = mo.split()
        mo_short = parts[0]   # plain abbreviated: "Jan", "Feb", "Dec"
        acc = month_acc[mo]
        cross_monthly.append({
            'month':      mo_short,
            'month_full': mo,
            'leads':      acc['leads'],
            'retail_ra':  acc['retail_ra'],
            'booking':    acc['booking'],
            'lost':       acc['lost'],
        })

    # Build combined_states (union of all OEM state data) — include retail_rate for JS chart
    state_acc: dict = {}
    for data in oem_data_full.values():
        for s in data.get('states', []):
            st = s.get('state', '')
            if not st: continue
            if st not in state_acc:
                state_acc[st] = {'state': st, 'leads': 0, 'retail': 0}
            state_acc[st]['leads']  += s.get('leads', 0)
            state_acc[st]['retail'] += s.get('retail', s.get('retail_attr', 0))
    # Compute retail_rate per state (used by renderCombinedStates JS)
    for st_data in state_acc.values():
        leads = st_data['leads'] or 1
        st_data['retail_rate'] = round(st_data['retail'] / leads * 100, 2)
    combined_states = sorted(state_acc.values(), key=lambda x: x['leads'], reverse=True)[:20]

    md = {
        'all_kpis':        all_kpis,
        'oem_compare':     oem_compare,
        'cross_monthly':   cross_monthly,
        'combined_states': combined_states,
        'data_as_of':      datetime.now().strftime('%-d %b %Y'),   # e.g. "27 Mar 2026"
        **oem_blocks,
    }

    # Sanitize all string values (removes literal control chars like tab in dealer names)
    # MUST happen before json.dumps, so re.sub replacement string won't interpret \t as tab
    md_clean = _sanitize_deep(md)
    master_json = json.dumps(md_clean, ensure_ascii=False, separators=(',', ':'))

    # Patch const MD = {...}; in HTML
    # Use lambda replacement to avoid re.sub interpreting \t/\n etc. in the JSON string
    replacement = f'const MD = {master_json};'
    subs_made = [0]
    def _replacer(m):
        subs_made[0] += 1
        return replacement
    new_content = re.sub(
        r'const MD\s*=\s*\{.*?\};',
        _replacer,
        content, count=1, flags=re.DOTALL
    )
    if subs_made[0] == 0:
        print("  WARNING: const MD pattern not found in Master dashboard — HTML may be out of date")
        new_content = content  # leave unchanged rather than corrupt it

    html_path.write_text(new_content, encoding='utf-8')
    print(f"  ✓ Updated: Master_OEM_Dashboard.html")


# ─────────────────────────────────────────────────────────────────────────────
# DETERMINE PERIOD STRING
# ─────────────────────────────────────────────────────────────────────────────

def infer_period(df: pd.DataFrame, create_col='CreateTime') -> str:
    if create_col not in df.columns:
        return ''
    dates = pd.to_datetime(df[create_col], errors='coerce').dropna()
    if dates.empty:
        return ''
    months = sorted(dates.dt.strftime('%b %Y').unique(),
                    key=lambda x: datetime.strptime(x, '%b %Y'))
    if len(months) == 1:
        return months[0]
    first, last = months[0], months[-1]
    first_mo, first_yr = first.split()
    last_mo,  last_yr  = last.split()
    # Same year: "Dec–Mar 2026" → no, show both years if they differ
    if first_yr == last_yr:
        return f"{first_mo}\u2013{last_mo} {last_yr}"
    else:
        return f"{first_mo} {first_yr}\u2013{last_mo} {last_yr}"


# ─────────────────────────────────────────────────────────────────────────────
# OEM CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

OEM_CONFIG = {
    'TVS': {
        'lms_folder':  LMS_DIR / 'TVS CPS',
        'ld_brand':    'TVS',
        'gen_brand':   'TVS',   # matches "TVS - Gen *.xlsx" filenames
        'html_file':   WORKSPACE / 'TVS_Lead_Performance_Dashboard.html',
    },
    'Ampere': {
        'lms_folder':  LMS_DIR / 'Ampere CPS',
        'ld_brand':    'Ampere Electric',
        'gen_brand':   'Ampere',
        'html_file':   WORKSPACE / 'Ampere_Lead_Performance_Dashboard.html',
    },
    'Jawa': {
        'lms_folder':  LMS_DIR / 'Jawa CPS',
        'ld_brand':    'Jawa',
        'gen_brand':   'Jawa',
        'html_file':   WORKSPACE / 'Jawa_Lead_Performance_Dashboard.html',
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # 1. Scan current files (LD, LMS, and Gen)
    current = scan_files(LD_DIR)
    current.update(scan_files(LMS_DIR))
    current.update(scan_files(GEN_DIR))
    current.update(scan_files(LATEST_STATUS_DIR))

    # 2. Load manifest
    old = load_manifest()

    # 3. Detect changes
    changed_files = {k for k, v in current.items() if old.get(k) != v}
    new_files = changed_files - set(old.keys())

    if not changed_files:
        print("No changes detected. All dashboards are up to date.")
        return

    print(f"\nDetected {len(changed_files)} changed file(s):")
    for f in sorted(changed_files):
        tag = '  NEW:' if f in new_files else '  MOD:'
        print(f"  {tag} {f}")

    # 4. Determine which OEMs are affected
    affected_oems = set()
    for f in changed_files:
        f_lower = f.lower()
        if 'tvs' in f_lower or 'ld rawdata' in f_lower:
            affected_oems.update(['TVS', 'Ampere', 'Jawa'])  # LD affects all
        if 'tvs' in f_lower:
            affected_oems.add('TVS')
        if 'ampere' in f_lower:
            affected_oems.add('Ampere')
        if 'jawa' in f_lower:
            affected_oems.add('Jawa')

    print(f"\nAffected OEMs: {', '.join(sorted(affected_oems))}")

    # 5. Rebuild affected OEM dashboards
    latest_status_overrides = load_latest_status_overrides()
    oem_data_full = {}  # full data dicts for master dashboard
    rebuilt = []

    for oem, cfg in OEM_CONFIG.items():
        if oem not in affected_oems:
            continue

        print(f"\n── Rebuilding {oem} dashboard ────────────────────────")
        try:
            lms_df    = load_lms_files(cfg['lms_folder'])
            ld_df     = load_ld_files(cfg['ld_brand'])

            # ── Gen data: BQ (live) or Excel (checkpoint) ──────────────────
            bq_gen_df = pd.DataFrame()
            if USE_BQ_FOR_GEN:
                bq_gen_df = load_gen_from_bq(oem)
            if not USE_BQ_FOR_GEN or bq_gen_df.empty:
                gen_df = load_gen_files(cfg.get('gen_brand', ''))
            else:
                gen_df = pd.DataFrame()   # BQ mode: skip Excel gen files

            if lms_df.empty:
                print(f"  WARNING: No LMS data found for {oem}")
                continue

            lms_df = apply_status_overrides(lms_df, latest_status_overrides)

            period = infer_period(lms_df)
            data   = build_oem_data(oem, lms_df, ld_df, period, gen_df,
                                    bq_gen_df=bq_gen_df)
            flags  = data.pop('flags', [])

            update_dashboard_html(cfg['html_file'], data, flags)

            data['period'] = period   # ensure period is in data for master
            oem_data_full[oem] = data
            rebuilt.append(oem)

        except Exception:
            print(f"  ERROR rebuilding {oem}:")
            traceback.print_exc()

    # 6. Update master dashboard
    if oem_data_full:
        print(f"\n── Updating Master Dashboard ─────────────────────────")
        update_master_dashboard(oem_data_full)

    # 7. Save updated manifest
    save_manifest(current)

    print(f"\n✅ Rebuild complete. Dashboards updated: {', '.join(rebuilt)}")
    print(f"   Changed files: {', '.join(sorted(changed_files))}")
    for oem, data in oem_data_full.items():
        k = data.get('kpis', {})
        print(f"   {oem}: {k.get('total_leads',0):,} leads | {k.get('retail_ra',0):,} retail ({k.get('retail_rate',0):.2f}%) | period: {data.get('period','?')}")


if __name__ == '__main__':
    main()
