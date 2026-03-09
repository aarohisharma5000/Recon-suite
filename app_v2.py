# app_v2.py — Paytm Reco Tool v2.0 (DuckDB + Parquet + Caching + Big File Mode)
# Supports: CSV / XLSX / XLS / XLSB (Excel files are converted to Parquet first)
# Big-file safe design: reconciliation runs in DuckDB (disk-based), not pandas merges

import os
import re
import json
import time
import hashlib
import uuid  # ✅ FIX: required for uuid.uuid4() used in excel_to_parquet
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import duckdb
import io
import zipfile
import subprocess
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# ===============================
# Column Normalization (ADD HERE)
# ===============================
def normalize_colname(col: str) -> str:
    col = str(col).strip().lower()
    col = col.replace("_", " ")
    col = " ".join(col.split())   # collapse multiple spaces
    col = re.sub(r"[^a-z0-9]+", "", col)
    return col

# ===============================
# Helpers for Downloads
# ===============================

# ---- Export cache init ----
if "export_cache_sig" not in st.session_state:
    st.session_state["export_cache_sig"] = None

if "export_cache" not in st.session_state:
    st.session_state["export_cache"] = None




def parquet_to_csv_bytes(con: duckdb.DuckDBPyConnection, pq_path: Path, max_rows: int = 800000):
    
    """
    Returns:
      files: dict {filename: bytes}
      meta:  dict {total_rows:int, parts:int, part_rows:list[int], base_stem:str}
    Splits into multiple CSV files if rows exceed max_rows.
    """
    if (pq_path is None) or (not pq_path.exists()):
        return {}, {"total_rows": 0, "parts": 0, "part_rows": [], "base_stem": pq_path.stem if pq_path else ""}

    total_rows = con.execute(
        f"SELECT COUNT(*) FROM parquet_scan('{pq_path.as_posix()}')"
    ).fetchone()[0]

    files = {}
    base_stem = pq_path.stem

    if total_rows <= max_rows:
        tmp_csv = pq_path.with_suffix(".tmp.csv")
        con.execute(f"""
            COPY (SELECT * FROM parquet_scan('{pq_path.as_posix()}'))
            TO '{tmp_csv.as_posix()}'
            (HEADER, DELIMITER ',');
        """)
        files[base_stem + ".csv"] = tmp_csv.read_bytes()
        tmp_csv.unlink(missing_ok=True)
        meta = {"total_rows": int(total_rows), "parts": 1, "part_rows": [int(total_rows)], "base_stem": base_stem}
        return files, meta

    parts = (total_rows // max_rows) + (1 if (total_rows % max_rows) else 0)
    part_rows = []

    for i in range(parts):
        offset = i * max_rows
        rows_this = int(min(max_rows, total_rows - offset))
        part_rows.append(rows_this)

        tmp_csv = pq_path.with_suffix(f".part{i+1}.tmp.csv")
        con.execute(f"""
            COPY (
                SELECT * FROM parquet_scan('{pq_path.as_posix()}')
                LIMIT {max_rows} OFFSET {offset}
            )
            TO '{tmp_csv.as_posix()}'
            (HEADER, DELIMITER ',');
        """)
        files[f"{base_stem}_part{i+1}.csv"] = tmp_csv.read_bytes()
        tmp_csv.unlink(missing_ok=True)

    meta = {"total_rows": int(total_rows), "parts": int(parts), "part_rows": part_rows, "base_stem": base_stem}
    return files, meta

def build_zip_bytes(files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name, b in files.items():
            if b:
                z.writestr(name, b)
    buf.seek(0)
    return buf.getvalue()



def parquet_scan_sql(files):
    """Return a DuckDB table function call for reading one or many parquet files.

    DuckDB supports: read_parquet('file.parquet') or read_parquet(['a.parquet','b.parquet']).
    We build the SQL snippet safely with forward slashes.
    """
    if isinstance(files, (list, tuple)):
        cleaned = []
        for f in files:
            s = str(f).replace('\\', '/').replace("'", "''")
            cleaned.append(f"'{s}'")
        return f"read_parquet([{', '.join(cleaned)}])"
    s = str(files).replace('\\', '/').replace("'", "''")
    return f"read_parquet('{s}')"

def build_summary_pdf(pdf_path: Path, summary_df: pd.DataFrame, app_version: str, focus_col: str, outdir: Path):
    pdf_path = Path(pdf_path)

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm
    )

    styles = getSampleStyleSheet()
    story = []

    title = Paragraph(f"<b>Paytm • Reconciliation Summary Report</b>", styles["Title"])
    subtitle = Paragraph(
        f"Tool Version: {app_version}<br/>"
        f"Generated On: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}<br/>"
        f"Summary Particular: {focus_col}",
        styles["Normal"]
    )

    story.append(title)
    story.append(Spacer(1, 8))
    story.append(subtitle)
    story.append(Spacer(1, 14))

    # table data
    table_data = [list(summary_df.columns)] + summary_df.astype(str).values.tolist()

    tbl = Table(table_data, repeatRows=1)

    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0d6efd")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cfd8e3")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fbff")]),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]))

    story.append(tbl)
    story.append(Spacer(1, 14))

    footer = Paragraph(
        "This report is generated from Paytm Reconciliation Suite and runs locally on the user machine.",
        styles["Italic"]
    )
    story.append(footer)

    doc.build(story)


# ---------------------------
# App Meta
# ---------------------------
import gc


def _num_clean_sql(col_sql: str) -> str:
    """Return a DuckDB SQL expression that safely parses a numeric string.
    - removes commas
    - trims spaces
    - treats empty/blank as NULL
    """
    # col_sql should already be a column reference like '"amount"' or 'a1."amount"'
    s = f"TRIM(CAST({col_sql} AS VARCHAR))"
    s = f"REPLACE({s}, ',', '')"
    # empty -> NULL then TRY_CAST
    return f"TRY_CAST(NULLIF({s}, '') AS DOUBLE)"

def _safe_atomic_replace(tmp_path: Path, final_path: Path, retries: int = 12, sleep_s: float = 0.5) -> None:
    """Windows-safe atomic replace with retries (handles transient file locks)."""
    tmp_path = Path(tmp_path)
    final_path = Path(final_path)

    last_err = None
    for _ in range(retries):
        try:
            # Ensure source exists
            if not tmp_path.exists():
                raise FileNotFoundError(f"Temp parquet not found: {tmp_path}")

            # Try to replace (atomic on Windows if same drive)
            os.replace(tmp_path, final_path)
            return
        except Exception as e:
            last_err = e
            # Try to release any lingering handles
            try:
                gc.collect()
            except Exception:
                pass
            time.sleep(sleep_s)

    # If still failing, raise the last error with context
    raise RuntimeError(f"Failed to atomically replace parquet after {retries} retries: {final_path}. Last error: {last_err}")

APP_VERSION = "v2.0"
LAST_UPDATED = "2026-02-17"
COMPANY_NAME = "Reconciliation Suite"

st.set_page_config(page_title="Reco Tool v2.0", layout="wide")

show_help = st.toggle("📘 Show Help & User Guide", value=False, key="toggle_help_tool2")


st.title(f"📌 {COMPANY_NAME} • Reconciliation Tool • {APP_VERSION}")

st.info(
    f"🟢 **Version:** {APP_VERSION}   |   🕒 **Last Updated:** {LAST_UPDATED}   |   "
    f"⚡ Engine: DuckDB + Parquet (Big-file mode)",
    icon="ℹ️"
)


if show_help:
    st.markdown("""
    <style>
    .help-card {
        padding: 16px 18px;
        border-radius: 14px;
        background: #f8fbff;
        border: 1px solid #d9e8ff;
        margin-bottom: 12px;
    }
    .help-title {
        font-size: 18px;
        font-weight: 700;
        margin-bottom: 6px;
        color: #0b5ed7;
    }
    .help-sub {
        font-size: 13px;
        color: #5b6470;
        margin-bottom: 0px;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-card">
        <div class="help-title">📘 Paytm Reconciliation Suite – Help & User Guide</div>
        <div class="help-sub">
            This guide explains how to use the tool step by step, what file formats are supported,
            what features are available, and how to download outputs correctly.
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-card">
        <div class="help-title">🔒 Data Security</div>
        <div class="help-sub">
            This tool runs locally on your machine. Files are processed on your system only and are not uploaded externally.
        </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Supported Formats", "CSV / XLSX / XLS / XLSB")
    c2.metric("Engine", "Pandas + DuckDB")
    c3.metric("Mode", "Local Machine")

    st.divider()

    with st.expander("🚀 Quick Start", expanded=True):
        st.markdown("""
        **Step 1** — Place or upload files in **File 1** and **File 2** sections  
        **Step 2** — Select the **Key Column / matching logic**  
        **Step 3** — Select columns to compare  
        **Step 4** — Apply filter if required  
        **Step 5** — Click **Run Reco**  
        **Step 6** — Review summary, preview, duplicates, and downloads
        """)

    with st.expander("📂 What files can be used?", expanded=False):
        st.markdown("""
        The tool supports:

        - CSV
        - XLSX
        - XLS
        - XLSB

        You can upload:
        - one file vs one file
        - multiple files vs multiple files

        Multiple uploaded files on each side are combined before reconciliation.
        """)

    with st.expander("🧩 What does this tool support?", expanded=False):
        st.markdown("""
        Main capabilities:

        - Multi-file upload on both sides
        - Key-based matching
        - Common-column comparison
        - Selected-column comparison
        - Duplicate detection
        - Match / mismatch / only-in outputs
        - CSV / ZIP downloads
        - Large-file handling on local machine
        """)

    with st.expander("🧹 What cleaning is applied?", expanded=False):
        st.markdown("""
        The tool can normalize and clean values to improve reconciliation quality.

        Typical cleaning includes:

        - removing leading/trailing spaces
        - removing hidden spaces / non-breaking spaces
        - removing commas in numeric fields
        - trimming text before comparison
        - case-insensitive matching where applicable

        Example:
        - `1,234.00` → `1234.00`
        - ` lender_id ` → `lender_id`
        """)

    with st.expander("🔑 How key matching works", expanded=False):
        st.markdown("""
        The **Key Column** is used to match rows between File 1 and File 2.

        Good examples:
        - `loan_account_number`
        - `lender_loan_account_number`
        - `customer_id`
        - `merchant_id`

        Avoid using:
        - amount columns
        - balance columns
        - calculated fields

        If multiple key candidates are selected, the tool follows the configured priority order.
        """)

    with st.expander("📊 How compare columns work", expanded=False):
        st.markdown("""
        Compare columns are the fields checked after row matching.

        Examples:
        - `total_outstanding`
        - `principal`
        - `interest`
        - `cpb`

        Typical outputs:
        - Match
        - Mismatch
        - Only in File1
        - Only in File2
        - Duplicates
        """)

    with st.expander("🎯 How filter works", expanded=False):
        st.markdown("""
        Filter helps run reconciliation for a subset of records only.

        Example:
        - Filter Column: `lender_id`
        - Filter Value: `24`

        This means only records matching that value will be reconciled on both sides.
        """)

    with st.expander("🧾 Outputs generated", expanded=False):
        st.markdown("""
        After reconciliation, the tool can generate:

        - Summary
        - Matched records
        - Mismatched records
        - Only in File1
        - Only in File2
        - Duplicates
        - ZIP download containing all outputs

        Large outputs may be split into multiple files automatically.
        """)

    with st.expander("⚠️ Common issues and fixes", expanded=False):
        st.markdown("""
        **1. XLSB file error**  
        Install dependency:
        `pyxlsb`

        **2. Column not found / Binder Error**  
        Usually caused by mismatch between source column name and selected logical column

        **3. No valid compare columns found**  
        Re-check selected compare columns and common columns

        **4. File locked / permission denied**  
        Close the source Excel file before running reconciliation

        **5. Missing dependency error**  
        Install required package in the same virtual environment
        """)

    with st.expander("✅ Best Practices", expanded=False):
        st.markdown("""
        Recommended usage:

        - Keep source files closed while running reco
        - Verify selected key and compare columns
        - Review duplicate outputs before final conclusion
        - Use filter only when needed
        - Download ZIP for full output pack
        """)

    st.divider()

    st.subheader("📞 Support")
    st.info(
        "For issues, reach out to Finance Ops / tool owner.\n\n"
        "Suggested contact: aarohisharma5000@gmail.com"
    )


# ---------------------------
# Paths
# ---------------------------
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
IN_DIR_1 = DATA_DIR / "input" / "file1"
IN_DIR_2 = DATA_DIR / "input" / "file2"
CACHE_PQ_1 = DATA_DIR / "cache" / "parquet" / "file1"
CACHE_PQ_2 = DATA_DIR / "cache" / "parquet" / "file2"
OUT_RUNS = DATA_DIR / "output" / "runs"

for p in [IN_DIR_1, IN_DIR_2, CACHE_PQ_1, CACHE_PQ_2, OUT_RUNS]:
    p.mkdir(parents=True, exist_ok=True)

# ---------------------------
# Config: Key candidates (COALESCE priority)
# ---------------------------
KEY_CANDIDATES_PRIORITY = [
    normalize_colname("lender loan account number"),
    normalize_colname("loan account number"),
    normalize_colname("parent loan account number"),
    normalize_colname("customer id"),
    normalize_colname("id"),
]

# Fast schema/key detection sample size (first N rows)
SAMPLE_ROWS_FOR_DETECT = 200_000



# -----------------------------
# Key auto-detection helpers (v2.0 patch)
# -----------------------------
import hashlib

def _norm_colname(s: str) -> str:
    s = str(s).strip().lower()
    # keep alphanumerics only for matching
    return re.sub(r"[^a-z0-9]+", "", s)

_KEY_HINTS = [
    "lenderloanaccountnumber",
    "loanaccountnumber",
    "loanacctnumber",
    "loanaccnumber",
    "lan",
    "accountnumber",
    "loanid",
    "lenderaccountnumber",
]

def suggest_key_candidates(cols: list, top_k: int = 5) -> list:
    """Return best-guess key columns from a list of column names."""
    scored = []
    for c in cols:
        n = _norm_colname(c)
        score = 0
        for h in _KEY_HINTS:
            if h in n:
                score += 10
        if n.endswith("id") or n.endswith("number"):
            score += 2
        if "date" in n or "amount" in n or "sum" in n:
            score -= 5
        scored.append((score, c))
    scored.sort(key=lambda x: (-x[0], str(x[1]).lower()))
    # keep only positive-ish scores, else empty
    out = [c for sc, c in scored if sc >= 2][:top_k]
    return out

def key_sig(key_cols: list) -> str:
    s = "|".join([str(x) for x in key_cols])
    return hashlib.md5(s.encode("utf-8")).hexdigest()[:10]

def ensure_cache_key_sig(cache_dir: str, key_cols: list) -> bool:
    """Returns True if key signature changed (meaning: force reconvert)."""
    import os
    os.makedirs(cache_dir, exist_ok=True)
    sig = key_sig(key_cols)
    p = os.path.join(cache_dir, "_key_sig.txt")
    old = None
    if os.path.exists(p):
        try:
            old = open(p, "r", encoding="utf-8").read().strip()
        except Exception:
            old = None
    changed = (old != sig)
    if changed:
        try:
            open(p, "w", encoding="utf-8").write(sig)
        except Exception:
            pass
    return changed
# ---------------------------
# Helpers
# ---------------------------
def list_input_files(folder: Path):
    exts = {".csv", ".xlsx", ".xls", ".xlsb"}

    files = []
    for p in folder.glob("*"):
        if not p.is_file():
            continue
        if p.suffix.lower() not in exts:
            continue

        # ✅ FIX: ignore Excel temporary/lock files like "~$xxxx.xlsx"
        if p.name.startswith("~$"):
            continue

        files.append(p)

    files.sort(key=lambda x: (x.suffix.lower(), x.name.lower()))
    return files

def file_signature(paths):
    # stable signature for caching: path + size + mtime
    items = []
    for p in paths:
        try:
            stt = p.stat()
            items.append(f"{p.name}|{stt.st_size}|{int(stt.st_mtime)}")
        except Exception:
            items.append(f"{p.name}|NA|NA")
    raw = "||".join(items)
    return hashlib.md5(raw.encode("utf-8")).hexdigest()

def normalize_colname(col: str) -> str:
    """
    Normalize column names for matching:
    - lower case
    - strip spaces
    - remove underscores
    - collapse internal spaces
    - keep alphanumerics only
    """
    col = str(col).strip().lower()
    col = col.replace("_", " ")
    col = " ".join(col.split())   # collapse multiple spaces
    col = re.sub(r"[^a-z0-9]+", "", col)
    return col


def safe_cols(cols):
    # normalize for internal logic
    return [normalize_colname(c) for c in cols]

SCI_RE = re.compile(r"^\s*-?\d+(\.\d+)?[eE][+-]?\d+\s*$")

def clean_key_text(v):
    """Alphanumeric-preserving key clean (good for loan account numbers)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    t = str(v).strip().lstrip("'").strip('"')
    if t == "":
        return ""
    t = t.replace(",", "").replace(" ", "")
    if t in {"0", "0.0"}:
        return ""
    if t.endswith(".0"):
        t = t[:-2]
    # scientific notation best-effort (only works if it looks like sci)
    if SCI_RE.match(t):
        # DuckDB-side scientific handling is hard; here we do python best-effort for excel inputs
        try:
            from decimal import Decimal
            d = Decimal(t)
            t = format(d.quantize(Decimal(1)), "f")
        except Exception:
            pass
    return t.upper()

def pick_best_sheet_excel(path: Path, engine=None):
    """Read all sheets, choose sheet with max rows (same as your v1 logic)."""
    sheets = pd.read_excel(path, sheet_name=None, engine=engine)
    best_sheet, best_df, best_rows = None, None, -1
    for sh, dfx in sheets.items():
        if dfx is None:
            continue
        rows = len(dfx)
        if rows > best_rows:
            best_rows = rows
            best_sheet = sh
            best_df = dfx
    if best_df is None:
        best_df = pd.DataFrame()
    return best_sheet or "UNKNOWN", best_df

def excel_to_parquet(src: Path, out_pq: Path, key_cols: list):
    """Convert XLSX/XLS/XLSB to Parquet (adds _SOURCE_FILE, _SHEET, _KEY)."""
    if src.suffix.lower() == ".xlsb":
        sheet, df = pick_best_sheet_excel(src, engine="pyxlsb")
    else:
        sheet, df = pick_best_sheet_excel(src, engine=None)

    # Normalize headers
    cols = safe_cols(df.columns)

    # ✅ FIX: make duplicate column names unique (target, target__2, target__3...)
    seen = {}
    new_cols = []
    for c in cols:
        if c not in seen:
            seen[c] = 1
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}__{seen[c]}")

    df.columns = new_cols
    df["_SOURCE_FILE"] = src.name
    df["_SHEET"] = str(sheet)

    # build COALESCE raw key
    raw_key = pd.Series([""] * len(df), index=df.index, dtype=str)
    for c in key_cols:
        if c in df.columns:
            s = df[c]
            mask = (raw_key == "") & s.notna() & (s.astype(str).str.strip() != "")
            raw_key.loc[mask] = s.loc[mask].astype(str)

    df["_KEY"] = raw_key.apply(clean_key_text)
    # ✅ IMPORTANT: do NOT drop rows here. Runtime DuckDB key will be used.
    # df = df[df["_KEY"] != ""].copy()

    out_pq.parent.mkdir(parents=True, exist_ok=True)
    tmp_pq = out_pq.with_name(f"__build_{out_pq.stem}_{uuid.uuid4().hex}.parquet")
    # ✅ FIX: prevent pyarrow mixed-type conversion errors (e.g., date column having numbers + strings)
    # ✅ Fix pyarrow mixed-type conversion safely
    for c in df.columns:
        try:
            if df[c].dtype == "object":
                df[c] = df[c].astype(str)
        except Exception:
            # Handles duplicate column names returning DataFrame instead of Series
            df[c] = df[c].astype(str)
    df.to_parquet(tmp_pq, index=False)
    _safe_atomic_replace(tmp_pq, out_pq)
    return len(df)


def csv_to_parquet_duckdb(con, csv_path: Path, out_parquet: Path, key_cols=None):
    """
    Convert CSV -> Parquet using DuckDB.
    IMPORTANT (v2.0 big-file mode):
      - Do NOT generate _KEY here (key columns may be unknown / user-selected later).
      - Keep everything as VARCHAR to avoid scientific-notation / float coercions.
      - Add _SOURCE_FILE + _SHEET for trace.
    """
    csv_path = Path(csv_path)
    out_parquet = Path(out_parquet)

    tmp_parquet = out_parquet.with_suffix(out_parquet.suffix + ".tmp")

    # Use DuckDB's CSV auto-reader; keep all as text for stable joins
    sql = f"""
    COPY (
      SELECT
        *,
        '{csv_path.name}' AS _SOURCE_FILE,
        'CSV' AS _SHEET
      FROM read_csv_auto('{str(csv_path).replace("'", "''")}',
                         HEADER=TRUE,
                         ALL_VARCHAR=TRUE,
                         SAMPLE_SIZE={SAMPLE_ROWS_FOR_DETECT})
    ) TO '{str(tmp_parquet).replace("'", "''")}' (FORMAT PARQUET);
    """
    con.execute(sql)

    # Atomic-ish replace
    try:
        if out_parquet.exists():
            out_parquet.unlink()
    except Exception:
        pass
    tmp_parquet.replace(out_parquet)

def duck_clean_num_expr(expr: str) -> str:
    """Remove commas/spaces, treat blanks as NULL, then TRY_CAST to DOUBLE.
    This makes "1,234.00" == "1234".

    Pass an SQL expression (including table alias + quoting), e.g. f1."amount".
    """
    return (
        "TRY_CAST(NULLIF("
        "REGEXP_REPLACE(REGEXP_REPLACE(TRIM(CAST(" + expr + " AS VARCHAR)), ',', ''), ' ', ''),"
        "''"
        ") AS DOUBLE)"
    )

def detect_numeric_cols(con, pq_glob: str, compare_cols, key_exclude: set, sample_rows=200000):
    """
    Detect numeric columns by sampling: if TRY_CAST works for >=80% non-blank values.
    """
    # pull sample
    df = con.execute(f"SELECT * FROM parquet_scan('{pq_glob}') LIMIT {int(sample_rows)}").df()
    num_cols = []
    for c in compare_cols:
        if c in key_exclude:
            continue
        if c not in df.columns:
            continue
        s = df[c]
        nb = s.notna() & (s.astype(str).str.strip() != "") & (~s.astype(str).isin(["-", "--", "na", "n/a", "null", "none"]))
        if nb.sum() == 0:
            continue
        # Tolerant parsing: remove commas/spaces so "1,234.00" and "1234" both parse.
        s_clean = s.astype(str).str.replace(",", "", regex=False).str.strip()
        n = pd.to_numeric(s_clean, errors="coerce")
        ratio = (n.notna() & nb).sum() / max(1, nb.sum())
        if ratio >= 0.80:
            num_cols.append(c)
    return num_cols

def now_run_id():
    return datetime.now().strftime("run_%Y%m%d_%H%M%S")

def write_json(path: Path, obj):
    path.write_text(json.dumps(obj, indent=2), encoding="utf-8")

# ---------------------------
# UI: Input discovery
# ---------------------------
st.subheader("0) Big File Mode Inputs (Folder-based)")

files1 = list_input_files(IN_DIR_1)
files2 = list_input_files(IN_DIR_2)

cA, cB = st.columns(2)
with cA:
    st.caption(f"File1 folder: {IN_DIR_1}")
    st.dataframe(pd.DataFrame([{
        "File": f.name, "Ext": f.suffix.lower(),
        "Size_MB": round(f.stat().st_size/1024/1024, 2),
        "Modified": datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    } for f in files1]), use_container_width=True)

with cB:
    st.caption(f"File2 folder: {IN_DIR_2}")
    st.dataframe(pd.DataFrame([{
        "File": f.name, "Ext": f.suffix.lower(),
        "Size_MB": round(f.stat().st_size/1024/1024, 2),
        "Modified": datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    } for f in files2]), use_container_width=True)

if not files1 or not files2:
    st.warning("Please place at least 1 file in BOTH folders: data/input/file1 and data/input/file2")
    st.stop()

sig1 = file_signature(files1)
sig2 = file_signature(files2)

# ---------------------------
# UI: Config
st.subheader("1) Reco Configuration (Same as v1 UI)")

# ---- Discover columns early (so user can select BEFORE Run) ----
db_path = DATA_DIR / "cache" / "duckdb" / "reco.duckdb"
db_path.parent.mkdir(parents=True, exist_ok=True)
con_tmp = duckdb.connect(str(db_path))

try:    # Convert inputs to parquet if needed (but only if changed)
    def ensure_parquet(side_files, cache_dir):
        pq_files = []
        for f in side_files:
            stamp = int(f.stat().st_mtime_ns)
            out_pq = cache_dir / f"{f.stem}__{stamp}.parquet"
            if not out_pq.exists():
                if f.suffix.lower() == ".csv":
                    csv_to_parquet_duckdb(con_tmp, f, out_pq, KEY_CANDIDATES_PRIORITY)
                else:
                    excel_to_parquet(f, out_pq, KEY_CANDIDATES_PRIORITY)
            pq_files.append(str(out_pq).replace('\\','/'))

            # Best-effort cleanup: keep only newest 2 versions per input stem
            try:
                versions = sorted(cache_dir.glob(f"{f.stem}__*.parquet"), key=lambda p: p.stat().st_mtime, reverse=True)
                for old in versions[2:]:
                    try:
                        old.unlink()
                    except Exception:
                        pass
            except Exception:
                pass
        return pq_files

    with st.spinner("Preparing cached Parquet (for column discovery)..."):
        pq_files1 = ensure_parquet(files1, CACHE_PQ_1)
        pq_files2 = ensure_parquet(files2, CACHE_PQ_2)

    cols1 = con_tmp.execute(f"SELECT * FROM {parquet_scan_sql(pq_files1)} LIMIT 1").df().columns.tolist()
    cols2 = con_tmp.execute(f"SELECT * FROM {parquet_scan_sql(pq_files2)} LIMIT 1").df().columns.tolist()

finally:
    con_tmp.close()

cols1 = safe_cols(cols1)
cols2 = safe_cols(cols2)

internal_cols = {"_SOURCE_FILE", "_SHEET", "_KEY"}
common_cols = sorted(list((set(cols1) & set(cols2)) - internal_cols))

# Defaults present in both sides (used for auto-suggestion of key)
key_present = [c for c in KEY_CANDIDATES_PRIORITY if c in common_cols]

if not common_cols:
    st.error("No common columns found between both sides (after conversion). Please verify headers.")
    st.stop()


# ---- 🔑 Key selection (auto-detect + manual override) ----
st.markdown("### 🔑 Key Column (used for matching rows)")

# If your file doesn't have exact names, tool will suggest closest match.
auto_suggest = key_present if key_present else suggest_key_candidates(common_cols, top_k=5)

if key_present:
    st.success("✅ Default key column(s) found → " + " / ".join(key_present))
else:
    st.warning(
        "Default key names not found (this is ok). "
        "Please select the correct key column(s) from below."
    )

with st.expander("How key works (COALESCE)"):
    st.write(
        "Tool builds a single _KEY as: first non-blank value from your selected key columns "
        "(in the order you provide). This supports cases where some files have blank key in one column."
    )

# Allow user to choose key columns from common columns
key_cols_sel = st.multiselect(
    "Select key column(s) in BOTH sides (COALESCE priority will follow the order below)",
    options=common_cols,
    default=auto_suggest
)

# Optional: reorder by comma-separated priority (because multiselect ordering can be unclear)
priority_text_default = ", ".join(key_cols_sel) if key_cols_sel else ", ".join(auto_suggest)
priority_text = st.text_input(
    "Key priority (comma-separated, left-to-right). Leave as-is if ok.",
    value=priority_text_default
)

# Parse priority list
priority_list = [x.strip() for x in priority_text.split(",") if x.strip()]
# ✅ Only treat *selected key candidate columns* as "key columns" for exclusion,
# so user-selected compare columns never get accidentally removed.
_allowed_keys = set(key_cols_sel) if key_cols_sel else set(key_present)
KEY_CANDIDATES_ACTIVE = [c for c in priority_list if (c in common_cols and c in _allowed_keys)]
if not KEY_CANDIDATES_ACTIVE:
    KEY_CANDIDATES_ACTIVE = [c for c in (key_cols_sel or key_present) if c in common_cols]

if not KEY_CANDIDATES_ACTIVE:
    st.warning("⚠️ Key is not selected yet. You can still view columns, but Run will require a key selection.")
else:
    st.info("Key Mode: COALESCE → " + " / ".join(KEY_CANDIDATES_ACTIVE))

# ---- Same UI controls as your v1 ----
compare_mode = st.selectbox("Compare mode", ["Compare selected columns", "Compare common columns"])
numeric_tolerance = st.number_input("Numeric tolerance (+/-)", min_value=0.0, value=5.0, step=1.0)
treat_blanks_as_equal = st.checkbox("Treat blanks as equal", value=True)
show_rows = st.number_input("Rows to show in browser (preview)", min_value=100, value=2000, step=500)

st.markdown("### ✅ Matched Columns (common columns)")
st.caption(f"Common columns count: {len(common_cols)}")
with st.expander("Show common column names"):
    st.write(common_cols)

if compare_mode == "Compare common columns":
    compare_cols = [c for c in common_cols if c not in set(KEY_CANDIDATES_ACTIVE)]
else:
    options_cols = [c for c in common_cols if c not in set(KEY_CANDIDATES_ACTIVE)]

    # your previous default logic, but made safe
    suggested_defaults = common_cols[:min(12, len(common_cols))]
    default_cols = [c for c in suggested_defaults if c in options_cols]

    compare_cols = st.multiselect(
        "Select columns to compare (must exist in both)",
        options=options_cols,
        default=default_cols
    )
    if not compare_cols:
        st.warning("Please select at least one column to compare.")
        st.stop()

reco_focus_col = st.selectbox(
    "Summary 'Particular' column (choose ONE from selected/common columns)",
    options=compare_cols,
    index=0
)

st.markdown('### 🧹 Optional Filter (run reco on subset)')
st.caption('Example: if you have Lender ID in both files, you can run reco only for one lender by filtering both sides before reconciliation.')
filter_enabled = st.checkbox('Enable filter (apply on BOTH File1 & File2 before reco)', value=False)

# ---- Optional: dropdown of unique values for selected filter column (fast DuckDB DISTINCT) ----
@st.cache_data(show_spinner=False)
def _distinct_filter_values(sig1, sig2, filter_col_norm: str, limit=5000):
    """
    Robust distinct-value fetcher for filter dropdown.

    NOTE: UI uses normalized column names (via _norm_colname), e.g. "lenderid".
    Actual parquet may contain "lender_id" or different casing.
    We resolve the *real* parquet column name by matching normalized forms.
    """
    # --- helper: quote identifier safely for DuckDB ---
    def _qident(name: str) -> str:
        return '"' + str(name).replace('"', '""') + '"'

    # --- helper: resolve normalized col -> actual col in parquet schema ---
    def _resolve_actual_col(con, pq_files, norm_name: str):
        if not pq_files:
            return None
        try:
            # Use first parquet file to infer schema (consistent within the file-set)
            scan_sql = parquet_scan_sql([pq_files[0]])
            rows = con.execute(f"DESCRIBE SELECT * FROM {scan_sql}").fetchall()
            cols = [r[0] for r in rows]  # column names
        except Exception:
            return None

        # 1) exact match (case-insensitive)
        for c in cols:
            if str(c).lower() == str(norm_name).lower():
                return c

        # 2) normalized match (recommended)
        target = _norm_colname(norm_name)
        for c in cols:
            if _norm_colname(c) == target:
                return c
        return None

    con = duckdb.connect(database=":memory:")
    try:
        # Resolve actual column names per side
        actual1 = _resolve_actual_col(con, pq_files1, filter_col_norm)
        actual2 = _resolve_actual_col(con, pq_files2, filter_col_norm)

        # Prefer a column that exists on BOTH sides; else fallback to any one side
        actual = actual1 if (actual1 and actual2) else (actual1 or actual2)
        if not actual:
            return [], False

        col_sql = _qident(actual)

        vals = []
        for pq in (pq_files1 + pq_files2):
            scan_sql = parquet_scan_sql([pq])
            try:
                rows = con.execute(
                    f"SELECT DISTINCT TRIM(CAST({col_sql} AS VARCHAR)) AS v "
                    f"FROM {scan_sql} "
                    f"WHERE {col_sql} IS NOT NULL "
                    f"LIMIT {int(limit)}"
                ).fetchall()
                for r in rows:
                    v = r[0]
                    if v is None:
                        continue
                    v = str(v).strip()
                    if v and v not in vals:
                        vals.append(v)
                if len(vals) >= limit:
                    return vals[:limit], True
            except Exception:
                # If the column doesn't exist in a particular parquet, just skip it.
                continue

        return vals, False
    finally:
        con.close()
_filter_default = None
for _cand in ['lender_id','lender','lender code','lender_code','lender name','lender_name','partner_id','partner','merchant_id','merchant','lender loan account number']:
    _cand_norm = _norm_colname(_cand)
    if _cand_norm in common_cols:
        _filter_default = _cand_norm
        break
if _filter_default is None and common_cols:
    _filter_default = common_cols[0]

filter_col = st.selectbox('Filter column', options=common_cols, index=(common_cols.index(_filter_default) if _filter_default in common_cols else 0))

# If filter enabled, show a picker of available values (multiselect) + optional manual entry.
picked_values = []
if filter_enabled:
    use_picker = st.checkbox('Show available values (dropdown)', value=True)
    if use_picker:
        # Tie cache to current inputs (sig1/sig2) + column.
        vals, truncated = _distinct_filter_values(sig1, sig2, filter_col, limit=5000)
        if truncated:
            st.info("Showing first 5,000 unique values (too many to list). You can still type values manually below.")
        if not vals:
            st.info("No values found for this filter column.")
        picked_values = st.multiselect(
            "Pick filter value(s) (case-insensitive match)",
            options=vals,
            default=[]
        )
filter_value_raw = st.text_area(
    'Filter value(s) (exact match after TRIM, case-insensitive) — you can enter multiple values (comma or new line)',
    value='',
    height=90,
    placeholder='Example:\n85\n86\n87  (or 85,86,87)'
)
# Parse multiple values (comma/newline separated), de-duplicate (preserve order)
_tmp_vals = []
for _part in filter_value_raw.replace(',', '\n').splitlines():
    _v = _part.strip()
    if _v and (_v not in _tmp_vals):
        _tmp_vals.append(_v)
filter_values = _tmp_vals

# Merge dropdown-picked values (if any) into filter_values
for _v in (picked_values or []):
    _vv = str(_v).strip()
    if _vv and (_vv not in filter_values):
        filter_values.append(_vv)
if filter_enabled and not filter_values:
    st.warning('Filter is enabled but value is blank. Please enter at least one value or disable the filter.')
elif filter_enabled and filter_values:
    st.caption(f'Filter will run on **{len(filter_values)}** value(s).')
st.subheader("2) Run")

run_btn = st.button("✅ Run Reco (DuckDB Engine)", type="primary")

# ---------------------------

# Caching signature


# ---------------------------
filter_values_sig = ",".join([str(v).strip().upper() for v in (filter_values or []) if str(v).strip()])
config_sig_raw = f"{sig1}|{sig2}|{compare_mode}|{numeric_tolerance}|{treat_blanks_as_equal}|{reco_focus_col}|{','.join(compare_cols)}|{filter_enabled}|{filter_col}|{filter_values_sig}"

config_sig = hashlib.md5(config_sig_raw.encode("utf-8")).hexdigest()

# Store run cache state
if "run_cache_sig" not in st.session_state:
    st.session_state["run_cache_sig"] = None
if "run_cache_outdir" not in st.session_state:
    st.session_state["run_cache_outdir"] = None

# ✅ Run requires key selection
if run_btn and not KEY_CANDIDATES_ACTIVE:
    st.error("Please select the correct key column(s) above before running reconciliation.")
    st.stop()

if not run_btn:
    if st.session_state["run_cache_sig"] and st.session_state["run_cache_outdir"]:
        outdir = Path(st.session_state["run_cache_outdir"])
        use_cache = True
    else:
        st.stop()

# ---------------------------
# Execute Run (only when clicked OR cache exists)
# ---------------------------
use_cache = (st.session_state["run_cache_sig"] == config_sig and st.session_state["run_cache_outdir"])
if use_cache and not run_btn:
    outdir = Path(st.session_state["run_cache_outdir"])
else:
    outdir = OUT_RUNS / now_run_id()
    outdir.mkdir(parents=True, exist_ok=True)

log_lines = []
def log(msg):
    log_lines.append(msg)

# DuckDB connection (per run)
db_path = DATA_DIR / "cache" / "duckdb" / "reco.duckdb"
db_path.parent.mkdir(parents=True, exist_ok=True)
con = duckdb.connect(str(db_path))

try:
    prog = st.progress(0)
    status = st.empty()

    def step(pct, msg):
        prog.progress(pct)
        status.info(msg)
        log(msg)

    # ---------------------------
    # Step 1: Convert inputs to Parquet (with caching)
    # ---------------------------
    step(10, "Step 1/5: Converting inputs to Parquet (cached if unchanged)...")

    def to_cached_parquet(side_files, cache_dir):
        pq_files = []
        # if key mapping changed, reconvert all cached parquet (because _KEY depends on it)
        force = ensure_cache_key_sig(str(cache_dir), KEY_CANDIDATES_ACTIVE)

        for f in side_files:
            stamp = int(f.stat().st_mtime_ns)
            out_pq = cache_dir / f"{f.stem}__{stamp}.parquet"
            # convert only if missing or older than source
            if not out_pq.exists():
                t0 = time.time()
                if f.suffix.lower() == ".csv":
                    rows = csv_to_parquet_duckdb(con, f, out_pq, KEY_CANDIDATES_ACTIVE)
                else:
                    rows = excel_to_parquet(f, out_pq, KEY_CANDIDATES_ACTIVE)
                log(f"Converted {f.name} -> {out_pq.name} | rows={rows:,} | {round(time.time()-t0,2)}s")
            pq_files.append(out_pq)
        return pq_files

    pq1 = to_cached_parquet(files1, CACHE_PQ_1)
    pq2 = to_cached_parquet(files2, CACHE_PQ_2)

    pq_glob_1 = pq1[0]  # for detection preview.replace("\\", "/")
    pq_glob_2 = pq2[0]  # for detection preview.replace("\\", "/")

    # ---------------------------
    # Step 2: Validate selected columns exist in both sides
    # ---------------------------
    step(25, "Step 2/5: Validating selected columns...")

    p1 = str(pq_glob_1).replace("\\", "/").replace("'", "''")
    p2 = str(pq_glob_2).replace("\\", "/").replace("'", "''")

    # Keep RAW parquet column names for SQL mapping
    cols1_raw = con.execute(f"SELECT * FROM parquet_scan('{p1}') LIMIT 1").df().columns.tolist()
    cols2_raw = con.execute(f"SELECT * FROM parquet_scan('{p2}') LIMIT 1").df().columns.tolist()

    # Normalized names only for UI / matching logic
    cols1 = safe_cols(cols1_raw)
    cols2 = safe_cols(cols2_raw)
    internal_cols = {"_SOURCE_FILE", "_SHEET", "_KEY"}
    common_cols = sorted(list((set(cols1) & set(cols2)) - internal_cols))

    # Use config-time selections (do NOT re-render widgets here)
    focus_col = reco_focus_col

    if compare_mode == "Compare common columns":
        compare_cols = [c for c in common_cols if c not in set(KEY_CANDIDATES_ACTIVE)]
    else:
        # Keep only columns that still exist in BOTH sides
        compare_cols = [c for c in compare_cols if c in common_cols]
    if not compare_cols:
        st.error("No valid compare columns found in BOTH files. Please go back to section (1) and select common columns.")
        st.stop()

    # numeric detection (sample)
    step(35, "Step 3/5: Detecting numeric columns (sample-based)...")
    numeric_cols = detect_numeric_cols(con, pq_glob_1, compare_cols, key_exclude=set(KEY_CANDIDATES_ACTIVE))

    # ---------------------------
    # Step 3: Aggregate by key per side
    # ---------------------------
    step(55, "Step 4/5: Aggregating duplicates by key and preparing joined views...")

    # Build SQL select list for agg
    def qident(name):
        return '"' + name.replace('"', '""') + '"'

    def parquet_scan_sql(files):
        if isinstance(files, str):
            files = [files]

        cleaned = []
        for f in files:
            f = str(f).replace('\\', '/')
            f = f.replace("'", "''")   # escape single quotes for DuckDB
            cleaned.append(f"'{f}'")

        arr = "[" + ",".join(cleaned) + "]"
        return f"parquet_scan({arr})"

    # numeric sum, text any_value, plus trace
    agg_select_parts = ["_KEY"]
    agg_group_by = "_KEY"

    # Map normalized compare cols -> actual parquet cols (use raw schema)
    _colmap1_cmp = {_norm_colname(c): c for c in cols1_raw}
    _colmap2_cmp = {_norm_colname(c): c for c in cols2_raw}

    for c in compare_cols:
        if c in KEY_CANDIDATES_ACTIVE:
            continue

        # Resolve normalized compare col -> actual parquet col
        real_c1 = _colmap1_cmp.get(_norm_colname(c), c)
        real_c2 = _colmap2_cmp.get(_norm_colname(c), c)

        # Prefer file1 real column; fallback to file2; else keep c
        real_c = real_c1 if real_c1 in cols1_raw else (real_c2 if real_c2 in cols2_raw else c)

        if c in numeric_cols:
            # Clean commas/spaces and cast safely so 1,234.00 == 1234
            agg_select_parts.append(f"SUM({duck_clean_num_expr(qident(real_c))}) AS {qident(c)}")
        else:
            agg_select_parts.append(f"ANY_VALUE({qident(real_c)}) AS {qident(c)}")
            
    # trace fields
    agg_select_parts.append("STRING_AGG(DISTINCT _SOURCE_FILE, '; ') AS _SOURCE_FILE")
    agg_select_parts.append("STRING_AGG(DISTINCT _SHEET, '; ') AS _SHEET")
    # ✅ Duplicate count per key (used later as DUP_F1 / DUP_F2)
    agg_select_parts.append("COUNT(*) AS _DUP_COUNT")

    agg_select = ",\n        ".join(agg_select_parts)

    # Build runtime _KEY using selected candidates (COALESCE with blank-as-null)
    if KEY_CANDIDATES_ACTIVE:

        def _clean_sql(col_sql: str) -> str:
            # Clean text in DuckDB safely (no regex, no \t/\n escapes)
            # Removes: NBSP, spaces, commas, tab, CR, LF, apostrophe, double-quote
            return (
                "upper("
                "replace(replace(replace(replace(replace(replace(replace(replace("
                f"trim(replace(cast({col_sql} as varchar), chr(160), ''))"
                ", ' ', ''), ',', ''), chr(9), ''), chr(13), ''), chr(10), ''),"
                " chr(39), ''), chr(34), ''), '\"', '')"
                ")"
            )

        _colmap1_key = {_norm_colname(c): c for c in cols1_raw}
        _colmap2_key = {_norm_colname(c): c for c in cols2_raw}

        def _key_coalesce_expr(colmap):
            parts = []
            for c in KEY_CANDIDATES_ACTIVE:
                real = colmap.get(_norm_colname(c), c)
                parts.append(f"NULLIF({_clean_sql(qident(real))}, '')")
            return "coalesce(" + ",".join(parts) + ")"

        _co1 = _key_coalesce_expr(_colmap1_key)
        _co2 = _key_coalesce_expr(_colmap2_key)

        key_expr_sql_1 = f"COALESCE({_co1}, '')"
        key_expr_sql_2 = f"COALESCE({_co2}, '')"

    else:
        key_expr_sql_1 = "''"
        key_expr_sql_2 = "''"
    
    # ---- Optional filter (applied to BOTH sides before agg/join) ----
    filter_clause_1 = ""
    filter_clause_2 = ""

    if filter_enabled and filter_values:
        _colmap1 = {_norm_colname(c): c for c in cols1_raw}
        _colmap2 = {_norm_colname(c): c for c in cols2_raw}

        # ✅ ALWAYS normalize filter_col for lookup
        _fkey = _norm_colname(filter_col)

        # ✅ Resolve to actual parquet column names per file
        _real1 = _colmap1.get(_fkey)
        _real2 = _colmap2.get(_fkey)

        _fcol1 = qident(_real1 if _real1 else filter_col)
        _fcol2 = qident(_real2 if _real2 else filter_col)

        _vals = []
        for _v in filter_values:
            _vv = str(_v).strip().upper().replace("'", "''")
            if _vv:
                _vals.append(f"'{_vv}'")

        if _vals:
            _fexpr1 = f"upper(trim(cast({_fcol1} as varchar))) IN ({', '.join(_vals)})"
            _fexpr2 = f"upper(trim(cast({_fcol2} as varchar))) IN ({', '.join(_vals)})"
            filter_clause_1 = f"WHERE {_fexpr1}"
            filter_clause_2 = f"WHERE {_fexpr2}"

    con.execute("DROP VIEW IF EXISTS f1_raw")
    con.execute("DROP VIEW IF EXISTS f2_raw")
    # overwrite/define _KEY at runtime (do NOT rely on cached parquet _KEY)
    p1 = str(pq_glob_1).replace("\\", "/").replace("'", "''")
    p2 = str(pq_glob_2).replace("\\", "/").replace("'", "''")

    # --- SAFE CREATE VIEW: some source files don't have _KEY, so EXCLUDE(_KEY) can fail ---
    try:
        con.execute(
            f"CREATE VIEW f1_raw AS "
            f"SELECT * EXCLUDE (_KEY), {key_expr_sql_1} AS _KEY "
            f"FROM {parquet_scan_sql(pq_files1)} {filter_clause_1}"
        )
    except Exception as e:
        # If _KEY doesn't exist in source, fallback to no-EXCLUDE
        if "EXCLUDE" in str(e) and "_KEY" in str(e):
            con.execute(
                f"CREATE VIEW f1_raw AS "
                f"SELECT *, {key_expr_sql_1} AS _KEY "
                f"FROM {parquet_scan_sql(pq_files1)} {filter_clause_1}"
            )
        else:
            raise
        
    # --- SAFE CREATE VIEW for file2 ---
    try:
        con.execute(
            f"CREATE VIEW f2_raw AS "
            f"SELECT * EXCLUDE (_KEY), {key_expr_sql_2} AS _KEY "
            f"FROM {parquet_scan_sql(pq_files2)} {filter_clause_2}"
        )
    except Exception as e:
        if "EXCLUDE" in str(e) and "_KEY" in str(e):
            con.execute(
                f"CREATE VIEW f2_raw AS "
                f"SELECT *, {key_expr_sql_2} AS _KEY "
                f"FROM {parquet_scan_sql(pq_files2)} {filter_clause_2}"
            )
        else:
            raise

    dups = con.execute("""
        SELECT _KEY, COUNT(*) AS c
        FROM f1_raw
        GROUP BY _KEY
        HAVING COUNT(*) > 1
        ORDER BY c DESC
        LIMIT 50
        """).df()
    st.write("Top duplicates in File1 by _KEY", dups)

    con.execute("DROP TABLE IF EXISTS f1_agg")
    con.execute("DROP TABLE IF EXISTS f2_agg")

    con.execute(f"""
    CREATE TABLE f1_agg AS
    SELECT
        {agg_select}
    FROM f1_raw
    WHERE _KEY IS NOT NULL AND _KEY <> ''
    GROUP BY {agg_group_by};
    """)

    con.execute(f"""
    CREATE TABLE f2_agg AS
    SELECT
        {agg_select}
    FROM f2_raw
    WHERE _KEY IS NOT NULL AND _KEY <> ''
    GROUP BY {agg_group_by};
    """)

    # ---------------------------
    # Step 4: Full outer join + mismatch tagging (based on focus column)
    # ---------------------------
    step(75, "Step 5/5: Full join + generating outputs...")

    # Prepare join select with suffixes
    def sel_cols(side_prefix):
        parts = []
        for c in compare_cols:
            if c in KEY_CANDIDATES_ACTIVE:
                continue
            parts.append(f"{side_prefix}.{qident(c)} AS {qident(c + '_' + side_prefix)}")
        return parts

    f1_cols = sel_cols("f1")
    f2_cols = sel_cols("f2")

    # trace
    trace = [
        "f1._SOURCE_FILE AS File1_Name",
        "f1._SHEET AS Sheet1",
        "f2._SOURCE_FILE AS File2_Name",
        "f2._SHEET AS Sheet2",
        "COALESCE(f1._KEY, f2._KEY) AS _KEY",
        "CASE WHEN f1._KEY IS NULL THEN 0 ELSE f1._DUP_COUNT END AS DUP_F1",
        "CASE WHEN f2._KEY IS NULL THEN 0 ELSE f2._DUP_COUNT END AS DUP_F2",
    ]

    join_select = ",\n        ".join(trace + f1_cols + f2_cols)


    # Focus compare (used only for Summary table 'Particular')
    f1_focus = f'"{focus_col}_f1"'
    f2_focus = f'"{focus_col}_f2"'

    # Numeric tolerance for numeric compares
    TOLERANCE = float(numeric_tolerance)

    # Build per-column STATUS_<col> and DIFF_<col> expressions
    status_exprs = []
    diff_exprs = []
    mismatch_terms = []

    for base_col in compare_cols:
        f1_alias = f'"{base_col}_f1"'
        f2_alias = f'"{base_col}_f2"'

        is_num = base_col in numeric_cols

        if is_num:
            mismatch_expr_col = f"ABS(COALESCE({f1_alias},0) - COALESCE({f2_alias},0)) > {TOLERANCE}"
            diff_expr_col = f"(COALESCE({f1_alias},0) - COALESCE({f2_alias},0)) AS {qident('DIFF_' + base_col)}"
        else:
            
            # ✅ normalize text: trim + remove NBSP/spaces/tabs/CR/LF/commas + remove quotes/apostrophes + upper
            t1 = (
                "upper("
                "replace(replace(replace(replace(replace(replace(replace("
                "trim(replace(cast(" + f1_alias + " as varchar), chr(160), '')),"
                " ' ', ''), '\\t', ''), '\\r', ''), '\\n', ''), ',', ''), chr(39), ''), chr(34), '')"
                ")"
            )
            t2 = (
                "upper("
                "replace(replace(replace(replace(replace(replace(replace("
                "trim(replace(cast(" + f2_alias + " as varchar), chr(160), '')),"
                " ' ', ''), '\\t', ''), '\\r', ''), '\\n', ''), ',', ''), chr(39), ''), chr(34), '')"
                ")"
            )
            mismatch_expr_col = f"COALESCE({t1}, '') <> COALESCE({t2}, '')"
            diff_expr_col = f"NULL AS {qident('DIFF_' + base_col)}"

        if treat_blanks_as_equal:
            mismatch_expr_col = f"({mismatch_expr_col}) AND NOT ({f1_alias} IS NULL AND {f2_alias} IS NULL)"

        status_expr_col = f"""CASE
          WHEN f1._KEY IS NULL THEN 'ONLY_IN_FILE2'
          WHEN f2._KEY IS NULL THEN 'ONLY_IN_FILE1'
          WHEN {mismatch_expr_col} THEN 'MISMATCH'
          ELSE 'MATCH'
        END AS {qident('STATUS_' + base_col)}"""

        status_exprs.append(status_expr_col)
        diff_exprs.append(diff_expr_col)
        mismatch_terms.append(mismatch_expr_col)

    any_mismatch_expr = " OR ".join([f"({t})" for t in mismatch_terms]) if mismatch_terms else "FALSE"

    final_status_expr = f"""CASE
          WHEN f1._KEY IS NULL THEN 'ONLY_IN_FILE2'
          WHEN f2._KEY IS NULL THEN 'ONLY_IN_FILE1'
          WHEN {any_mismatch_expr} THEN 'MISMATCH'
          ELSE 'MATCH'
        END AS FINAL_STATUS"""

    # Extend select list: raw cols + per-column status/diff + FINAL_STATUS
    # Interleave per selected column: <col>_f1, <col>_f2, STATUS_<col>, DIFF_<col>
    compare_select = []
    for i, _c in enumerate(compare_cols):
        if i < len(f1_cols):
            compare_select.append(f1_cols[i])
        if i < len(f2_cols):
            compare_select.append(f2_cols[i])
        if i < len(status_exprs):
            compare_select.append(status_exprs[i])
        if i < len(diff_exprs):
            compare_select.append(diff_exprs[i])


    join_select = ",\n".join(trace + compare_select + [final_status_expr])

    con.execute("DROP VIEW IF EXISTS joined_all")
    con.execute(f"""
    CREATE VIEW joined_all AS
    SELECT
        {join_select}
    FROM f1_agg f1
    FULL OUTER JOIN f2_agg f2
    ON f1._KEY = f2._KEY;
    """)

    # Output paths
    out_summary_csv = outdir / "summary.csv"
    out_matched_pq = outdir / "matched.parquet"
    out_mism_pq = outdir / "mismatched.parquet"
    out_only1_pq = outdir / "only_in_file1.parquet"
    out_only2_pq = outdir / "only_in_file2.parquet"
    out_dups_pq  = outdir / "duplicates.parquet"

    # Write Parquet outputs
    con.execute(f"COPY (SELECT * FROM joined_all WHERE FINAL_STATUS='MATCH') TO '{out_matched_pq.as_posix()}' (FORMAT PARQUET)")
    con.execute(f"COPY (SELECT * FROM joined_all WHERE FINAL_STATUS='MISMATCH') TO '{out_mism_pq.as_posix()}' (FORMAT PARQUET)")
    con.execute(f"COPY (SELECT * FROM joined_all WHERE FINAL_STATUS='ONLY_IN_FILE1') TO '{out_only1_pq.as_posix()}' (FORMAT PARQUET)")
    con.execute(f"COPY (SELECT * FROM joined_all WHERE FINAL_STATUS='ONLY_IN_FILE2') TO '{out_only2_pq.as_posix()}' (FORMAT PARQUET)")
    con.execute(f"COPY (SELECT * FROM joined_all WHERE DUP_F1 > 1 OR DUP_F2 > 1) TO '{out_dups_pq.as_posix()}' (FORMAT PARQUET)")
    
    # Summary
    matched_cnt = con.execute("SELECT COUNT(*) FROM joined_all WHERE FINAL_STATUS='MATCH'").fetchone()[0]
    mism_cnt = con.execute("SELECT COUNT(*) FROM joined_all WHERE FINAL_STATUS='MISMATCH'").fetchone()[0]
    only1_cnt = con.execute("SELECT COUNT(*) FROM joined_all WHERE FINAL_STATUS='ONLY_IN_FILE1'").fetchone()[0]
    only2_cnt = con.execute("SELECT COUNT(*) FROM joined_all WHERE FINAL_STATUS='ONLY_IN_FILE2'").fetchone()[0]
    dup_keys_f1 = con.execute("SELECT COUNT(*) FROM f1_agg WHERE _DUP_COUNT > 1").fetchone()[0]
    dup_keys_f2 = con.execute("SELECT COUNT(*) FROM f2_agg WHERE _DUP_COUNT > 1").fetchone()[0]
    dup_rows    = con.execute("SELECT COUNT(*) FROM joined_all WHERE DUP_F1 > 1 OR DUP_F2 > 1").fetchone()[0]

    
    # -----------------------------
    # Focus sums (robust: works even if focus cols are text)
    # -----------------------------
    def _sum_focus(where_sql: str, expr_sql: str) -> float:
        # TRY_CAST handles text numbers safely; non-convertible becomes NULL -> treated as 0
        q = f"""
            SELECT COALESCE(SUM(COALESCE(TRY_CAST({expr_sql} AS DOUBLE), 0)), 0)
            FROM joined_all
            WHERE {where_sql}
        """
        return float(con.execute(q).fetchone()[0] or 0)

    # Build per-bucket focus sums
    s_match_f1 = _sum_focus("FINAL_STATUS='MATCH'", f1_focus)
    s_match_f2 = _sum_focus("FINAL_STATUS='MATCH'", f2_focus)

    s_mism_f1  = _sum_focus("FINAL_STATUS='MISMATCH'", f1_focus)
    s_mism_f2  = _sum_focus("FINAL_STATUS='MISMATCH'", f2_focus)

    s_only1_f1 = _sum_focus("FINAL_STATUS='ONLY_IN_FILE1'", f1_focus)
    s_only1_f2 = 0.0

    s_only2_f1 = 0.0
    s_only2_f2 = _sum_focus("FINAL_STATUS='ONLY_IN_FILE2'", f2_focus)

    s_dups_f1  = _sum_focus("(DUP_F1 > 1 OR DUP_F2 > 1)", f1_focus)
    s_dups_f2  = _sum_focus("(DUP_F1 > 1 OR DUP_F2 > 1)", f2_focus)

    # -----------------------------
    # Summary table (add Diff + TOTAL)
    # -----------------------------
    summary_rows = [
        {"SNO": 1, "Particular": focus_col, "Count": int(matched_cnt), "Focus_f1_Sum": s_match_f1, "Focus_f2_Sum": s_match_f2, "Diff": s_match_f1 - s_match_f2, "Remarks": "Match"},
        {"SNO": 2, "Particular": focus_col, "Count": int(mism_cnt),    "Focus_f1_Sum": s_mism_f1,  "Focus_f2_Sum": s_mism_f2,  "Diff": s_mism_f1  - s_mism_f2,  "Remarks": "Mismatch"},
        {"SNO": 3, "Particular": focus_col, "Count": int(only1_cnt),   "Focus_f1_Sum": s_only1_f1, "Focus_f2_Sum": s_only1_f2, "Diff": s_only1_f1 - s_only1_f2, "Remarks": "Only in File1"},
        {"SNO": 4, "Particular": focus_col, "Count": int(only2_cnt),   "Focus_f1_Sum": s_only2_f1, "Focus_f2_Sum": s_only2_f2, "Diff": s_only2_f1 - s_only2_f2, "Remarks": "Only in File2"},
        {"SNO": 5, "Particular": focus_col, "Count": int(dup_rows),    "Focus_f1_Sum": s_dups_f1,  "Focus_f2_Sum": s_dups_f2,  "Diff": s_dups_f1  - s_dups_f2,  "Remarks": f"Duplicates (keys: F1={int(dup_keys_f1)}, F2={int(dup_keys_f2)})"},
    ]

    # TOTAL row (IMPORTANT: do NOT include Duplicates bucket in totals)
    _non_dup_rows = [r for r in summary_rows if not str(r.get("Remarks", "")).startswith("Duplicates")]

    total_count = sum(r["Count"] for r in _non_dup_rows)
    total_f1 = sum(r["Focus_f1_Sum"] for r in _non_dup_rows)
    total_f2 = sum(r["Focus_f2_Sum"] for r in _non_dup_rows)

    summary_rows.append({
        "SNO": 6,
        "Particular": focus_col,
        "Count": int(total_count),
        "Focus_f1_Sum": float(total_f1),
        "Focus_f2_Sum": float(total_f2),
        "Diff": float(total_f1 - total_f2),
        "Remarks": "TOTAL",
    })

    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_csv(out_summary_csv, index=False)
    out_summary_pdf = Path(outdir) / "summary_report.pdf"
    build_summary_pdf(out_summary_pdf, summary_df, APP_VERSION, focus_col, Path(outdir))

    # Save config + log
    cfg = {
        "run_id": outdir.name,
        "file1_sig": sig1,
        "file2_sig": sig2,
        "compare_mode": compare_mode,
        "compare_cols": compare_cols,
        "focus_col": focus_col,
        "numeric_tolerance": TOLERANCE,
        "filter_enabled": bool(filter_enabled),
        "filter_column": filter_col if filter_enabled else None,
        "filter_values": filter_values if filter_enabled else None,
        "treat_blanks_as_equal": bool(treat_blanks_as_equal),
        "numeric_cols_detected": numeric_cols,
        "generated": {
            "summary": str(out_summary_csv),
            "matched_parquet": str(out_matched_pq),
            "mismatched_parquet": str(out_mism_pq),
            "only_in_file1_parquet": str(out_only1_pq),
            "only_in_file2_parquet": str(out_only2_pq),
        }
    }
    write_json(outdir / "config.json", cfg)
    (outdir / "run.log").write_text("\n".join(log_lines), encoding="utf-8")

    # Cache run state for UI
    st.session_state["run_cache_sig"] = config_sig
    st.session_state["run_cache_outdir"] = str(outdir)

    prog.progress(100)
    status.success("✅ Completed")

    # ---------------------------
    # Display results
    # ---------------------------
    st.subheader("3) Summary")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Matched", f"{matched_cnt:,}")
    k2.metric("Mismatched", f"{mism_cnt:,}")
    k3.metric("Only in File1", f"{only1_cnt:,}")
    k4.metric("Only in File2", f"{only2_cnt:,}")
    st.dataframe(summary_df, use_container_width=True)

    st.subheader("4) Preview (Top rows)")

    tabs = st.tabs([
        "✅ Matched",
        "❌ Mismatched",
        "⬅ Only in File1",
        "➡ Only in File2",
        "🟡 Duplicates",
        "🧾 Run Log"
    ])

    def preview_parquet(path: Path, n=2000):
        if not path.exists():
            return pd.DataFrame()
        return con.execute(
            f"SELECT * FROM parquet_scan('{path.as_posix()}') LIMIT {int(n)}"
        ).df()

    with tabs[0]:
        st.dataframe(preview_parquet(out_matched_pq, show_rows), use_container_width=True)

    with tabs[1]:
        st.dataframe(preview_parquet(out_mism_pq, show_rows), use_container_width=True)

    with tabs[2]:
        st.dataframe(preview_parquet(out_only1_pq, show_rows), use_container_width=True)

    with tabs[3]:
        st.dataframe(preview_parquet(out_only2_pq, show_rows), use_container_width=True)

    with tabs[4]:
        st.dataframe(preview_parquet(out_dups_pq, show_rows), use_container_width=True)

    with tabs[5]:
        st.code(
            (outdir / "run.log").read_text(encoding="utf-8")
            if (outdir / "run.log").exists()
            else "",
            language="text"
        )



    # ===============================
    # 5) Downloads Section
    # ===============================
    st.subheader("5) Downloads")

    summary_bytes = out_summary_csv.read_bytes()
    summary_pdf_bytes = out_summary_pdf.read_bytes() if out_summary_pdf.exists() else b""
    # Save summary physically to run folder
    summary_path = Path(outdir) / "summary.csv"
    with open(summary_path, "wb") as f:
        f.write(summary_bytes)
    st.download_button(
        "⬇ Download Summary (CSV)",
        data=summary_bytes,
        file_name="summary.csv",
        mime="text/csv"
    )
    # ✅ PDF summary download
    if summary_pdf_bytes:
        st.download_button(
            "⬇ Download Summary Report (PDF)",
            data=summary_pdf_bytes,
            file_name="summary_report.pdf",
            mime="application/pdf"
        )

    # ✅ Open a fresh DuckDB connection ONLY for exports
    export_sig = str(outdir)

    if st.session_state["export_cache_sig"] == export_sig and st.session_state["export_cache"] is not None:
        cached = st.session_state["export_cache"]
        matched_files = cached["matched_files"]
        mism_files    = cached["mism_files"]
        only1_files   = cached["only1_files"]
        only2_files   = cached["only2_files"]
        dups_files    = cached.get("dups_files", {})   # ✅ ADD
        rows_table    = cached["rows_table"]
        zip_bytes     = cached["zip_bytes"]
        summary_bytes = cached["summary_bytes"]
    else:
        # ✅ generate exports once
        con_dl = duckdb.connect(database=":memory:")
        try:
            # ---- Export progress UI ----
            st.markdown("### ⏳ Export Progress")
            export_prog = st.progress(0)
            export_msg = st.empty()

            rows_table = []

            def export_one(label: str, pq_path: Path, pct_start: int, pct_end: int):
                export_msg.info(f"Exporting: {label} ...")
                export_prog.progress(pct_start)

                files, meta = parquet_to_csv_bytes(con_dl, pq_path, max_rows=800000)

                # save CSVs physically to run folder
                for fname, data in files.items():
                    (Path(outdir) / fname).write_bytes(data)

                # rows table
                if meta["parts"] == 1:
                    for fname in files.keys():
                        rows_table.append({"Category": label, "File": fname, "Rows": meta["total_rows"]})
                else:
                    for i, r in enumerate(meta["part_rows"], start=1):
                        fname = f"{meta['base_stem']}_part{i}.csv"
                        if fname in files:
                            rows_table.append({"Category": label, "File": fname, "Rows": int(r)})

                export_prog.progress(pct_end)
                return files

            matched_files = export_one("Matched", out_matched_pq, 5, 30)
            mism_files    = export_one("Mismatched", out_mism_pq, 30, 55)
            only1_files   = export_one("Only in File1", out_only1_pq, 55, 80)
            only2_files   = export_one("Only in File2", out_only2_pq, 80, 95)
            dups_files    = export_one("Duplicates", out_dups_pq, 95, 99)

            export_msg.success("✅ Export ready")
            export_prog.progress(100)

        finally:
            con_dl.close()

        # summary bytes (and save physically too)
        summary_bytes = out_summary_csv.read_bytes()
        (Path(outdir) / "summary.csv").write_bytes(summary_bytes)

        # zip bytes
        
        zip_payload = {
            "summary.csv": summary_bytes,
            "summary_report.pdf": summary_pdf_bytes
        }
        zip_payload.update(matched_files)
        zip_payload.update(mism_files)
        zip_payload.update(only1_files)
        zip_payload.update(only2_files)
        zip_payload.update(dups_files)

        zip_bytes = build_zip_bytes(zip_payload)

        # ✅ cache everything
        st.session_state["export_cache_sig"] = export_sig
        st.session_state["export_cache"] = {
        "matched_files": matched_files,
        "mism_files": mism_files,
        "only1_files": only1_files,
        "only2_files": only2_files,
        "dups_files": dups_files,        # ✅ ADD
        "rows_table": rows_table,
        "zip_bytes": zip_bytes,
        "summary_bytes": summary_bytes,
    }

    # ---- Rows count table ----
    st.markdown("### 📊 Rows count per output file")
    if rows_table:
        st.dataframe(pd.DataFrame(rows_table), use_container_width=True)


    # ---- Build category summary from rows_table + file size ----
    from collections import defaultdict

    category_summary = defaultdict(lambda: {"files": 0, "rows": 0, "size_bytes": 0})

    def accumulate_summary(category_name, files_dict):
        if not files_dict:
            return
        for fname, data in files_dict.items():
            category_summary[category_name]["files"] += 1
            category_summary[category_name]["size_bytes"] += len(data)

    for r in rows_table:
        cat = r["Category"]
        category_summary[cat]["rows"] += int(r["Rows"])

    accumulate_summary("Matched", matched_files)
    accumulate_summary("Mismatched", mism_files)
    accumulate_summary("Only in File1", only1_files)
    accumulate_summary("Only in File2", only2_files)
    accumulate_summary("Duplicates", dups_files)   # ✅ ADD

    # ---- Total Export Size badge ----
    total_bytes = 0
    for v in category_summary.values():
        total_bytes += int(v.get("size_bytes", 0))

    total_mb = total_bytes / (1024 * 1024)

    st.markdown(
        f"""
        <div style="
            display:inline-block;
            background:#0d6efd;
            color:white;
            padding:6px 12px;
            border-radius:10px;
            font-size:13px;
            font-weight:700;
            margin-bottom:10px;
        ">
            📦 Total download size: {total_mb:,.2f} MB
        </div>
        """,
        unsafe_allow_html=True
    )

    def summary_label(category_name: str, color: str):
        data = category_summary.get(category_name, {"files": 0, "rows": 0, "size_bytes": 0})
        size_mb = data["size_bytes"] / (1024 * 1024)

        badge = f"""
        <span style="
            background-color:{color};
            color:white;
            padding:3px 8px;
            border-radius:8px;
            font-size:12px;
            font-weight:600;
        ">
            {data['files']} file(s)
        </span>
        """

        return f"""{badge} &nbsp; {data['rows']:,} rows • {size_mb:,.2f} MB"""

    # ---- Download buttons (grouped) ----
    
    st.markdown("### 📥 Download Outputs (Grouped)")

    with st.expander("✅ Matched Files", expanded=False):
        st.markdown(summary_label("Matched", "#28a745"), unsafe_allow_html=True)
        if not matched_files:
            st.info("No matched files generated.")
        else:
            for name, data in matched_files.items():
                st.download_button(f"⬇ Download {name}", data=data, file_name=name, mime="text/csv")

    with st.expander("❌ Mismatched Files", expanded=True):
        st.markdown(summary_label("Mismatched", "#dc3545"), unsafe_allow_html=True)
        if not mism_files:
            st.info("No mismatched files generated.")
        else:
            for name, data in mism_files.items():
                st.download_button(f"⬇ Download {name}", data=data, file_name=name, mime="text/csv")

    with st.expander("⬅ Only in File1", expanded=False):
        st.markdown(summary_label("Only in File1", "#fd7e14"), unsafe_allow_html=True)
        if not only1_files:
            st.info("No only-in-file1 files generated.")
        else:
            for name, data in only1_files.items():
                st.download_button(f"⬇ Download {name}", data=data, file_name=name, mime="text/csv")

    with st.expander("➡ Only in File2", expanded=False):
        st.markdown(summary_label("Only in File2", "#fd7e14"), unsafe_allow_html=True)
        if not only2_files:
            st.info("No only-in-file2 files generated.")
    
        else:
            for name, data in only2_files.items():
                st.download_button(f"⬇ Download {name}", data=data, file_name=name, mime="text/csv")
    # ✅ ADD THIS (Duplicates expander)
    with st.expander("🟡 Duplicates", expanded=False):
        st.markdown(summary_label("Duplicates", "#ffc107"), unsafe_allow_html=True)
        if not dups_files:
            st.info("No duplicate files generated.")
        else:
            for name, data in dups_files.items():
                st.download_button(f"⬇ Download {name}", data=data, file_name=name, mime="text/csv")
    # ---- ZIP ----
    zip_payload = {
    "summary.csv": summary_bytes,
    "summary_report.pdf": summary_pdf_bytes
    }
    zip_payload.update(matched_files)
    zip_payload.update(mism_files)
    zip_payload.update(only1_files)
    zip_payload.update(only2_files)
    zip_payload.update(dups_files)
    zip_bytes = build_zip_bytes(zip_payload)
    st.download_button(
        "⬇ Download ALL (ZIP)",
        data=zip_bytes,
        file_name="reco_outputs.zip",
        mime="application/zip"
    )

    st.caption(f"Run folder saved at: {outdir}")
    colA, colB = st.columns([1, 3])

    with colA:
        if st.button("📂 Open run folder"):
            p = str(Path(outdir).resolve())
            try:
                os.startfile(p)  # Windows
                st.success("Opened run folder.")
            except Exception:
                try:
                    subprocess.Popen(["explorer", p])
                    st.success("Opened run folder.")
                except Exception as e:
                    st.error(f"Could not open automatically. Open manually:\n{p}\nError: {e}")


finally:
    con.close()
