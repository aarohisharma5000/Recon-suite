import streamlit as st
from pathlib import Path
from datetime import datetime
import importlib.util
import textwrap
import os, sys
import platform

from pathlib import Path

def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent

BASE_DIR = app_base_dir()

st.set_page_config(page_title="Reco Suite Launcher", layout="wide")

BASE   = BASE_DIR
ASSETS = BASE / "assets"
LOGO_PATH = ASSETS / "logo.png"

DATA_DIR  = BASE / "data"
INPUT_DIR = DATA_DIR / "input"
FILE1_DIR = INPUT_DIR / "file1"
FILE2_DIR = INPUT_DIR / "file2"
CACHE_DIR = DATA_DIR / "cache"
RUNS_DIR  = DATA_DIR / "runs"

for p in [ASSETS, DATA_DIR, INPUT_DIR, FILE1_DIR, FILE2_DIR, CACHE_DIR, RUNS_DIR]:
    p.mkdir(parents=True, exist_ok=True)

LAUNCHER_VERSION = "v1.1"

def html(s: str):
    st.markdown(textwrap.dedent(s).strip(), unsafe_allow_html=True)

def fmt_ts(dt):
    if not dt:
        return "—"
    return dt.strftime("%d-%b-%Y %I:%M %p")

def check_deps(mods):
    missing = [m for m in mods if importlib.util.find_spec(m) is None]
    return (len(missing) == 0), missing

# ---------- session state ----------
if "selected_tool"    not in st.session_state: st.session_state["selected_tool"]    = "tool1"
if "launched"         not in st.session_state: st.session_state["launched"]         = False
if "last_launched_at" not in st.session_state: st.session_state["last_launched_at"] = {"tool1": None, "tool2": None}
if "dark_mode"        not in st.session_state: st.session_state["dark_mode"]        = False

# ---------- query params ----------
try:
    qp      = dict(st.query_params)
    qp_tool = qp.get("tool", "")
    qp_auto = qp.get("auto", "0")
except Exception:
    qp      = st.experimental_get_query_params()
    qp_tool = (qp.get("tool", [""])[0])
    qp_auto = (qp.get("auto", ["0"])[0])

# ---------- Tool registry ----------
TOOLS = {
    "tool1": {
        "name": "Tool 1",
        "desc": "Recon Tool 1 (app.py)",
        "icon": "📄",
        "version": "v1.6",
        "path": BASE / "app.py",
        "deps": ["streamlit", "pandas", "openpyxl"],
        "tag": "Classic",
        "tag_color": "blue",
        "features": [
            "✅ Smart auto-detect for ANY column name",
            "✅ COALESCE key — multi-column fallback",
            "✅ Duplicate Include / Exclude mode",
            "✅ Before vs After duplicate comparison",
            "✅ Search any Loan Account Number",
            "✅ Filter by Lender / Subset",
            "✅ Numeric tolerance setting",
            "✅ Per-column Match / Diff analysis",
            "✅ PDF Summary Report",
            "✅ Excel + CSV + ZIP downloads",
            "✅ Multi-file upload (CSV/XLSX/XLS/XLSB)",
        ],
        "best_for": "Standard file comparison — upload and run",
        "feature_count": "11 Features",
    },
    "tool2": {
        "name": "Tool 2",
        "desc": "Recon Tool 2 (app_v2.py)",
        "icon": "🧠",
        "version": "v2.1",
        "path": BASE / "app_v2.py",
        "deps": ["streamlit", "pandas", "duckdb", "pyarrow"],
        "tag": "Big File",
        "tag_color": "purple",
        "features": [
            "✅ DuckDB engine — disk-based SQL",
            "✅ Parquet caching — skip reconvert",
            "✅ Handles files too large for RAM",
            "✅ Duplicate preview + export",
            "✅ ZIP downloads + summaries",
            "✅ Filter applied at SQL level",
            "✅ Auto CSV split for 800k+ rows",
            "✅ Run folder saved to disk",
            "✅ Per-column STATUS + DIFF output",
        ],
        "best_for": "Large datasets / faster DuckDB processing",
        "feature_count": "9 Features",
    },
}

if qp_tool in TOOLS and str(qp_auto) == "1":
    st.session_state["selected_tool"] = qp_tool
    st.session_state["launched"]      = True

selected = st.session_state["selected_tool"]

# ---------- dark mode toggle ----------
toggle_col_l, toggle_col_r = st.columns([0.85, 0.15], vertical_alignment="center")
with toggle_col_r:
    st.session_state["dark_mode"] = st.toggle("🌙 Dark", value=st.session_state["dark_mode"])

is_dark = st.session_state["dark_mode"]

# ---------- CSS ----------
html(f"""
<style>
  header[data-testid="stHeader"] {{ height:0rem!important; visibility:hidden!important; }}
  footer {{ visibility:hidden; }}

  :root {{
    --bg:     {"#0b1220"             if is_dark else "#f4f7fb"};
    --panel:  {"#0f172a"             if is_dark else "#ffffff"};
    --panel2: {"#0b1326"             if is_dark else "#f0f6ff"};
    --text:   {"#e5e7eb"             if is_dark else "#111827"};
    --muted:  {"#a1a1aa"             if is_dark else "#6b7280"};
    --border: {"rgba(148,163,184,0.25)" if is_dark else "#dde6f5"};
    --brand:  #0d6efd;
    --good:   #198754;
    --bad:    #dc3545;
    --purple: #7c3aed;
    --shadow: {"rgba(0,0,0,0.35)"    if is_dark else "rgba(13,110,253,0.10)"};
  }}

  .stApp {{ background: var(--bg); }}

  /* ── Hero ── */
  .hero {{
    border:1px solid var(--border);
    border-radius:22px;
    padding:22px 28px;
    background:
      radial-gradient(800px 260px at 10% 0%,  rgba(13,110,253,0.18), transparent 60%),
      radial-gradient(800px 260px at 90% 10%, rgba(124,58,237,0.14), transparent 55%),
      var(--panel);
    box-shadow:0 12px 40px var(--shadow);
    margin-bottom:12px;
    text-align:center;
  }}
  .heroTitle {{
    font-size:38px; font-weight:900; letter-spacing:-0.5px;
    color:var(--text); margin:0; line-height:1.1;
  }}
  .heroVer {{
    font-size:12px; font-weight:800; padding:3px 10px;
    border-radius:999px; border:1px solid var(--border);
    color:var(--muted); display:inline-block; margin-left:8px;
    vertical-align:middle; background:rgba(255,255,255,0.04);
  }}
  .heroSub {{
    color:var(--muted); font-size:13px; margin-top:6px;
  }}

  /* ── Section headers ── */
  .sec-title {{
    font-size:20px; font-weight:900; color:var(--text);
    margin:0 0 14px 0; display:flex; align-items:center; gap:8px;
  }}

  /* ── Tool Card — COMPACT ── */
  .tool-card {{
    border:1.5px solid var(--border);
    border-radius:16px;
    padding:14px 16px;
    background:var(--panel);
    transition:transform 160ms ease, box-shadow 160ms ease, border-color 160ms ease;
  }}
  .tool-card:hover {{
    transform:translateY(-2px);
    box-shadow:0 8px 24px var(--shadow);
  }}
  .tool-card.selected {{
    border:2px solid var(--brand);
    background:var(--panel2);
    box-shadow:0 6px 20px var(--shadow);
  }}
  .tc-header {{
    display:flex; align-items:center; gap:10px; margin-bottom:8px;
  }}
  .tc-icon {{
    width:38px; height:38px; border-radius:10px;
    display:flex; align-items:center; justify-content:center;
    font-size:20px;
    background:{"#1e293b" if is_dark else "#eef3ff"};
    flex:0 0 38px;
  }}
  .tc-title {{ font-size:16px; font-weight:900; color:var(--text); margin:0; }}
  .tc-sub   {{ font-size:11px; color:var(--muted); margin:1px 0 0 0; }}
  .tc-divider {{
    height:1px; background:var(--border); margin:8px 0;
  }}
  .tc-bestfor {{
    font-size:11px; color:var(--muted);
    margin-bottom:6px;
    padding:4px 8px;
    border-radius:6px;
    background:{"rgba(255,255,255,0.03)" if is_dark else "rgba(13,110,253,0.05)"};
    border:1px solid var(--border);
  }}
  .tc-features {{
    font-size:12px; line-height:1.75; color:var(--text);
    margin:0; padding:0; list-style:none;
  }}

  /* ── Badges ── */
  .badge {{
    display:inline-block; padding:2px 8px;
    border-radius:999px; font-size:10px; font-weight:800;
    border:1px solid var(--border);
    background:rgba(255,255,255,0.04);
    color:var(--text); white-space:nowrap;
  }}
  .badge.blue   {{ background:rgba(13,110,253,0.16);  border-color:rgba(13,110,253,0.28);  color:{"#cfe0ff" if is_dark else "#0b5ed7"}; }}
  .badge.green  {{ background:rgba(25,135,84,0.18);   border-color:rgba(25,135,84,0.28);   color:{"#c7f9d6" if is_dark else "#0f5132"}; }}
  .badge.red    {{ background:rgba(220,53,69,0.18);   border-color:rgba(220,53,69,0.30);   color:{"#ffd0d6" if is_dark else "#842029"}; }}
  .badge.purple {{ background:rgba(124,58,237,0.16);  border-color:rgba(124,58,237,0.28);  color:{"#ddd6fe" if is_dark else "#5b21b6"}; }}
  .badge.gray   {{ background:rgba(255,255,255,0.04); }}
  .badge.orange {{ background:rgba(253,126,20,0.16);  border-color:rgba(253,126,20,0.28);  color:{"#fed8aa" if is_dark else "#7c3500"}; }}
  .badges {{ display:flex; gap:4px; flex-wrap:wrap; margin-bottom:8px; }}

  /* ── Capability card ── */
  .cap-card {{
    border:1px solid var(--border); border-radius:12px;
    padding:12px 14px; background:var(--panel); margin-bottom:8px;
  }}
  .cap-title {{
    font-size:13px; font-weight:900; color:var(--brand);
    margin-bottom:6px;
  }}
  .cap-item {{
    font-size:12px; color:var(--text);
    line-height:1.6; display:flex; align-items:flex-start; gap:5px;
  }}

  /* ── Link box ── */
  .linkbox {{
    border:1px solid var(--border); border-radius:16px;
    padding:16px; background:var(--panel);
  }}
  .linktitle {{ font-weight:900; margin-bottom:10px; color:var(--text); font-size:15px; }}
  .hint {{ color:var(--muted); font-size:12px; margin-top:10px; }}
  .linkbox a {{ text-decoration:none; font-weight:800; color:var(--brand); }}
  .link-row {{
    display:flex; align-items:center; gap:8px;
    padding:8px 0; border-bottom:1px solid var(--border);
  }}
  .link-row:last-of-type {{ border-bottom:none; }}

  /* ── Action bar ── */
  .selline {{
    display:flex; gap:10px; flex-wrap:wrap;
    align-items:center; margin:6px 0 16px 0;
  }}

  .divider {{ height:1px; background:var(--border); margin:20px 0; }}

  /* ── Help section ── */
  .help-card {{
    border:1px solid var(--border); border-radius:14px;
    padding:16px 18px; background:var(--panel); margin-bottom:10px;
  }}
  .help-title {{
    font-size:16px; font-weight:800; color:var(--brand); margin-bottom:6px;
  }}
  .help-body {{
    font-size:13px; color:var(--text); line-height:1.8;
  }}
</style>
""")

# ═══════════════════════════════════════════
# HERO
# ═══════════════════════════════════════════
hero_l, hero_m, hero_r = st.columns([0.15, 0.70, 0.15], vertical_alignment="center")
with hero_l:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=120)
with hero_m:
    html(f"""
    <div class="hero">
      <div class="heroTitle">
        📌 Reconciliation Suite
        <span class="heroVer">Launcher {LAUNCHER_VERSION}</span>
      </div>
      <div class="heroSub">
        Select a tool → Launch. Use "Open in new tab" to run both tools simultaneously.
      </div>
    </div>
    """)
with hero_r:
    html(f"""
    <div style="text-align:right;">
      <span class="badge blue">☁️ Cloud</span>
    </div>
    """)

# ═══════════════════════════════════════════
# SYSTEM STATUS
# ═══════════════════════════════════════════
st.markdown("### ✅ System Status")
s1, s2, s3, s4 = st.columns(4)
s1.metric("Python", platform.python_version())
s2.metric("OS",     platform.system())
s3.metric("Base",   BASE_DIR.name)
s4.metric("Mode",   "Streamlit Cloud")

with st.expander("📌 Environment Details", expanded=False):
    st.code(f"""BASE_DIR:    {BASE_DIR}
Executable:  {sys.executable}
Frozen:      {getattr(sys, 'frozen', False)}
""".strip(), language="text")

html('<div class="divider"></div>')

# ═══════════════════════════════════════════
# SUITE OVERVIEW METRICS
# ═══════════════════════════════════════════
st.markdown("### 🏠 Suite Overview")
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Tools Available",    "2")
k2.metric("Supported Formats",  "CSV / XLSX / XLS / XLSB")
k3.metric("Engine",             "Pandas + DuckDB")
k4.metric("Runs On",            "Local Machine")
k5.metric("Data Privacy",       "100% Local")

html('<div class="divider"></div>')

# ═══════════════════════════════════════════
# MAIN LAYOUT — Tool Cards + Side Panel
# ═══════════════════════════════════════════
st.markdown("### 🚀 Available Tools")

card_col1, card_col2, side_col = st.columns([1, 1, 0.65], gap="large")

# ── Tool cards — COMPACT render_card ──
def render_card(tool_key: str, col):
    t       = TOOLS[tool_key]
    is_sel  = (selected == tool_key)
    last_ts = st.session_state["last_launched_at"].get(tool_key)
    ok, missing = check_deps(t.get("deps", []))

    health_badge = (
        "<span class='badge green'>✅ deps OK</span>"
        if ok else
        f"<span class='badge red'>❌ missing: {', '.join(missing)}</span>"
    )
    sel_badge  = "<span class='badge green'>Selected</span>" if is_sel else ""
    tag_color  = t.get("tag_color", "blue")
    tag_label  = t.get("tag", "")
    feat_count = t.get("feature_count", "")

    # Show first 5 features; rest hidden under collapsible
    all_feats  = t["features"]
    top_feats  = all_feats[:5]
    more_feats = all_feats[5:]

    top_html  = "".join(f"<li class='tc-features' style='list-style:none;font-size:12px;line-height:1.75;color:var(--text);'>{f}</li>" for f in top_feats)
    more_html = "".join(f"<li style='list-style:none;font-size:12px;line-height:1.75;color:var(--text);'>{f}</li>" for f in more_feats)
    more_block = (
        f"<details style='margin-top:4px;'>"
        f"<summary style='font-size:11px;color:var(--muted);cursor:pointer;list-style:none;'>+ {len(more_feats)} more features</summary>"
        f"<ul style='padding:0;margin:4px 0 0 0;'>{more_html}</ul>"
        f"</details>"
    ) if more_feats else ""

    card_cls = "tool-card selected" if is_sel else "tool-card"

    card_html = f"""
    <div class="{card_cls}">
      <div class="tc-header">
        <div class="tc-icon">{t['icon']}</div>
        <div>
          <div class="tc-title">{t['name']}</div>
          <div class="tc-sub">{t['desc']}</div>
        </div>
      </div>
      <div class="badges">
        <span class="badge {tag_color}">{tag_label}</span>
        <span class="badge blue">{t['version']}</span>
        <span class="badge orange">⚡ {feat_count}</span>
        <span class="badge gray">🕒 {fmt_ts(last_ts)}</span>
        {health_badge}
        {sel_badge}
      </div>
      <div class="tc-bestfor">
        🎯 <b>Best for:</b> {t['best_for']}
      </div>
      <div class="tc-divider"></div>
      <ul style="padding:0; margin:0;">
        {top_html}
      </ul>
      {more_block}
    </div>
    """

    with col:
        st.markdown(card_html, unsafe_allow_html=True)
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        if st.button(f"Select {t['name']}", key=f"select_{tool_key}", use_container_width=True):
            st.session_state["selected_tool"] = tool_key
            st.session_state["launched"]      = False
            st.rerun()

render_card("tool1", card_col1)
render_card("tool2", card_col2)

# ── Side Panel ──
with side_col:

    # Suite Capabilities
    html("""
    <div class="cap-card">
      <div class="cap-title">✨ Suite Capabilities</div>
      <ul style="padding:0; margin:0; list-style:none;">
        <li class="cap-item">📂 Multi-file reconciliation</li>
        <li class="cap-item">🔑 Smart key auto-detection</li>
        <li class="cap-item">🧠 COALESCE key logic</li>
        <li class="cap-item">🟡 Duplicate detection + handling</li>
        <li class="cap-item">📊 Before vs After dup comparison</li>
      </ul>
      <details style="margin-top:4px;">
        <summary style="font-size:11px;color:var(--muted);cursor:pointer;list-style:none;">+ 7 more capabilities</summary>
        <ul style="padding:0; margin:4px 0 0 0; list-style:none;">
          <li class="cap-item">🔍 Search Loan Account Number</li>
          <li class="cap-item">🧹 Filter-based subset reco</li>
          <li class="cap-item">📋 Common / selected column compare</li>
          <li class="cap-item">📄 PDF Summary Report</li>
          <li class="cap-item">💾 Excel + CSV + ZIP outputs</li>
          <li class="cap-item">🔒 Local machine processing</li>
          <li class="cap-item">🛡️ Data never leaves machine</li>
        </ul>
      </details>
    </div>
    """)

    # Security & Support
    html("""
    <div class="cap-card">
      <div class="cap-title">🔒 Security & Support</div>
      <div style="font-size:13px; color:var(--text); line-height:1.8;">
        Data stays on your machine.<br>
        No external upload.<br>
        100% local processing.<br><br>
        <b>Support:</b><br>
        📧 aarohisharma5000@gmail.com
      </div>
    </div>
    """)

    # Open in new tab
    html("""
    <div class="linkbox">
      <div class="linktitle">🔗 Open in New Tab</div>
      <div class="link-row">
        📄 <a href="?tool=tool1&auto=1" target="_blank">Tool 1 — Classic Reco</a>
      </div>
      <div class="link-row">
        🧠 <a href="?tool=tool2&auto=1" target="_blank">Tool 2 — Big File Reco</a>
      </div>
      <div class="hint">
        💡 Tip: Open both tools in separate tabs to run simultaneously.
      </div>
    </div>
    """)

html('<div class="divider"></div>')

# ═══════════════════════════════════════════
# ACTION BAR
# ═══════════════════════════════════════════
t      = TOOLS[selected]
ok_sel, missing_sel = check_deps(t.get("deps", []))

html(f"""
<div class="selline">
  <div><b style="color:var(--text);">Selected:</b>
       <span class="badge green">{t['name']}</span></div>
  <div><b style="color:var(--text);">Version:</b>
       <span class="badge blue">{t['version']}</span></div>
  <div>{"<span class='badge green'>Ready to launch</span>"
        if ok_sel else
        "<span class='badge red'>Fix deps first</span>"}</div>
</div>
""")

a, b, c = st.columns([1.2, 1.3, 2.5], gap="medium")
with a:
    launch = st.button("🚀 Launch Selected Tool", type="primary",
                       use_container_width=True, disabled=(not ok_sel))
with b:
    back = st.button("✏️ Back to Launcher (Stop Tool)", use_container_width=True)
with c:
    if ok_sel:
        st.caption("Launch runs the selected tool. Back returns to this selection screen.")
    else:
        st.error(f"Cannot launch. Missing: {', '.join(missing_sel)}")

if launch:
    st.session_state["launched"] = True
    st.session_state["last_launched_at"][selected] = datetime.now()
    st.rerun()

if back:
    st.session_state["launched"] = False
    try:
        st.query_params.clear()
    except Exception:
        st.experimental_set_query_params()
    st.rerun()

html('<div class="divider"></div>')

# ═══════════════════════════════════════════
# HELP & USER GUIDE
# ═══════════════════════════════════════════
show_help = st.toggle("📘 Show Help & User Guide", value=False, key="launcher_help")

if show_help:
    st.markdown("## 📘 Help & User Guide")

    h1, h2 = st.columns(2, gap="large")

    with h1:
        html("""
        <div class="help-card">
          <div class="help-title">🚀 Quick Start</div>
          <div class="help-body">
            <b>Step 1</b> — Select Tool 1 or Tool 2<br>
            <b>Step 2</b> — Click Launch Selected Tool<br>
            <b>Step 3</b> — Upload your files (both sides)<br>
            <b>Step 4</b> — Select key column for matching<br>
            <b>Step 5</b> — Select columns to compare<br>
            <b>Step 6</b> — Apply filter if needed<br>
            <b>Step 7</b> — Click Run Reco<br>
            <b>Step 8</b> — Review summary + download outputs
          </div>
        </div>
        """)

        html("""
        <div class="help-card">
          <div class="help-title">📂 Supported File Formats</div>
          <div class="help-body">
            Tool 1 and Tool 2 both support:<br><br>
            • <b>CSV</b> — comma separated values<br>
            • <b>XLSX</b> — Excel 2007+ format<br>
            • <b>XLS</b> — Excel 97-2003 format<br>
            • <b>XLSB</b> — Excel Binary format<br><br>
            Multiple files can be uploaded on each side.<br>
            All files are combined before reconciliation.
          </div>
        </div>
        """)

        html("""
        <div class="help-card">
          <div class="help-title">🔑 How Key Matching Works</div>
          <div class="help-body">
            The tool uses <b>COALESCE key logic</b>:<br><br>
            • Auto-detects best key column from your headers<br>
            • Works with ANY column name — not just standard names<br>
            • If first key column is blank → uses next column<br>
            • Cleans keys — removes spaces, commas, quotes<br>
            • Handles scientific notation (Excel float issue)<br>
            • Keys are normalized to UPPERCASE for matching<br><br>
            <b>Good key columns:</b><br>
            Loan Account Number, LAN, Customer ID, Reference No, Reco, ID
          </div>
        </div>
        """)

        html("""
        <div class="help-card">
          <div class="help-title">🟡 How Duplicate Handling Works</div>
          <div class="help-body">
            When duplicates are detected the tool shows:<br><br>
            • Count of duplicate keys in File 1 and File 2<br>
            • Preview of top duplicate records<br>
            • Choice of how to handle them:<br><br>
            <b>✅ Include duplicates (recommended)</b><br>
            Aggregates / sums all rows per key before comparing.<br>
            Best for monthly transaction files where same loan appears multiple times.<br><br>
            <b>🚫 Exclude duplicates</b><br>
            Keeps only the first occurrence of each duplicate key.<br>
            Best when duplicates are accidental double uploads.<br><br>
            After running both modes a <b>Before vs After comparison table</b>
            is generated automatically showing the impact of removing duplicates.
          </div>
        </div>
        """)

    with h2:
        html("""
        <div class="help-card">
          <div class="help-title">🔍 Search Loan Account Number</div>
          <div class="help-body">
            The Search feature lets you instantly check if any Loan Account Number
            exists in your uploaded files — before running reconciliation.<br><br>
            <b>How to use:</b><br>
            1. Upload your files on both sides<br>
            2. Scroll to the Search section<br>
            3. Type any LAN / Lender LAN<br>
            4. Tool shows instantly:<br>
            &nbsp;&nbsp;&nbsp;• ✅ Present in File 1<br>
            &nbsp;&nbsp;&nbsp;• ✅ Present in File 2<br>
            &nbsp;&nbsp;&nbsp;• ❌ Not found<br><br>
            The search cleans the key the same way reconciliation does —
            so "DELPPL001.0" and "DELPPL001" will both match correctly.
          </div>
        </div>
        """)

        html("""
        <div class="help-card">
          <div class="help-title">📊 Before vs After Duplicate Comparison</div>
          <div class="help-body">
            This feature shows the impact of removing duplicates on your reconciliation result.<br><br>
            <b>How to get it:</b><br>
            1. Run Reco with <b>Include duplicates</b> selected<br>
            2. Switch to <b>Exclude duplicates</b><br>
            3. Click Run Reco again<br>
            4. Comparison table appears automatically<br><br>
            <b>What it shows:</b><br>
            • Match count WITH vs WITHOUT dups<br>
            • Mismatch count WITH vs WITHOUT dups<br>
            • Only-in counts WITH vs WITHOUT dups<br>
            • Sum differences for the focus column<br><br>
            This saves you from manually re-running with corrected files.
          </div>
        </div>
        """)

        html("""
        <div class="help-card">
          <div class="help-title">🧹 Filter Feature</div>
          <div class="help-body">
            Filter lets you run reconciliation on a <b>subset of records</b> only.<br><br>
            <b>Example:</b><br>
            Filter Column: Lender ID<br>
            Filter Value: 24<br><br>
            Only records matching Lender ID = 24 will be reconciled on both sides.<br><br>
            <b>Supports:</b><br>
            • Single value filter<br>
            • Multi-value filter (comma or newline separated)<br>
            • Dropdown picker of available values<br>
            • Case-insensitive matching
          </div>
        </div>
        """)

        html("""
        <div class="help-card">
          <div class="help-title">📥 Outputs Generated</div>
          <div class="help-body">
            After reconciliation the tool generates:<br><br>
            • 📋 <b>Summary CSV + PDF</b> — counts and sums by category<br>
            • ✅ <b>Matched</b> — records matching on both sides<br>
            • ❌ <b>Mismatched</b> — records with value differences<br>
            • ⬅ <b>Only in File 1</b> — records missing from File 2<br>
            • ➡ <b>Only in File 2</b> — records missing from File 1<br>
            • 🟡 <b>Duplicates</b> — File1 / File2 / Both<br>
            • 📦 <b>ZIP</b> — all files in one download<br>
            • 📊 <b>Excel</b> — formatted with subtotals + formulas<br><br>
            Large outputs are automatically split into multiple CSV files.
          </div>
        </div>
        """)

        html("""
        <div class="help-card">
          <div class="help-title">⚠️ Common Issues & Fixes</div>
          <div class="help-body">
            <b>1. XLS file error</b><br>
            Ensure xlrd>=2.0.1 is in requirements.txt<br><br>
            <b>2. XLSB file error</b><br>
            Ensure pyxlsb is in requirements.txt<br><br>
            <b>3. Key not detected</b><br>
            Manually select key column from the multiselect dropdown<br><br>
            <b>4. All zeros after Exclude duplicates</b><br>
            Your data has duplicates on every key — use Include mode instead<br><br>
            <b>5. File locked / permission denied</b><br>
            Close the source Excel file before uploading<br><br>
            <b>6. App goes to sleep</b><br>
            Streamlit Cloud free tier sleeps after inactivity — just reload the page
          </div>
        </div>
        """)

    html("""
    <div class="help-card">
      <div class="help-title">✅ Best Practices</div>
      <div class="help-body">
        • Keep source files closed while uploading (avoids file lock errors)<br>
        • Verify key column selection before clicking Run Reco<br>
        • Use Search LAN to spot-check a few records before full reconciliation<br>
        • Run with Include duplicates first — then Exclude if needed<br>
        • Download ZIP for the complete output pack<br>
        • Use Filter when you only need to check one lender at a time<br>
        • Use Tool 2 for files larger than 100MB
      </div>
    </div>
    """)

    st.divider()
    st.subheader("📞 Support")
    st.info(
        "For issues reach out to Finance Ops / tool owner.\n\n"
        "📧 Contact: aarohisharma5000@gmail.com"
    )

html('<div class="divider"></div>')

# ═══════════════════════════════════════════
# RUN TOOL
# ═══════════════════════════════════════════
if not st.session_state["launched"]:
    st.info("Select a tool, then click **Launch Selected Tool**.")
else:
    tool_key = st.session_state["selected_tool"]
    try:
        if tool_key == "tool1":
            # ✅ FIX: pass full globals so exec'd code has builtins + correct __name__
            _globals = globals().copy()
            _globals["__name__"] = "__main__"
            exec(
                open(str(BASE / "app.py"), encoding="utf-8").read(),
                _globals
            )
        elif tool_key == "tool2":
            # ✅ FIX: pass full globals so exec'd code has builtins + correct __name__
            _globals = globals().copy()
            _globals["__name__"] = "__main__"
            exec(
                open(str(BASE / "app_v2.py"), encoding="utf-8").read(),
                _globals
            )
    except SystemExit:
        pass
    except Exception as e:
        st.error(f"Error loading tool: {e}")
