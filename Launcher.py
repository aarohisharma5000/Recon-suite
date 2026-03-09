import streamlit as st 
from pathlib import Path
from datetime import datetime
import importlib.util
import textwrap
import os, sys
import platform  # ✅ added (for pro status bar)

from pathlib import Path

def app_base_dir() -> Path:
    """
    Portable base folder:
    - In EXE: folder containing the .exe
    - In normal python: folder containing this launcher.py
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent

BASE_DIR = app_base_dir()

st.set_page_config(page_title= "Reco Suite Launcher", layout="wide")

BASE = BASE_DIR  # ✅ use portable base
ASSETS = BASE / "assets"
LOGO_PATH = ASSETS / "logo.png"

# ✅ Standard portable folders (create if missing)
DATA_DIR  = BASE / "data"
INPUT_DIR = DATA_DIR / "input"
FILE1_DIR = INPUT_DIR / "file1"
FILE2_DIR = INPUT_DIR / "file2"
CACHE_DIR = DATA_DIR / "cache"
RUNS_DIR  = DATA_DIR / "runs"

for p in [ASSETS, DATA_DIR, INPUT_DIR, FILE1_DIR, FILE2_DIR, CACHE_DIR, RUNS_DIR]:
    p.mkdir(parents=True, exist_ok=True)

LAUNCHER_VERSION = "v1.0"


# ---------- helpers ----------
def html(s: str):
    st.markdown(textwrap.dedent(s).strip(), unsafe_allow_html=True)

def fmt_ts(dt):
    if not dt:
        return "—"
    return dt.strftime("%d-%b-%Y %I:%M %p")


# ✅ 5-line deps checker (Point 3)
def check_deps(mods):
    missing = [m for m in mods if importlib.util.find_spec(m) is None]
    return (len(missing) == 0), missing


# ---------- session state ----------
if "selected_tool" not in st.session_state:
    st.session_state["selected_tool"] = "tool1"
if "launched" not in st.session_state:
    st.session_state["launched"] = False
if "last_launched_at" not in st.session_state:
    st.session_state["last_launched_at"] = {"tool1": None, "tool2": None}
if "dark_mode" not in st.session_state:
    st.session_state["dark_mode"] = False


# ---------- query params (new tab open) ----------
try:
    qp = dict(st.query_params)
    qp_tool = qp.get("tool", "")
    qp_auto = qp.get("auto", "0")
except Exception:
    qp = st.experimental_get_query_params()
    qp_tool = (qp.get("tool", [""])[0])
    qp_auto = (qp.get("auto", ["0"])[0])


# ---------- Tool registry ----------
TOOLS = {
    "tool1": {
        "name": "Tool 1",
        "desc": "Recon Tool 1 (app.py)",
        "icon": "📄",
        "version": "v1.5",
        "path": BASE / "app.py",
        "deps": ["streamlit", "pandas", "openpyxl"],  # ✅ openpyxl needed by Tool1
    },
    "tool2": {
        "name": "Tool 2",
        "desc": "Recon Tool 2 (app_v2.py)",
        "icon": "🧠",
        "version": "v2.0",
        "path": BASE / "app_v2.py",  # ✅ FIX OPTION A (your file is in root)
        "deps": ["streamlit", "pandas", "duckdb", "pyarrow"],
    },
}

# auto launch for new-tab links
if qp_tool in TOOLS and str(qp_auto) == "1":
    st.session_state["selected_tool"] = qp_tool
    st.session_state["launched"] = True

selected = st.session_state["selected_tool"]


# ---------- dark mode toggle (Streamlit widget) ----------
toggle_col_l, toggle_col_r = st.columns([0.85, 0.15], vertical_alignment="center")
with toggle_col_r:
    st.session_state["dark_mode"] = st.toggle("🌙 Dark", value=st.session_state["dark_mode"])


# ---------- Premium CSS ----------
is_dark = st.session_state["dark_mode"]

html(f"""
<style>
  header[data-testid="stHeader"] {{
    height: 0rem !important;
    visibility: hidden !important;
  }}
  section[data-testid="stSidebar"] {{
    border-right: 1px solid rgba(229,231,235,0.7);
  }}
  footer {{ visibility: hidden; }}

  :root {{
    --bg: {"#0b1220" if is_dark else "#f7fafc"};
    --panel: {"#0f172a" if is_dark else "#ffffff"};
    --panel2: {"#0b1326" if is_dark else "#f6faff"};
    --text: {"#e5e7eb" if is_dark else "#111827"};
    --muted: {"#a1a1aa" if is_dark else "#6b7280"};
    --border: {"rgba(148,163,184,0.25)" if is_dark else "#e5e7eb"};
    --brand: #0d6efd;
    --good: #198754;
    --bad: #dc3545;
    --shadow: {"rgba(0,0,0,0.35)" if is_dark else "rgba(13,110,253,0.12)"};
  }}

  .stApp {{
    background: var(--bg);
  }}

  .hero {{
    border: 1px solid var(--border);
    border-radius: 20px;
    padding: 18px 18px;
    background:
      radial-gradient(900px 300px at 20% 0%, rgba(13,110,253,0.18), transparent 60%),
      radial-gradient(900px 300px at 90% 10%, rgba(25,135,84,0.16), transparent 55%),
      linear-gradient(180deg, rgba(255,255,255,0.06), rgba(255,255,255,0.02));
    box-shadow: 0 14px 50px rgba(0,0,0,0.06);
    margin-top: 8px;
    margin-bottom: 10px;
  }}

  .heroRow {{
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap: 14px;
    flex-wrap: wrap;
  }}

  .heroTitle {{
    font-size: 40px;
    font-weight: 900;
    letter-spacing: -0.6px;
    color: var(--text);
    line-height: 1.0;
    margin: 0;
  }}

  .heroVer {{
    font-size: 12px;
    font-weight: 900;
    padding: 4px 10px;
    border-radius: 999px;
    border: 1px solid var(--border);
    color: var(--muted);
    display:inline-block;
    margin-left: 10px;
    vertical-align: middle;
    background: rgba(255,255,255,0.04);
  }}

  .heroSub {{
    color: var(--muted);
    font-size: 13px;
    margin: 0;
  }}

  .pill {{
    display:inline-block;
    padding:6px 12px;
    border-radius:999px;
    font-size:12px;
    font-weight:900;
    border:1px solid var(--border);
    background: rgba(255,255,255,0.04);
    color: var(--text);
    white-space:nowrap;
  }}
  .pill.blue {{
    background: rgba(13,110,253,0.16);
    border-color: rgba(13,110,253,0.28);
    color: {"#cfe0ff" if is_dark else "#0b5ed7"};
  }}

  .divider{{
    height: 1px;
    background: var(--border);
    margin: 18px 0 18px 0;
  }}

  .card {{
    border:1px solid var(--border);
    border-radius:18px;
    padding:18px;
    background: var(--panel);
    transition: transform 160ms ease, box-shadow 160ms ease, border-color 160ms ease;
  }}
  .card:hover {{
    transform: translateY(-2px);
    box-shadow: 0 12px 28px rgba(0,0,0,0.08);
  }}
  .card.selected {{
    border:2px solid var(--brand);
    background: var(--panel2);
    box-shadow: 0 10px 26px var(--shadow);
  }}

  .row {{ display:flex; align-items:flex-start; gap:12px; }}
  .icon {{
    width:46px; height:46px;
    border-radius:14px;
    display:flex; align-items:center; justify-content:center;
    font-size:22px;
    background: {"#111827" if not is_dark else "#020617"};
    color:white;
    flex: 0 0 46px;
  }}

  .title {{
    font-size:18px;
    font-weight:900;
    margin:0;
    line-height:1.1;
    color: var(--text);
  }}
  .desc {{
    color: var(--muted);
    font-size:13px;
    margin-top:4px;
  }}

  .meta {{
    color: var(--muted);
    font-size:12px;
    margin-top:10px;
    line-height:1.4;
    word-break: break-word;
  }}

  .badges {{
    display:flex;
    gap:8px;
    margin-top:10px;
    flex-wrap:wrap;
    align-items:center;
  }}

  .badge {{
    display:inline-block;
    padding:5px 10px;
    border-radius:999px;
    font-size:12px;
    font-weight:900;
    border:1px solid var(--border);
    background: rgba(255,255,255,0.04);
    color: var(--text);
    white-space:nowrap;
  }}
  .badge.blue {{ background: rgba(13,110,253,0.16); border-color: rgba(13,110,253,0.28); color: {"#cfe0ff" if is_dark else "#0b5ed7"}; }}
  .badge.green {{ background: rgba(25,135,84,0.18); border-color: rgba(25,135,84,0.28); color: {"#c7f9d6" if is_dark else "#0f5132"}; }}
  .badge.red {{ background: rgba(220,53,69,0.18); border-color: rgba(220,53,69,0.30); color: {"#ffd0d6" if is_dark else "#842029"}; }}
  .badge.gray {{ background: rgba(255,255,255,0.03); }}

  .linkbox {{
    border:1px solid var(--border);
    border-radius:18px;
    padding:16px;
    background: var(--panel);
  }}
  .linktitle {{ font-weight:900; margin-bottom:8px; color: var(--text); }}
  .hint {{ color: var(--muted); font-size:12px; margin-top:10px; }}
  .linkbox a {{ text-decoration:none; font-weight:800; }}

  .selline{{
    display:flex;
    gap:12px;
    flex-wrap:wrap;
    align-items:center;
    margin: 6px 0 18px 0;
  }}

</style>
""")


# ---------- SaaS hero header ----------
env_name = "localhost"

hero_left, hero_mid, hero_right = st.columns([0.18, 0.64, 0.18], vertical_alignment="center")
with hero_left:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=130)
with hero_mid:
    html(f"""
    <div class="hero">
      <div class="heroRow" style="justify-content:center;">
        <div style="text-align:center;">
          <div class="heroTitle">
                • Reconciliation Suite
            <span class="heroVer">Launcher {LAUNCHER_VERSION}</span>
          </div>
          <div class="heroSub">
            Select tool → Launch (no auto-run / no loops). Use “Open in new tab” if you want both tools at once.
          </div>
        </div>
      </div>
    </div>
    """)
with hero_right:
    html(f"""
    <div style="display:flex; justify-content:flex-end;">
      <span class="pill blue">Environment: {env_name}</span>
    </div>
    """)

# ✅ Point 4 upgrade: Pro System Status bar
st.markdown("### ✅ System Status")
s1, s2, s3, s4 = st.columns(4)
s1.metric("Python", platform.python_version())
s2.metric("OS", platform.system())
s3.metric("Base Folder", BASE_DIR.name)
s4.metric("URL", "Cloud")

with st.expander("📌 Environment Details", expanded=False):
    st.code(f"""BASE_DIR: {BASE_DIR}
Executable: {sys.executable}
Frozen: {getattr(sys, "frozen", False)}
""".strip(), language="text")


html('<div class="divider"></div>')

# ---------- Professional Home Dashboard UI ----------
st.markdown("### 🏠 Suite Overview")

k1, k2, k3, k4 = st.columns(4)
k1.metric("Tools Available", "2")
k2.metric("Supported Formats", "CSV / XLSX / XLS / XLSB")
k3.metric("Engine", "Pandas + DuckDB")
k4.metric("Runs On", "Local Machine")

st.markdown("---")

left_home, right_home = st.columns([2.2, 1], gap="large")

with left_home:
    st.markdown("### 🚀 Available Tools")

    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.markdown("""
        <div style="
            border:1px solid #d9e8ff;
            border-radius:16px;
            padding:18px;
            background:#f8fbff;
            min-height:260px;
        ">
            <div style="font-size:22px; font-weight:700; color:#0b5ed7;">📄 Tool 1</div>
            <div style="font-size:14px; color:#5b6470; margin-top:6px;">
                Classic Reconciliation Tool
            </div>
            <hr style="margin:12px 0;">
            <div style="font-size:14px; line-height:1.8;">
                <b>Best for:</b> Standard file comparison<br>
                <b>Supports:</b> CSV / XLSX / XLS / XLSB<br>
                <b>Highlights:</b><br>
                • Multi-file comparison<br>
                • Duplicate key outputs<br>
                • Excel + CSV downloads<br>
                • COALESCE key logic
            </div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown("""
        <div style="
            border:1px solid #d9e8ff;
            border-radius:16px;
            padding:18px;
            background:#f8fbff;
            min-height:260px;
        ">
            <div style="font-size:22px; font-weight:700; color:#0b5ed7;">🧠 Tool 2</div>
            <div style="font-size:14px; color:#5b6470; margin-top:6px;">
                Large File Reconciliation Tool
            </div>
            <hr style="margin:12px 0;">
            <div style="font-size:14px; line-height:1.8;">
                <b>Best for:</b> Large datasets / faster processing<br>
                <b>Supports:</b> CSV / XLSX / XLS / XLSB<br>
                <b>Highlights:</b><br>
                • DuckDB engine<br>
                • Parquet caching<br>
                • Duplicate preview + export<br>
                • ZIP downloads + summaries
            </div>
        </div>
        """, unsafe_allow_html=True)

with right_home:
    st.markdown("""
    <div style="
        border:1px solid #d9e8ff;
        border-radius:16px;
        padding:18px;
        background:#ffffff;
        margin-bottom:14px;
    ">
        <div style="font-size:18px; font-weight:700; color:#0b5ed7;">✨ Suite Capabilities</div>
        <div style="font-size:14px; color:#5b6470; margin-top:10px; line-height:1.8;">
            • Multi-file reconciliation<br>
            • Duplicate detection<br>
            • Filter-based subset reco<br>
            • Common / selected column compare<br>
            • Local machine processing<br>
            • CSV / ZIP / Excel outputs
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="
        border:1px solid #d9e8ff;
        border-radius:16px;
        padding:18px;
        background:#ffffff;
    ">
        <div style="font-size:18px; font-weight:700; color:#0b5ed7;">🔒 Security & Support</div>
        <div style="font-size:14px; color:#5b6470; margin-top:10px; line-height:1.8;">
            Data stays on your machine.<br>
            No external upload.<br><br>
            <b>Support:</b><br>
            aarohi.sharma@paytm.com
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")


# ---------- layout cards ----------

left, mid, right = st.columns([1, 1, 0.65], gap="large")


def render_card(tool_key: str, col):
    t = TOOLS[tool_key]
    is_sel = (selected == tool_key)
    last_ts = st.session_state["last_launched_at"].get(tool_key)

    ok, missing = check_deps(t.get("deps", []))
    health_badge = (
        "<span class='badge green'>✅ deps OK</span>"
        if ok
        else f"<span class='badge red'>❌ missing: {', '.join(missing)}</span>"
    )
    sel_badge = "<span class='badge gray'>Selected</span>" if is_sel else ""

    card_html = "\n".join([
        f"<div class='card {'selected' if is_sel else ''}'>",
        "<div class='row'>",
        f"<div class='icon'>{t['icon']}</div>",
        "<div class='content'>",
        f"<div class='title'>{t['name']}</div>",
        f"<div class='desc'>{t['desc']}</div>",
        "<div class='badges'>",
        f"<span class='badge blue'>{t['version']}</span>",
        f"<span class='badge gray'>Last launch: {fmt_ts(last_ts)}</span>",
        f"{health_badge}",
        f"{sel_badge}",
        "</div>",
        f"<div class='meta'><b>Path:</b> {t['path']}</div>",
        "</div>",
        "</div>",
        "</div>",
    ])

    with col:
        st.markdown(card_html, unsafe_allow_html=True)
        if st.button(f"Select {t['name']}", key=f"select_{tool_key}", use_container_width=True):
            st.session_state["selected_tool"] = tool_key
            st.session_state["launched"] = False
            st.rerun()

render_card("tool1", left)
render_card("tool2", mid)

with right:
    html("""
    <div class="linkbox">
      <div class="linktitle">Open in new tab</div>
      🔗 <a href="?tool=tool1&auto=1" target="_blank">Tool 1</a><br/><br/>
      🔗 <a href="?tool=tool2&auto=1" target="_blank">Tool 2</a>
      <div class="hint">Tip: Open Tool 1 and Tool 2 in separate tabs if you want both running together.</div>
    </div>
    """)

html('<div class="divider"></div>')


# ---------- action bar ----------
t = TOOLS[selected]
ok_sel, missing_sel = check_deps(t.get("deps", []))

html(f"""
<div class="selline">
  <div><b style="color:var(--text);">Selected:</b> <span class="badge green">{t['name']}</span></div>
  <div><b style="color:var(--text);">Version:</b> <span class="badge blue">{t['version']}</span></div>
  <div>{"<span class='badge green'>Ready to launch</span>" if ok_sel else "<span class='badge red'>Fix deps first</span>"}</div>
</div>
""")

a, b, c = st.columns([1.2, 1.1, 2.7], gap="medium")
with a:
    launch = st.button("🚀 Launch Selected Tool", type="primary", use_container_width=True, disabled=(not ok_sel))
with b:
    back = st.button("🧹 Back to Launcher (Stop Tool)", use_container_width=True)
with c:
    if ok_sel:
        st.caption("Launch runs the selected tool. Back returns to this selection screen.")
    else:
        st.error(f"Cannot launch. Missing dependencies: {', '.join(missing_sel)}")

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

# ---------- run tool only after launch ----------
if not st.session_state["launched"]:
    st.info("Select a tool, then click **Launch Selected Tool**.")
else:
    tool_key = st.session_state["selected_tool"]
    try:
        if tool_key == "tool1":
            import app
            app.main() if hasattr(app, 'main') else None
        elif tool_key == "tool2":
            import app_v2
            app_v2.main() if hasattr(app_v2, 'main') else None
    except Exception as e:
        st.error(f"Error loading tool: {e}")




