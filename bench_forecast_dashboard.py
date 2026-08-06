"""
NA Bench Forecast Dashboard
Streamlit dashboard replicating three Excel sheets:
  1. Bench Forecast  - editable table (per-center weekly values)
  2. Historic Bench Data - stacked bar chart by center with NA FNC line overlay
  3. Q3 2025 tab - Q3 26 Forecast vs Q3 25 Bench vs Q3 Original Forecast line chart
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pathlib import Path

# -----------------------------------------------------------------------------
# Palette  -  Deep teal / amber / slate
# -----------------------------------------------------------------------------
# Background / surfaces
BG          = "#0d1117"      # near-black canvas
SURFACE     = "#161b22"      # card background
SURFACE2    = "#21262d"      # raised / hover
BORDER      = "#30363d"      # subtle divider
TEXT_PRI    = "#e6edf3"      # primary text
TEXT_SEC    = "#8b949e"      # muted / labels
ACCENT      = "#38bdf8"      # sky-blue accent (headers, highlights)
ACCENT2     = "#f59e0b"      # amber (secondary accent)
GOOD        = "#34d399"      # emerald green (positive)
WARN        = "#fb923c"      # orange (warn)
DANGER      = "#f87171"      # soft red

# Chart series colours - distinct, accessible, dark-bg friendly
CENTER_COLORS = {
    "Baton Rouge":  "#38bdf8",   # sky blue
    "Buffalo":      "#818cf8",   # indigo
    "Calgary":      "#34d399",   # emerald
    "Halifax":      "#f59e0b",   # amber
    "Lansing":      "#fb923c",   # orange
    "Monroe":       "#e879f9",   # fuchsia
    "Quebec":       "#a3e635",   # lime
    "East Lansing": "#818cf8",   # indigo (same family as Buffalo for historic tab)
}

# Q3 comparison series
C_FORECAST   = "#38bdf8"   # sky blue - Q3 26 Forecast
C_ACTUAL     = "#34d399"   # emerald  - Q3 25 Actual
C_ORIG       = "#94a3b8"   # slate    - Q3 Original Forecast

# Table row colours - per-center (background, text)
CENTER_ROW_COLORS = {
    "Baton Rouge":  ("#0e2a3a", "#38bdf8"),
    "Buffalo":      ("#1a1a3e", "#818cf8"),
    "Calgary":      ("#0d2e22", "#34d399"),
    "Halifax":      ("#2e2200", "#f59e0b"),
    "Lansing":      ("#2e1800", "#fb923c"),
    "Monroe":       ("#2a0e2a", "#e879f9"),
    "Quebec":       ("#1a2200", "#a3e635"),
    "East Lansing": ("#1a1a3e", "#818cf8"),
}
# Header row for styled read-only tables
TBL_HEADER_BG   = "#0f2437"
TBL_HEADER_FG   = ACCENT
TBL_TOTAL_BG    = "#1c1c1c"
TBL_TOTAL_FG    = "#ffffff"

# -----------------------------------------------------------------------------
# Config
# -----------------------------------------------------------------------------
XLSX_PATH = Path(__file__).parent / "NA Bench Forecast.xlsx"
CENTERS   = ["Baton Rouge", "Buffalo", "Calgary", "Halifax", "Lansing", "Monroe", "Quebec"]
WEEKS     = [f"Wk {i:02d}" for i in range(1, 14)]

st.set_page_config(
    page_title="NA Bench Forecast",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# -----------------------------------------------------------------------------
# Access gate - password required
# Set via Streamlit Cloud Secrets:  [auth]  password = "your-password"
# For local dev without secrets, set env var:  BENCH_PASSWORD=your-password
# -----------------------------------------------------------------------------
import os

def _get_password():  # -> str | None
    """Return the required password, or None if no gate is configured."""
    try:
        return st.secrets["auth"]["password"]
    except Exception:
        return os.environ.get("BENCH_PASSWORD")

_required_pw = _get_password()

if _required_pw:
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.markdown(f"""
        <style>
        html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"] {{
            background: {BG} !important;
        }}
        .lock-wrap {{
            display: flex; flex-direction: column; align-items: center;
            justify-content: center; min-height: 80vh;
        }}
        .lock-card {{
            background: {SURFACE}; border: 1px solid {BORDER};
            border-top: 3px solid {ACCENT}; border-radius: 12px;
            padding: 40px 44px; max-width: 380px; width: 100%;
            text-align: center;
        }}
        .lock-title {{
            font-size: 18px; font-weight: 700; color: {TEXT_PRI};
            margin-bottom: 4px;
        }}
        .lock-sub {{
            font-size: 13px; color: {TEXT_SEC}; margin-bottom: 28px;
        }}
        </style>
        <div class="lock-wrap">
          <div class="lock-card">
            <div class="lock-title">NA Bench Forecast</div>
            <div class="lock-sub">IBM FNC Workforce Planning &mdash; restricted access</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        pw_input = st.text_input(
            "Access code", type="password", placeholder="Enter access code...",
            label_visibility="collapsed",
        )
        col_btn, _ = st.columns([1, 3])
        with col_btn:
            if st.button("Unlock", use_container_width=True):
                if pw_input == _required_pw:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Incorrect access code.")
        st.stop()

# -----------------------------------------------------------------------------
# Editing lock - admin toggle via Streamlit Cloud Secrets
# Set  [lock]  editing = true  in Secrets to make the forecast table read-only.
# Remove or set to false to re-enable editing.
# -----------------------------------------------------------------------------
try:
    EDITING_LOCKED = bool(st.secrets["lock"]["editing"])
except Exception:
    EDITING_LOCKED = False

# -----------------------------------------------------------------------------
# Global CSS - dark theme
# -----------------------------------------------------------------------------
st.markdown(f"""
<style>
    /* -- canvas -- */
    html, body, [data-testid="stAppViewContainer"],
    [data-testid="stMain"], .main .block-container {{
        background: {BG} !important;
        color: {TEXT_PRI};
    }}
    [data-testid="stHeader"] {{ background: {BG} !important; }}
    [data-testid="stSidebar"] {{ background: {SURFACE} !important; }}

    /* -- typography -- */
    * {{ font-family: "Inter", "Segoe UI", system-ui, sans-serif !important; }}
    h1, h2, h3, h4 {{ color: {TEXT_PRI} !important; }}
    p, li, label {{ color: {TEXT_SEC}; }}

    /* -- hero header -- */
    .dash-hero {{
        background: linear-gradient(135deg, {SURFACE} 0%, #0f2437 100%);
        border: 1px solid {BORDER};
        border-radius: 10px;
        padding: 28px 36px 22px;
        margin-bottom: 28px;
        display: flex;
        align-items: center;
        gap: 20px;
    }}
    .dash-hero .logo-circle {{
        width: 52px; height: 52px;
        background: {ACCENT};
        border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.6rem; flex-shrink: 0;
    }}
    .dash-hero h1 {{
        margin: 0 0 4px; font-size: 1.55rem; font-weight: 700;
        color: {TEXT_PRI} !important;
        letter-spacing: -0.3px;
    }}
    .dash-hero .sub {{
        font-size: 0.83rem; color: {TEXT_SEC}; margin: 0;
    }}
    .dash-hero .badge {{
        margin-left: auto;
        background: rgba(56,189,248,0.12);
        border: 1px solid rgba(56,189,248,0.3);
        color: {ACCENT};
        border-radius: 20px;
        padding: 4px 14px;
        font-size: 0.78rem; font-weight: 600;
        white-space: nowrap;
    }}

    /* -- section titles -- */
    .sec-title {{
        font-size: 0.82rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: {ACCENT};
        border-left: 3px solid {ACCENT};
        padding-left: 10px;
        margin: 0 0 16px;
    }}

    /* -- stat cards -- */
    .stat-row {{ display: flex; gap: 14px; margin-bottom: 22px; flex-wrap: wrap; }}
    .stat-card {{
        flex: 1; min-width: 120px;
        background: {SURFACE};
        border: 1px solid {BORDER};
        border-radius: 8px;
        padding: 14px 18px;
    }}
    .stat-card .val  {{ font-size: 1.6rem; font-weight: 700; color: {TEXT_PRI}; line-height: 1; }}
    .stat-card .lbl  {{ font-size: 0.75rem; color: {TEXT_SEC}; margin-top: 4px; }}
    .stat-card .delta-up   {{ color: {GOOD};  font-size: 0.78rem; }}
    .stat-card .delta-down {{ color: {DANGER}; font-size: 0.78rem; }}

    /* -- tabs -- */
    [data-testid="stTabs"] [role="tablist"] {{
        background: {SURFACE};
        border-radius: 8px 8px 0 0;
        padding: 4px 8px 0;
        border-bottom: 1px solid {BORDER};
        gap: 4px;
    }}
    [data-testid="stTabs"] button[role="tab"] {{
        color: {TEXT_SEC} !important;
        border-radius: 6px 6px 0 0;
        padding: 8px 18px;
        font-size: 0.85rem;
        font-weight: 500;
        border: none;
        background: transparent;
    }}
    [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {{
        color: {ACCENT} !important;
        background: rgba(56,189,248,0.08) !important;
        border-bottom: 2px solid {ACCENT} !important;
    }}

    /* -- data editor / dataframe -- */
    [data-testid="stDataEditor"] {{ border-radius: 8px; border: 1px solid {BORDER}; }}

    /* -- data editor: header row -- */
    [data-testid="stDataEditor"] thead tr th {{
        background: {TBL_HEADER_BG} !important;
        color: {TBL_HEADER_FG} !important;
        font-size: 0.78rem !important;
        font-weight: 700 !important;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        border-bottom: 2px solid {ACCENT} !important;
    }}
    /* -- data editor: center (first) column -- */
    [data-testid="stDataEditor"] tbody tr td:first-child {{
        background: {SURFACE2} !important;
        color: {ACCENT} !important;
        font-weight: 600 !important;
        font-size: 0.82rem !important;
    }}
    /* -- data editor: data cells -- */
    [data-testid="stDataEditor"] tbody tr td {{
        background: {SURFACE} !important;
        color: {TEXT_PRI} !important;
        border-color: {BORDER} !important;
    }}
    /* -- data editor: alternating rows -- */
    [data-testid="stDataEditor"] tbody tr:nth-child(even) td {{
        background: {SURFACE2} !important;
    }}

    /* -- buttons -- */
    [data-testid="stDownloadButton"] button {{
        background: transparent !important;
        border: 1px solid {ACCENT} !important;
        color: {ACCENT} !important;
        border-radius: 6px;
        font-size: 0.83rem;
        font-weight: 600;
        padding: 6px 16px;
    }}
    [data-testid="stDownloadButton"] button:hover {{
        background: rgba(56,189,248,0.12) !important;
    }}

    /* -- metrics -- */
    [data-testid="metric-container"] {{
        background: {SURFACE};
        border: 1px solid {BORDER};
        border-radius: 8px;
        padding: 12px 16px !important;
    }}
    [data-testid="stMetricValue"] {{ color: {TEXT_PRI} !important; font-size: 1.4rem !important; }}
    [data-testid="stMetricLabel"] {{ color: {TEXT_SEC} !important; font-size: 0.75rem !important; }}

    /* -- multiselect -- */
    [data-baseweb="select"] {{ background: {SURFACE2} !important; border-color: {BORDER} !important; }}

    /* -- number input -- */
    [data-testid="stNumberInput"] input {{
        background: {SURFACE2} !important;
        border: 1px solid {BORDER} !important;
        color: {TEXT_PRI} !important;
        border-radius: 4px;
        text-align: center !important;
        font-size: 0.85rem !important;
        padding: 4px 2px !important;
    }}
    /* tighten spinner buttons */
    [data-testid="stNumberInput"] button {{
        background: {SURFACE2} !important;
        border-color: {BORDER} !important;
        color: {TEXT_SEC} !important;
    }}
    /* expander header styling */
    [data-testid="stExpander"] summary {{
        background: {SURFACE} !important;
        border: 1px solid {BORDER} !important;
        border-radius: 6px;
        color: {ACCENT} !important;
        font-weight: 600;
        font-size: 0.85rem;
    }}
    [data-testid="stExpander"] > div[data-testid="stExpanderDetails"] {{
        background: {SURFACE} !important;
        border: 1px solid {BORDER} !important;
        border-top: none !important;
        border-radius: 0 0 6px 6px;
        padding: 12px 8px 16px;
    }}

    /* -- checkbox -- */
    [data-testid="stCheckbox"] label {{ color: {TEXT_SEC} !important; }}

    /* -- dividers -- */
    hr {{ border-color: {BORDER} !important; }}

    footer {{ visibility: hidden; }}
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Hero Header
# -----------------------------------------------------------------------------
st.markdown(f"""
<div class="dash-hero">
    <div class="logo-circle"></div>
    <div>
        <h1>NA Bench Forecast</h1>
        <p class="sub">North America FNC . Bench planning, historic trends &amp; quarter comparisons</p>
    </div>
    <div class="badge">Q3 2026 Active</div>
</div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Shared chart layout defaults
# -----------------------------------------------------------------------------
CHART_LAYOUT = dict(
    plot_bgcolor   = SURFACE,
    paper_bgcolor  = SURFACE,
    font           = dict(color=TEXT_SEC, family="Inter, Segoe UI, system-ui"),
    legend         = dict(
        orientation="h", y=-0.22,
        font=dict(size=11, color=TEXT_SEC),
        bgcolor="rgba(0,0,0,0)",
    ),
    margin         = dict(l=12, r=12, t=44, b=60),
    xaxis          = dict(
        tickfont   = dict(size=10, color=TEXT_SEC),
        gridcolor  = BORDER,
        zerolinecolor = BORDER,
        linecolor  = BORDER,
        showgrid   = False,
    ),
    yaxis          = dict(
        tickfont   = dict(size=10, color=TEXT_SEC),
        gridcolor  = BORDER,
        zerolinecolor = BORDER,
        linecolor  = BORDER,
    ),
    hoverlabel     = dict(
        bgcolor    = SURFACE2,
        bordercolor= BORDER,
        font       = dict(color=TEXT_PRI, size=12),
    ),
)

# -----------------------------------------------------------------------------
# Styled-table helper
# -----------------------------------------------------------------------------
def styled_table(df: pd.DataFrame, total_rows: list[str] | None = None) -> None:
    """Render a read-only DataFrame with per-center row colours, dark header, centred numbers."""
    total_rows = total_rows or []

    def _row_style(row):
        center = str(row["Center"]) if "Center" in row.index else ""
        if center in total_rows:
            bg, fg = TBL_TOTAL_BG, TBL_TOTAL_FG
        elif center in CENTER_ROW_COLORS:
            bg, fg = CENTER_ROW_COLORS[center]
        else:
            bg, fg = SURFACE, TEXT_PRI
        return [f"background-color:{bg}; color:{fg}"] * len(row)

    # Centre every column except the first (Center name)
    def _col_align(col):
        if col.name == "Center":
            return ["text-align:left; padding-left:12px"] * len(col)
        return ["text-align:center"] * len(col)

    styler = (
        df.style
        .apply(_row_style, axis=1)
        .apply(_col_align, axis=0)
        .set_table_styles([
            # header row - centred, except first th (Center label stays left)
            {"selector": "thead tr th",
             "props": f"background-color:{TBL_HEADER_BG}; color:{TBL_HEADER_FG}; "
                      f"font-weight:700; font-size:0.78rem; text-transform:uppercase; "
                      f"letter-spacing:0.05em; border-bottom:2px solid {ACCENT}; "
                      f"text-align:center; padding:7px 10px;"},
            # first header cell - left-align the "Center" column header
            {"selector": "thead tr th:first-child",
             "props": "text-align:left; padding-left:12px;"},
            # all data cells
            {"selector": "td",
             "props": f"border:1px solid {BORDER}; font-size:0.85rem; padding:6px 10px;"},
            # table wrapper
            {"selector": "",
             "props": f"border-collapse:collapse; width:100%;"},
        ])
        .format(precision=0, na_rep="-")
    )
    st.markdown(
        "<div style='overflow-x:auto;border-radius:8px;border:1px solid "
        + BORDER + ";'>" + styler.to_html(index=False) + "</div>",
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# Data loaders
# -----------------------------------------------------------------------------
@st.cache_data
def load_bench_forecast() -> pd.DataFrame:
    raw = pd.read_excel(XLSX_PATH, sheet_name="Bench Forecast", header=0)
    raw.columns = ["Center"] + WEEKS + (list(raw.columns[14:]) if len(raw.columns) > 14 else [])
    raw = raw[raw["Center"].isin(CENTERS)].reset_index(drop=True)
    for w in WEEKS:
        raw[w] = pd.to_numeric(raw[w], errors="coerce").fillna(0).astype(int)
    return raw[["Center"] + WEEKS]


@st.cache_data
def load_historic() -> pd.DataFrame:
    raw = pd.read_excel(XLSX_PATH, sheet_name="Historic Bench Data", header=0)
    raw = raw.iloc[:, :12]
    raw.columns = [
        "Year", "Quarter", "Week",
        "Baton Rouge", "East Lansing", "Monroe", "Buffalo",
        "Halifax", "Quebec", "Calgary", "NA FNC", "Wk",
    ]
    raw = raw[raw["Year"].notna() & raw["Quarter"].notna()].copy()
    raw["Year"] = raw["Year"].astype(int)
    raw["Label"] = raw["Year"].astype(str) + " " + raw["Quarter"] + " " + raw["Week"]
    return raw


@st.cache_data
def load_q3_comparison() -> pd.DataFrame:
    raw = pd.read_excel(XLSX_PATH, sheet_name="Q3 2025", header=None)
    pivot_start = None
    for i, row in raw.iterrows():
        if str(row.iloc[0]).strip() == "Total Bench":
            pivot_start = i
            break
    if pivot_start is None:
        return pd.DataFrame()
    weeks_row = raw.iloc[pivot_start]
    weeks = [str(w).strip() for w in weeks_row.iloc[1:14].tolist()]
    rows = {}
    for offset in [1, 2, 3]:
        r = raw.iloc[pivot_start + offset]
        rows[str(r.iloc[0]).strip()] = pd.to_numeric(r.iloc[1:14], errors="coerce").tolist()
    return pd.DataFrame(rows, index=weeks)


@st.cache_data
def load_q3_center_detail() -> pd.DataFrame:
    raw = pd.read_excel(XLSX_PATH, sheet_name="Q3 2025", header=None)
    start = None
    for i, row in raw.iterrows():
        if str(row.iloc[0]).strip() == "Center" and str(row.iloc[1]).strip() in ("W1", "Wk1"):
            if start is not None:
                start = i
                break
            start = i
    if start is None:
        return pd.DataFrame()
    week_row  = raw.iloc[start]
    weeks_net = [str(week_row.iloc[j]).strip() for j in range(1, 14)]
    centers_q3 = ["Baton Rouge", "Lansing", "Monroe", "Buffalo", "Halifax", "Quebec", "Calgary"]
    rows_out = []
    for offset in range(2, 9):
        r = raw.iloc[start + offset]
        if str(r.iloc[0]).strip() not in centers_q3:
            continue
        rows_out.append([str(r.iloc[0]).strip()] + pd.to_numeric(r.iloc[1:14], errors="coerce").tolist())
    return pd.DataFrame(rows_out, columns=["Center"] + weeks_net)


# -----------------------------------------------------------------------------
# Tabs
# -----------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "\U0001f4cb  Bench Forecast",
    "\U0001f4c8  Historic Bench",
    "\U0001f4ca  Q3 Comparison",
    "\U0001f4e7  Executive Summary",
])


# =============================================================================
# TAB 1 - Editable Bench Forecast
# =============================================================================
with tab1:
    import streamlit.components.v1 as components

    df_forecast = load_bench_forecast()
    if "forecast_data" not in st.session_state:
        st.session_state["forecast_data"] = df_forecast.copy()

    working = st.session_state["forecast_data"].copy()

    if EDITING_LOCKED:
        st.markdown(
            f'<div style="background:#2d1a00;border:1px solid {ACCENT2};border-left:4px solid {ACCENT2};'
            f'border-radius:6px;padding:10px 16px;margin-bottom:12px;font-size:13px;color:{ACCENT2};">'
            f'&#128274; <strong>Forecast table is locked.</strong> '
            f'Editing has been disabled by the administrator. Contact your manager to request changes.'
            f'</div>',
            unsafe_allow_html=True,
        )
        st.markdown(f'<div class="sec-title">Weekly Headcount by Center - read only</div>',
                    unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="sec-title">Weekly Headcount by Center - click any cell to edit</div>',
                    unsafe_allow_html=True)

    # -- Build one combined HTML table with <input> cells -----------------
    def build_editable_table(df: pd.DataFrame) -> str:
        th_style = (
            f"background:{TBL_HEADER_BG};color:{TBL_HEADER_FG};"
            f"font-size:0.75rem;font-weight:700;text-transform:uppercase;"
            f"letter-spacing:0.06em;padding:8px 6px;text-align:center;"
            f"border:1px solid {BORDER};border-bottom:2px solid {ACCENT};"
        )
        th_first = th_style + "text-align:left;padding-left:12px;"
        inp_style = (
            f"width:100%;background:transparent;border:none;color:inherit;"
            f"font-size:0.85rem;text-align:center;padding:0;margin:0;"
            f"-moz-appearance:textfield;"
        )
        td_num_style = (
            f"padding:5px 4px;border:1px solid {BORDER};"
            f"text-align:center;vertical-align:middle;"
        )

        header = "".join(
            f'<th style="{th_first}">Center</th>' +
            "".join(f'<th style="{th_style}">{w}</th>' for w in WEEKS)
        )

        rows_html = ""
        for _, row in df.iterrows():
            center = row["Center"]
            bg, fg = CENTER_ROW_COLORS.get(center, (SURFACE, TEXT_PRI))
            td_center = (
                f'<td style="background:{bg};color:{fg};font-weight:600;'
                f'font-size:0.82rem;padding:6px 12px;border:1px solid {BORDER};'
                f'white-space:nowrap;">{center}</td>'
            )
            cells = td_center
            for w in WEEKS:
                val = int(row[w])
                cid = f"{center.replace(' ', '_')}_{w.replace(' ', '_')}"
                cells += (
                    f'<td style="background:{bg};{td_num_style}">'
                    f'<input id="{cid}" type="number" min="0" max="9999" value="{val}" '
                    f'style="{inp_style}color:{fg};" '
                    f'oninput="recalc()" />'
                    f'</td>'
                )
            rows_html += f"<tr>{cells}</tr>\n"

        # Grand Total row (read-only, calculated by JS)
        total_td = (
            f'<td style="background:{TBL_TOTAL_BG};color:{TBL_TOTAL_FG};'
            f'font-weight:700;font-size:0.82rem;padding:6px 12px;'
            f'border:1px solid {BORDER};">Grand Total</td>'
        )
        for w in WEEKS:
            wid = w.replace(" ", "_")
            total_td += (
                f'<td id="tot_{wid}" style="background:{TBL_TOTAL_BG};'
                f'color:{TBL_TOTAL_FG};font-weight:700;{td_num_style}">0</td>'
            )
        rows_html += f"<tr>{total_td}</tr>\n"

        # Centers list for JS
        centers_js = "[" + ",".join(f'"{c}"' for c in CENTERS) + "]"
        weeks_js   = "[" + ",".join(f'"{w}"' for w in WEEKS) + "]"

        js = f"""
<script>
const CENTERS = {centers_js};
const WEEKS   = {weeks_js};

function getId(c, w) {{
  return c.replace(/ /g,'_') + '_' + w.replace(/ /g,'_');
}}

function recalc() {{
  WEEKS.forEach(w => {{
    let tot = 0;
    CENTERS.forEach(c => {{
      const el = document.getElementById(getId(c, w));
      if (el) tot += (parseInt(el.value) || 0);
    }});
    const td = document.getElementById('tot_' + w.replace(/ /g,'_'));
    if (td) td.textContent = tot;
  }});
  // send data back to Streamlit
  const data = {{}};
  CENTERS.forEach(c => {{
    data[c] = {{}};
    WEEKS.forEach(w => {{
      const el = document.getElementById(getId(c, w));
      data[c][w] = el ? (parseInt(el.value) || 0) : 0;
    }});
  }});
  window.parent.postMessage({{type:'bench_data', payload: data}}, '*');
}}

// run once on load to populate Grand Total
window.addEventListener('load', recalc);
</script>
"""

        return f"""
<style>
  body {{ margin:0; padding:0; background:{BG}; }}
  table {{ border-collapse:collapse; width:100%; font-family:Inter,system-ui,sans-serif; }}
  input[type=number] {{
    outline: none !important;
    box-shadow: none !important;
    -moz-appearance: textfield !important;
  }}
  input[type=number]:focus {{
    outline: none !important;
    box-shadow: none !important;
  }}
  input[type=number]::-webkit-outer-spin-button,
  input[type=number]::-webkit-inner-spin-button {{
    -webkit-appearance: none !important;
    appearance: none !important;
    margin: 0;
    display: none;
  }}
  tr:hover td {{ filter: brightness(1.12); }}
</style>
<div style="overflow-x:auto;border-radius:8px;border:1px solid {BORDER};">
  <table>
    <thead><tr>{header}</tr></thead>
    <tbody>{rows_html}</tbody>
  </table>
</div>
{js}
"""

    if EDITING_LOCKED:
        # Read-only styled table - reuse the existing styled_table helper
        styled_table(working, total_rows=[])
    else:
        table_html = build_editable_table(working)
        # Row height: 40px per center + 44px for grand-total row + 38px header
        components.html(table_html, height=len(CENTERS) * 40 + 82, scrolling=False)

    # Pull everything that follows up to close the iframe's bottom whitespace
    st.markdown("<div style='margin-top:-2rem'></div>", unsafe_allow_html=True)

    edited = working.copy()

    # -- Grand Total + Bench % - flush below the editable table -----------
    grand_total = edited[WEEKS].sum().to_frame().T
    grand_total.insert(0, "Center", "Grand Total")

    # HC input needed before bench_pct is calculated - rendered after tables
    hc = st.session_state.get("hc_input", 2053)

    bench_pct = grand_total[WEEKS].values[0] / hc * 100
    bench_pct_row = pd.DataFrame(
        [["Bench %"] + [f"{v:.1f}%" for v in bench_pct]],
        columns=["Center"] + WEEKS,
    )

    styled_table(grand_total, total_rows=["Grand Total"])
    st.markdown("<div style='margin-top:6px'></div>", unsafe_allow_html=True)
    styled_table(bench_pct_row, total_rows=["Bench %"])

    # -- Total HC input - below the tables ---------------------------------
    st.markdown("<br>", unsafe_allow_html=True)
    hc_col, _ = st.columns([1, 3])
    with hc_col:
        hc = st.number_input(
            "Total HC (Bench % denominator)",
            min_value=1, value=2053, step=1, key="hc_input",
        )
    # Recompute bench_pct with the confirmed hc value
    bench_pct = grand_total[WEEKS].values[0] / hc * 100

    # -- KPI strip ---------------------------------------------------------
    cur_wk_totals = grand_total[WEEKS].values[0]
    peak_wk_idx   = cur_wk_totals.argmax()
    st.markdown("<br>", unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Peak Week",       f"Wk {peak_wk_idx+1:02d}",         f"{int(cur_wk_totals[peak_wk_idx])} heads")
    k2.metric("Wk 01 Total",     f"{int(cur_wk_totals[0])}",         "")
    k3.metric("Wk 13 Total",     f"{int(cur_wk_totals[-1])}",        f"{int(cur_wk_totals[-1]-cur_wk_totals[0]):+d} vs Wk 01")
    k4.metric("Avg Bench %",     f"{bench_pct.mean():.1f}%",         "")

    # -- Filled area chart - one line per center ---------------------------
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f'<div class="sec-title">Weekly Bench by Center</div>', unsafe_allow_html=True)

    fig_prev = go.Figure()

    # Per-center filled area traces
    for center in CENTERS:
        row = edited[edited["Center"] == center]
        if row.empty:
            continue
        color     = CENTER_COLORS.get(center, TEXT_SEC)
        vals      = row[WEEKS].values[0].tolist()
        r, g, b   = int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)
        fill_rgba = f"rgba({r},{g},{b},0.15)"

        fig_prev.add_trace(go.Scatter(
            name=center,
            x=WEEKS,
            y=vals,
            mode="lines+markers",
            line=dict(color=color, width=2.2),
            marker=dict(size=6, color=color, symbol="circle",
                        line=dict(color=SURFACE, width=1.5)),
            fill="tozeroy",
            fillcolor=fill_rgba,
            hovertemplate=f"<b>{center}</b><br>%{{x}}: %{{y}} heads<extra></extra>",
        ))

    # Grand Total line - white, dashed, with data labels
    total_vals = edited[WEEKS].sum().tolist()
    fig_prev.add_trace(go.Scatter(
        name="Grand Total",
        x=WEEKS,
        y=total_vals,
        mode="lines+markers+text",
        line=dict(color="#ffffff", width=2.5, dash="dash"),
        marker=dict(size=7, color="#ffffff", symbol="diamond",
                    line=dict(color=SURFACE, width=1.5)),
        text=[str(int(v)) for v in total_vals],
        textposition="top center",
        textfont=dict(color="#ffffff", size=10, family="Inter, system-ui"),
        hovertemplate="<b>Grand Total</b><br>%{x}: %{y} heads<extra></extra>",
    ))

    layout1 = {**CHART_LAYOUT}
    layout1.update(dict(
        height=400,
        title=dict(text="", x=0),
        xaxis={**CHART_LAYOUT["xaxis"], "title": ""},
        yaxis=dict(**CHART_LAYOUT["yaxis"], title="Head Count"),
        hovermode="x unified",
    ))
    fig_prev.update_layout(**layout1)
    st.plotly_chart(fig_prev, use_container_width=True)

    # -- Download -----------------------------------------------------------
    st.markdown("<br>", unsafe_allow_html=True)

    @st.cache_data
    def to_excel_bytes(df: pd.DataFrame) -> bytes:
        import io
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Bench Forecast")
        return buf.getvalue()

    dl_df = pd.concat([edited, grand_total], ignore_index=True)
    st.download_button(
        label="  Download updated forecast (.xlsx)",
        data=to_excel_bytes(dl_df),
        file_name="NA_Bench_Forecast_updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =============================================================================
# TAB 2 - Historic Bench Trend
# =============================================================================
with tab2:
    st.markdown(f'<div class="sec-title">All Centers . NA FNC Total overlay</div>', unsafe_allow_html=True)

    df_hist = load_historic()
    years_avail    = sorted(df_hist["Year"].unique())
    quarters_avail = sorted(df_hist["Quarter"].unique())
    hist_centers   = ["Baton Rouge", "East Lansing", "Monroe", "Buffalo", "Halifax", "Quebec", "Calgary"]

    fc1, fc2, fc3 = st.columns([1, 1, 2])
    with fc1:
        sel_years    = st.multiselect("Year",             years_avail,    default=years_avail,    key="hist_years")
    with fc2:
        sel_quarters = st.multiselect("Quarter",          quarters_avail, default=quarters_avail, key="hist_quarters")
    with fc3:
        sel_centers  = st.multiselect("Centers to show",  hist_centers,   default=hist_centers,   key="hist_centers")

    df_f = df_hist[
        df_hist["Year"].isin(sel_years) & df_hist["Quarter"].isin(sel_quarters)
    ].sort_values(["Year", "Quarter", "Week"]).copy()

    if df_f.empty:
        st.info("No data for the selected filters.")
    else:
        x_labels = df_f["Label"].tolist()

        # -- KPI strip -----------------------------------------------------
        avg_cic = df_f["NA FNC"].mean()
        max_cic = df_f["NA FNC"].max()
        min_cic = df_f["NA FNC"].min()
        last_cic = df_f["NA FNC"].iloc[-1]

        st.markdown("<br>", unsafe_allow_html=True)
        h1, h2, h3, h4 = st.columns(4)
        h1.metric("Avg NA FNC",        f"{avg_cic:.0f}",    "")
        h2.metric("Peak NA FNC",        f"{int(max_cic)}",   "")
        h3.metric("Trough NA FNC",      f"{int(min_cic)}",   "")
        h4.metric("Latest Data Point",  f"{int(last_cic)}",  "")
        st.markdown("<br>", unsafe_allow_html=True)

        # -- Chart ---------------------------------------------------------
        fig_hist = go.Figure()

        for center in sel_centers:
            if center not in df_f.columns:
                continue
            fig_hist.add_trace(go.Bar(
                name=center,
                x=x_labels,
                y=df_f[center].tolist(),
                marker_color=CENTER_COLORS.get(center, TEXT_SEC),
                marker_line_width=0,
                hovertemplate=f"<b>{center}</b><br>%{{x}}: %{{y}}<extra></extra>",
            ))

        # NA FNC total line - amber, stands out on dark stacked bars
        fig_hist.add_trace(go.Scatter(
            name="NA FNC Total",
            x=x_labels,
            y=df_f["NA FNC"].tolist(),
            mode="lines+markers+text",
            line=dict(color=ACCENT2, width=2.5),
            marker=dict(size=4, color=ACCENT2, symbol="circle"),
            text=[str(int(v)) for v in df_f["NA FNC"].tolist()],
            textposition="top center",
            textfont=dict(color=ACCENT2, size=9, family="Inter, system-ui"),
            hovertemplate="<b>NA FNC Total</b><br>%{x}: %{y}<extra></extra>",
        ))

        # Quarter boundary dotted lines
        prev_q = None
        for i, (_, row) in enumerate(df_f.iterrows()):
            cur_q = (row["Year"], row["Quarter"])
            if prev_q and cur_q != prev_q:
                fig_hist.add_vline(
                    x=i - 0.5,
                    line_width=1,
                    line_dash="dot",
                    line_color=BORDER,
                )
            prev_q = cur_q

        layout2 = {**CHART_LAYOUT}
        layout2.update(dict(
            barmode="stack",
            height=480,
            bargap=0.06,
            title=dict(
                text="Historic Bench - Stacked by Center  .  NA FNC Total (amber line)",
                font=dict(size=13, color=TEXT_SEC),
                x=0,
            ),
            xaxis={**CHART_LAYOUT["xaxis"], "tickangle": -55, "tickfont": dict(size=9, color=TEXT_SEC)},
            yaxis=dict(**CHART_LAYOUT["yaxis"], title="Head Count on Bench"),
            legend={**CHART_LAYOUT["legend"], "y": -0.24},
        ))
        fig_hist.update_layout(**layout2)
        st.plotly_chart(fig_hist, use_container_width=True)


# =============================================================================
# TAB 3 - Q3 2025 vs Q3 2026 Comparison
# =============================================================================
with tab3:
    st.markdown(f'<div class="sec-title">Q3 26 Forecast . Q3 25 Actual . Q3 26 Original Forecast</div>', unsafe_allow_html=True)

    df_q3 = load_q3_comparison()

    if df_q3.empty:
        st.warning("Could not load Q3 comparison data.")
    else:
        show_detail = st.checkbox("Show per-center Net bench detail", value=False)

        # -- KPI strip -----------------------------------------------------
        if "Q3 26 Forecast" in df_q3.columns and "Q3 25 bench" in df_q3.columns:
            delta_last = int(df_q3["Q3 26 Forecast"].iloc[-1] - df_q3["Q3 25 bench"].iloc[-1])
            st.markdown("<br>", unsafe_allow_html=True)
            q1, q2, q3_col, q4 = st.columns(4)
            q1.metric("Q3 26 Forecast peak",   f"{int(df_q3['Q3 26 Forecast'].max())}")
            q2.metric("Q3 25 Bench trough",     f"{int(df_q3['Q3 25 bench'].min())}")
            q3_col.metric("Wk 13 Delta (26 vs 25)", f"{delta_last:+d}")
            q4.metric("Q3 Orig Forecast peak",  f"{int(df_q3['Q3 Original Forecast'].max())}")
            st.markdown("<br>", unsafe_allow_html=True)

        # -- Chart ---------------------------------------------------------
        fig_q3 = go.Figure()

        series_cfg = {
            "Q3 26 Forecast":       dict(color=C_FORECAST, dash="solid", width=2.8, symbol="circle",  size=8),
            "Q3 25 bench":          dict(color=C_ACTUAL,   dash="solid", width=2.8, symbol="square",  size=8),
            "Q3 Original Forecast": dict(color=C_ORIG,     dash="dash",  width=1.8, symbol="diamond", size=6),
        }

        for col in df_q3.columns:
            cfg = series_cfg.get(col, dict(color=TEXT_SEC, dash="solid", width=1.5, symbol="circle", size=6))
            fig_q3.add_trace(go.Scatter(
                name=col,
                x=df_q3.index.tolist(),
                y=df_q3[col].tolist(),
                mode="lines+markers+text",
                line=dict(color=cfg["color"], width=cfg["width"], dash=cfg["dash"]),
                marker=dict(size=cfg["size"], color=cfg["color"], symbol=cfg["symbol"],
                            line=dict(color=SURFACE, width=1.5)),
                text=[str(int(v)) for v in df_q3[col].tolist()],
                textposition="top center",
                textfont=dict(color=cfg["color"], size=10, family="Inter, system-ui"),
                hovertemplate=f"<b>{col}</b><br>%{{x}}: %{{y}} heads<extra></extra>",
            ))

        # Shaded band between Q3 26 forecast and Q3 25 bench
        if "Q3 26 Forecast" in df_q3.columns and "Q3 25 bench" in df_q3.columns:
            x_vals = df_q3.index.tolist()
            upper  = df_q3["Q3 26 Forecast"].tolist()
            lower  = df_q3["Q3 25 bench"].tolist()
            fig_q3.add_trace(go.Scatter(
                x=x_vals + x_vals[::-1],
                y=upper + lower[::-1],
                fill="toself",
                fillcolor="rgba(56,189,248,0.06)",
                line=dict(color="rgba(0,0,0,0)"),
                showlegend=False,
                hoverinfo="skip",
            ))

        layout3 = {**CHART_LAYOUT}
        layout3.update(dict(
            height=440,
            title=dict(
                text="Q3 Total Bench Comparison - 13-week view",
                font=dict(size=13, color=TEXT_SEC),
                x=0,
            ),
            xaxis={**CHART_LAYOUT["xaxis"], "title": "Week", "tickangle": -30},
            yaxis=dict(**CHART_LAYOUT["yaxis"], title="Total Bench Headcount"),
            legend={**CHART_LAYOUT["legend"], "y": -0.18},
        ))
        fig_q3.update_layout(**layout3)
        st.plotly_chart(fig_q3, use_container_width=True)

        # -- Per-center stacked detail (optional) --------------------------
        if show_detail:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(f'<div class="sec-title">Per-Center Q3 2025 Net Bench by Week</div>', unsafe_allow_html=True)
            df_detail = load_q3_center_detail()
            if not df_detail.empty:
                fig_det = go.Figure()
                for _, row_d in df_detail.iterrows():
                    cname = row_d["Center"]
                    fig_det.add_trace(go.Bar(
                        name=cname,
                        x=df_detail.columns[1:].tolist(),
                        y=row_d[1:].tolist(),
                        marker_color=CENTER_COLORS.get(cname, TEXT_SEC),
                        marker_line_width=0,
                    ))
                layout_det = {**CHART_LAYOUT}
                layout_det.update(dict(
                    barmode="stack",
                    height=330,
                    bargap=0.2,
                    xaxis={**CHART_LAYOUT["xaxis"], "title": "Week"},
                    yaxis=dict(**CHART_LAYOUT["yaxis"], title="Net Bench"),
                    legend={**CHART_LAYOUT["legend"], "y": -0.26},
                ))
                fig_det.update_layout(**layout_det)
                st.plotly_chart(fig_det, use_container_width=True)
                styled_table(df_detail)


# =============================================================================
# TAB 4 - Executive Summary
# =============================================================================
with tab4:
    st.markdown(f'<div class="sec-title">Executive Summary - Dynamic Email Draft</div>',
                unsafe_allow_html=True)
    st.caption("All highlighted figures update automatically as you edit the forecast data.")

    # -- Pull data from session state / already-computed values ------------
    _fc   = load_bench_forecast()
    _hc   = st.session_state.get("hc_input", 2053)
    _work = st.session_state.get("forecast_data", _fc).copy()

    # Per-center totals over all weeks
    _gt_vals  = _work[WEEKS].sum(axis=0).tolist()          # grand total per week
    _gt_w1    = int(_gt_vals[0])
    _gt_w13   = int(_gt_vals[-1])
    _peak_idx = int(pd.Series(_gt_vals).argmax())
    _peak_wk  = WEEKS[_peak_idx]
    _peak_val = int(_gt_vals[_peak_idx])
    _avg_pct  = float(pd.Series(_gt_vals).mean() / _hc * 100)
    _w1_pct   = _gt_w1 / _hc * 100
    _w13_pct  = _gt_w13 / _hc * 100
    _trend    = "increasing" if _gt_w13 > _gt_w1 else "decreasing" if _gt_w13 < _gt_w1 else "stable"

    # Highest and lowest centers at Wk 01
    _w1_series = {c: int(_work[_work["Center"] == c][WEEKS[0]].values[0])
                  for c in CENTERS if not _work[_work["Center"] == c].empty}
    _top_center    = max(_w1_series, key=_w1_series.get)
    _top_center_v  = _w1_series[_top_center]
    _low_center    = min(_w1_series, key=_w1_series.get)
    _low_center_v  = _w1_series[_low_center]

    # Q3 comparison data
    _q3 = load_q3_comparison()
    _q3_fc_w1  = int(_q3["Q3 26 Forecast"].iloc[0])  if not _q3.empty else 0
    _q3_ac_w1  = int(_q3["Q3 25 bench"].iloc[0])     if not _q3.empty else 0
    _q3_delta  = _q3_fc_w1 - _q3_ac_w1
    _q3_trend  = "above" if _q3_delta > 0 else "below" if _q3_delta < 0 else "in line with"
    _q3_fc_pk  = int(_q3["Q3 26 Forecast"].max())    if not _q3.empty else 0
    _q3_ac_trg = int(_q3["Q3 25 bench"].min())       if not _q3.empty else 0

    # Historic data
    _hist = load_historic()
    _hist_last = _hist.iloc[-1] if not _hist.empty else None
    _hist_lbl  = str(_hist_last["Label"]) if _hist_last is not None else "latest"
    _hist_fnc  = int(_hist_last["NA FNC"]) if _hist_last is not None else 0

    # -- Email meta controls -----------------------------------------------
    st.markdown("<br>", unsafe_allow_html=True)
    em1, em2, em3 = st.columns(3)
    with em1:
        sender_name  = st.text_input("From (your name)",    value="Workforce Planning Team", key="em_from")
        report_week  = st.text_input("Reporting week label", value=WEEKS[0],                 key="em_week")
    with em2:
        quarter_label = st.text_input("Quarter",    value="Q3 2026",    key="em_qtr")
        region_label  = st.text_input("Region",     value="NA FNC",     key="em_region")
    with em3:
        bench_target  = st.number_input("Bench % target (%)", min_value=1.0, max_value=30.0,
                                         value=8.5, step=0.5, key="em_target")
        tone          = st.selectbox("Tone", ["Professional", "Executive Summary", "Detailed"],
                                     key="em_tone")

    # -- Dynamic narrative engine -------------------------------------------
    def _hl(text: str, color: str = ACCENT) -> str:
        """Wrap text in a colored highlight span."""
        return f'<span style="color:{color};font-weight:700">{text}</span>'

    def _status_color(val: float, target: float) -> str:
        if val > target * 1.15: return DANGER
        if val > target:        return WARN
        return GOOD

    avg_color   = _status_color(_avg_pct, bench_target)
    w1_color    = _status_color(_w1_pct,  bench_target)
    w13_color   = _status_color(_w13_pct, bench_target)
    trend_color = GOOD if _trend == "decreasing" else WARN if _trend == "increasing" else ACCENT

    # Build paragraphs based on tone
    if tone == "Executive Summary":
        p_opening = (
            f"Please find below the {_hl(quarter_label)} Bench Forecast summary for "
            f"{_hl(region_label)}. As of {_hl(report_week)}, total bench stands at "
            f"{_hl(str(_gt_w1), w1_color)} resources, representing "
            f"{_hl(f'{_w1_pct:.1f}%', w1_color)} of total HC ({_hl(str(_hc))})."
        )
        p_trend = (
            f"The bench is {_hl(_trend, trend_color)} across the quarter, "
            f"projecting {_hl(str(_gt_w13), w13_color)} resources ({_hl(f'{_w13_pct:.1f}%', w13_color)}) "
            f"by {_hl(WEEKS[-1])}. Peak bench is forecasted at {_hl(_peak_wk)} "
            f"with {_hl(str(_peak_val))} resources."
        )
        p_centers = (
            f"{_hl(_top_center, CENTER_COLORS.get(_top_center, ACCENT))} leads with "
            f"{_hl(str(_top_center_v))} resources on bench in {_hl(report_week)}, "
            f"while {_hl(_low_center, CENTER_COLORS.get(_low_center, ACCENT))} "
            f"has the lowest count at {_hl(str(_low_center_v))}."
        )
        p_yoy = (
            f"Compared to {_hl('Q3 2025 actuals', C_ACTUAL)}, the current forecast is "
            f"{_hl(_q3_trend + ' Q3 2025', WARN if _q3_delta > 0 else GOOD)} "
            f"by {_hl(str(abs(_q3_delta)))} resources at week 1 "
            f"(Forecast: {_hl(str(_q3_fc_w1), C_FORECAST)} vs "
            f"Actual: {_hl(str(_q3_ac_w1), C_ACTUAL)})."
        )
        p_action = (
            f"Average bench rate of {_hl(f'{_avg_pct:.1f}%', avg_color)} is "
            + (f"{_hl('above', DANGER)} the {_hl(f'{bench_target:.1f}%')} target - "
               f"immediate placement actions are recommended."
               if _avg_pct > bench_target else
               f"{_hl('within', GOOD)} the {_hl(f'{bench_target:.1f}%')} target range.")
        )
        paragraphs = [p_opening, p_trend, p_centers, p_yoy, p_action]

    elif tone == "Detailed":
        p_opening = (
            f"Team, please find the detailed {_hl(quarter_label)} Bench Forecast report for "
            f"{_hl(region_label)}. This report covers weeks {_hl(WEEKS[0])} through "
            f"{_hl(WEEKS[-1])} and includes per-center breakdowns, YoY comparisons, "
            f"and trend analysis."
        )
        p_snapshot = (
            f"<b>Current Snapshot ({_hl(report_week)}):</b> Total bench = "
            f"{_hl(str(_gt_w1), w1_color)} ({_hl(f'{_w1_pct:.1f}%', w1_color)} of "
            f"{_hl(str(_hc))} HC). Target bench rate: {_hl(f'{bench_target:.1f}%')}. "
            f"Status: {_hl('ON TARGET' if _w1_pct <= bench_target else 'ABOVE TARGET', GOOD if _w1_pct <= bench_target else DANGER)}."
        )
        p_trend = (
            f"<b>Trend:</b> Bench is {_hl(_trend, trend_color)} from "
            f"{_hl(str(_gt_w1))} in {_hl(WEEKS[0])} to "
            f"{_hl(str(_gt_w13), w13_color)} in {_hl(WEEKS[-1])} "
            f"(net change: {_hl(f'{_gt_w13 - _gt_w1:+d}', GOOD if _gt_w13 < _gt_w1 else WARN)}). "
            f"Peak forecasted at {_hl(_peak_wk)} with {_hl(str(_peak_val))} resources "
            f"({_hl(f'{_peak_val/_hc*100:.1f}%', _status_color(_peak_val/_hc*100, bench_target))})."
        )
        _center_lines = " | ".join(
            f"{_hl(c, CENTER_COLORS.get(c, ACCENT))}: {_hl(str(_w1_series.get(c, 0)))}"
            for c in CENTERS
        )
        p_centers = f"<b>Per-Center ({_hl(WEEKS[0])}):</b> {_center_lines}."
        p_yoy = (
            f"<b>YoY Comparison:</b> Q3 2026 forecast ({_hl(str(_q3_fc_pk), C_FORECAST)} peak) vs "
            f"Q3 2025 actual ({_hl(str(_q3_ac_trg), C_ACTUAL)} trough). "
            f"Week 1 delta: {_hl(f'{_q3_delta:+d}', GOOD if _q3_delta < 0 else WARN)} resources. "
            f"Historic latest reading ({_hl(_hist_lbl)}): {_hl(str(_hist_fnc))} NA FNC total."
        )
        p_action = (
            f"<b>Recommended Actions:</b> "
            + (f"With bench at {_hl(f'{_avg_pct:.1f}%', avg_color)} avg, "
               f"prioritise placement of {_hl(_top_center, CENTER_COLORS.get(_top_center, ACCENT))} "
               f"resources ({_hl(str(_top_center_v))} on bench). "
               f"Monitor {_hl(_peak_wk)} closely as the projected peak period."
               if _avg_pct > bench_target else
               f"Bench rate is healthy. Continue monitoring {_hl(_top_center, CENTER_COLORS.get(_top_center, ACCENT))} "
               f"as the highest-volume center ({_hl(str(_top_center_v))}) and track trajectory to {_hl(WEEKS[-1])}.")
        )
        paragraphs = [p_opening, p_snapshot, p_trend, p_centers, p_yoy, p_action]

    else:  # Professional (default)
        p_opening = (
            f"Please find the {_hl(quarter_label)} Bench Forecast update for "
            f"{_hl(region_label)}. Current bench as of {_hl(report_week)}: "
            f"{_hl(str(_gt_w1), w1_color)} resources "
            f"({_hl(f'{_w1_pct:.1f}%', w1_color)} bench rate)."
        )
        p_trend = (
            f"The forecast shows a {_hl(_trend, trend_color)} trend through {_hl(WEEKS[-1])}, "
            f"with a projected peak of {_hl(str(_peak_val))} in {_hl(_peak_wk)}. "
            f"End-of-quarter bench is projected at {_hl(str(_gt_w13), w13_color)} "
            f"({_hl(f'{_w13_pct:.1f}%', w13_color)})."
        )
        p_centers = (
            f"{_hl(_top_center, CENTER_COLORS.get(_top_center, ACCENT))} continues to carry "
            f"the highest bench volume this week at {_hl(str(_top_center_v))} resources. "
            f"{_hl(_low_center, CENTER_COLORS.get(_low_center, ACCENT))} remains the "
            f"lowest at {_hl(str(_low_center_v))}."
        )
        p_yoy = (
            f"Versus Q3 2025, the current forecast is {_hl(_q3_trend + ' last year', WARN if _q3_delta > 0 else GOOD)} "
            f"by {_hl(str(abs(_q3_delta)))} heads at week 1 "
            f"({_hl(str(_q3_fc_w1), C_FORECAST)} forecast vs {_hl(str(_q3_ac_w1), C_ACTUAL)} actual)."
        )
        p_action = (
            f"Average bench rate of {_hl(f'{_avg_pct:.1f}%', avg_color)} "
            + (f"exceeds the {_hl(f'{bench_target:.1f}%')} target. "
               f"Placement focus is recommended for {_hl(_top_center, CENTER_COLORS.get(_top_center, ACCENT))}."
               if _avg_pct > bench_target else
               f"is within the {_hl(f'{bench_target:.1f}%')} target. No immediate action required.")
        )
        paragraphs = [p_opening, p_trend, p_centers, p_yoy, p_action]

    # -- Render the email preview -------------------------------------------
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f'<div class="sec-title">Preview</div>', unsafe_allow_html=True)

    # Pre-build all dynamic fragments before the single f-string
    para_html = "".join(
        f"<p style='margin:0 0 14px;line-height:1.75'>{p}</p>"
        for p in paragraphs
    )

    yoy_pill_color = '#f87171' if _q3_delta > 0 else '#34d399'

    def _center_row(c: str) -> str:
        bg, fg   = CENTER_ROW_COLORS.get(c, (SURFACE, TEXT_PRI))
        row_src  = _work[_work["Center"] == c]
        v1       = int(row_src[WEEKS[0]].values[0])  if not row_src.empty else 0
        v13      = int(row_src[WEEKS[-1]].values[0]) if not row_src.empty else 0
        if v13 < v1:
            trend_html = '<span style="color:#34d399">&#9660; Down</span>'
        elif v13 > v1:
            trend_html = '<span style="color:#fb923c">&#9650; Up</span>'
        else:
            trend_html = '<span style="color:#8b949e">&#8212; Flat</span>'
        return (
            f'<tr>'
            f'<td style="padding:5px 10px;border:1px solid {BORDER};background:{bg};color:{fg};font-weight:600">{c}</td>'
            f'<td style="text-align:center;padding:5px 10px;border:1px solid {BORDER};background:{bg};color:{fg}">{v1}</td>'
            f'<td style="text-align:center;padding:5px 10px;border:1px solid {BORDER};background:{bg};color:{fg}">{v13}</td>'
            f'<td style="text-align:center;padding:5px 10px;border:1px solid {BORDER};background:{bg}">{trend_html}</td>'
            f'</tr>'
        )

    center_rows_html = "".join(_center_row(c) for c in CENTERS)

    email_html = (
        f'<div style="background:{SURFACE};border:1px solid {BORDER};border-radius:10px;'
        f'padding:32px 36px;font-family:Inter,Segoe UI,system-ui,sans-serif;'
        f'color:{TEXT_PRI};font-size:0.93rem;line-height:1.7;max-width:100%;">'

        # Header
        f'<div style="border-bottom:2px solid {ACCENT};padding-bottom:16px;margin-bottom:24px;">'
        f'<div style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.1em;color:{TEXT_SEC};margin-bottom:6px;">'
        f'{region_label} Workforce Communication</div>'
        f'<div style="font-size:1.25rem;font-weight:700;color:{TEXT_PRI}">'
        f'{quarter_label} Bench Forecast Update &#8212; {report_week}</div>'
        f'<div style="font-size:0.8rem;color:{TEXT_SEC};margin-top:4px">'
        f'From: {sender_name} &nbsp;&middot;&nbsp; Report generated automatically from dashboard data</div>'
        f'</div>'

        # KPI pills
        f'<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:22px;">'
        f'<div style="background:{SURFACE2};border:1px solid {BORDER};border-radius:20px;padding:5px 14px;font-size:0.8rem;">'
        f'&#128205; <b>{report_week}</b> bench: <span style="color:{w1_color};font-weight:700">{_gt_w1} ({_w1_pct:.1f}%)</span></div>'
        f'<div style="background:{SURFACE2};border:1px solid {BORDER};border-radius:20px;padding:5px 14px;font-size:0.8rem;">'
        f'&#128200; Peak: <span style="color:{ACCENT};font-weight:700">{_peak_val} @ {_peak_wk}</span></div>'
        f'<div style="background:{SURFACE2};border:1px solid {BORDER};border-radius:20px;padding:5px 14px;font-size:0.8rem;">'
        f'&#127919; Target: <span style="color:{ACCENT};font-weight:700">{bench_target:.1f}%</span></div>'
        f'<div style="background:{SURFACE2};border:1px solid {BORDER};border-radius:20px;padding:5px 14px;font-size:0.8rem;">'
        f'&#128202; YoY &#916; W1: <span style="color:{yoy_pill_color};font-weight:700">{_q3_delta:+d}</span></div>'
        f'<div style="background:{SURFACE2};border:1px solid {BORDER};border-radius:20px;padding:5px 14px;font-size:0.8rem;">'
        f'&#128201; Avg bench rate: <span style="color:{avg_color};font-weight:700">{_avg_pct:.1f}%</span></div>'
        f'</div>'

        # Body paragraphs
        f'{para_html}'

        # Center snapshot table
        f'<div style="margin-top:20px;border-top:1px solid {BORDER};padding-top:16px;">'
        f'<div style="font-size:0.72rem;text-transform:uppercase;letter-spacing:0.08em;'
        f'color:{TEXT_SEC};margin-bottom:10px;font-weight:700">Center Snapshot &#8212; {report_week}</div>'
        f'<table style="border-collapse:collapse;width:100%;font-size:0.82rem;">'
        f'<thead><tr>'
        f'<th style="text-align:left;padding:5px 10px;background:{TBL_HEADER_BG};color:{TBL_HEADER_FG};border:1px solid {BORDER};">Center</th>'
        f'<th style="text-align:center;padding:5px 10px;background:{TBL_HEADER_BG};color:{TBL_HEADER_FG};border:1px solid {BORDER};">Bench (Wk 01)</th>'
        f'<th style="text-align:center;padding:5px 10px;background:{TBL_HEADER_BG};color:{TBL_HEADER_FG};border:1px solid {BORDER};">Bench (Wk 13)</th>'
        f'<th style="text-align:center;padding:5px 10px;background:{TBL_HEADER_BG};color:{TBL_HEADER_FG};border:1px solid {BORDER};">Trend</th>'
        f'</tr></thead>'
        f'<tbody>{center_rows_html}</tbody>'
        f'</table></div>'

        # Footer
        f'<div style="margin-top:24px;border-top:1px solid {BORDER};padding-top:12px;'
        f'font-size:0.75rem;color:{TEXT_SEC};">'
        f'This communication was generated automatically from the NA Bench Forecast Dashboard. '
        f'Data reflects the latest forecast inputs. For questions, contact {sender_name}.</div>'
        f'</div>'
    )

    st.markdown(email_html, unsafe_allow_html=True)

    # -- Plain-text copy version --------------------------------------------
    import re, json
    st.markdown("<br>", unsafe_allow_html=True)

    def _strip_html(s: str) -> str:
        return re.sub(r"<[^>]+>", "", s)

    plain_paras = [_strip_html(p) for p in paragraphs]

    FONT = "Aptos, Aptos Display, Calibri, sans-serif"

    # -- SVG helper --------------------------------------------------------
    def _svg_grid_and_labels(w, h, pad, vals, x_labels, max_val):
        """Return (x_fn, y_fn, grid_svg) for a line chart area."""
        iw = w - pad["l"] - pad["r"]
        ih = h - pad["t"] - pad["b"]
        n  = len(vals)

        def xf(i):  return pad["l"] + i * iw / max(n - 1, 1)
        def yf(v):  return pad["t"] + ih - (v / max_val * ih)

        grid = ""
        step = max(1, int(max_val // 4))
        for gi in range(0, int(max_val) + 1, step):
            gy = yf(gi)
            grid += (f'<line x1="{pad["l"]}" y1="{gy:.1f}" x2="{pad["l"]+iw}" y2="{gy:.1f}" '
                     f'stroke="#30363d" stroke-width="1"/>'
                     f'<text x="{pad["l"]-5}" y="{gy+4:.1f}" text-anchor="end" '
                     f'font-size="9" fill="#8b949e" font-family="{FONT}">{gi}</text>')
        for i, lbl in enumerate(x_labels):
            grid += (f'<text x="{xf(i):.1f}" y="{h-5}" text-anchor="middle" '
                     f'font-size="9" fill="#8b949e" font-family="{FONT}">{lbl}</text>')
        return xf, yf, grid

    # -- Chart 1: Q3 Comparison - 3-line chart ----------------------------
    q3_series = {
        "Q3 26 Forecast":       (C_FORECAST, list(_q3["Q3 26 Forecast"])       if not _q3.empty else []),
        "Q3 25 Bench":          (C_ACTUAL,   list(_q3["Q3 25 bench"])           if not _q3.empty else []),
        "Q3 Orig Forecast":     (C_ORIG,     list(_q3["Q3 Original Forecast"])  if not _q3.empty else []),
    }
    q3_wk_labels = list(_q3.index) if not _q3.empty else WEEKS

    q3_all_vals = [v for _, vals in q3_series.values() for v in vals if v == v]
    q3_max      = max(q3_all_vals) * 1.12 if q3_all_vals else 300
    q3_w, q3_h  = 560, 200
    q3_pad      = dict(l=40, r=20, t=24, b=32)

    xf3, yf3, q3_grid = _svg_grid_and_labels(
        q3_w, q3_h, q3_pad, list(range(len(q3_wk_labels))), q3_wk_labels, q3_max
    )

    q3_lines = q3_grid
    legend_x  = q3_pad["l"]
    for idx, (name, (col, vals)) in enumerate(q3_series.items()):
        if not vals:
            continue
        pts = " ".join(f"{xf3(i):.1f},{yf3(v):.1f}" for i, v in enumerate(vals))
        dash = "6,3" if name == "Q3 Orig Forecast" else "none"
        q3_lines += (f'<polyline points="{pts}" fill="none" stroke="{col}" '
                     f'stroke-width="2.2" stroke-dasharray="{dash}" stroke-linejoin="round"/>')
        for i, v in enumerate(vals):
            cx, cy = xf3(i), yf3(v)
            q3_lines += (f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="3.5" fill="{col}" '
                         f'stroke="#161b22" stroke-width="1.5"/>')
            q3_lines += (f'<text x="{cx:.1f}" y="{cy-8:.1f}" text-anchor="middle" '
                         f'font-size="9" fill="{col}" font-weight="bold" '
                         f'font-family="{FONT}">{int(v)}</text>')
        # Legend dot + label
        lx = legend_x + idx * 185
        q3_lines += (f'<circle cx="{lx+6}" cy="10" r="4" fill="{col}"/>'
                     f'<text x="{lx+14}" y="14" font-size="9" fill="{col}" '
                     f'font-family="{FONT}" font-weight="600">{name}</text>')

    q3_chart_svg = (
        f'<svg width="{q3_w}" height="{q3_h}" xmlns="http://www.w3.org/2000/svg">'
        f'<rect width="{q3_w}" height="{q3_h}" fill="#161b22" rx="6"/>'
        f'{q3_lines}'
        f'</svg>'
    )

    # -- Chart 2: Historic Bench - stacked by center + NA FNC amber line ---
    _hist_data   = load_historic()
    _hb_centers  = ["Baton Rouge", "East Lansing", "Monroe", "Buffalo", "Halifax", "Quebec", "Calgary"]
    # Keep only center columns that actually exist in the dataframe
    _hb_centers  = [c for c in _hb_centers if c in _hist_data.columns]

    hb_w, hb_h   = 560, 240
    hb_pad       = dict(l=45, r=20, t=28, b=52)
    hb_iw        = hb_w - hb_pad["l"] - hb_pad["r"]
    hb_ih        = hb_h - hb_pad["t"] - hb_pad["b"]

    hb_svg = (f'<svg width="{hb_w}" height="{hb_h}" xmlns="http://www.w3.org/2000/svg">'
              f'<rect width="{hb_w}" height="{hb_h}" fill="#161b22" rx="6"/>')

    if not _hist_data.empty and _hb_centers:
        _hb_df    = _hist_data.sort_values(["Year", "Quarter", "Week"]).copy()
        hb_labels = _hb_df["Label"].tolist()
        hb_totals = _hb_df["NA FNC"].tolist()
        n_hb      = len(hb_labels)
        hb_max    = max(hb_totals) * 1.15 if hb_totals else 1
        gap       = hb_iw / n_hb
        bar_w     = gap * 0.78

        def hb_xf(i):  return hb_pad["l"] + i * gap + gap / 2
        def hb_yf(v):  return hb_pad["t"] + hb_ih - (v / hb_max * hb_ih)

        # Y gridlines
        _hb_step = max(1, int(hb_max // 5))
        for gi in range(0, int(hb_max) + 1, _hb_step):
            gy = hb_yf(gi)
            hb_svg += (f'<line x1="{hb_pad["l"]}" y1="{gy:.1f}" '
                       f'x2="{hb_pad["l"]+hb_iw}" y2="{gy:.1f}" '
                       f'stroke="#30363d" stroke-width="1"/>'
                       f'<text x="{hb_pad["l"]-5}" y="{gy+4:.1f}" text-anchor="end" '
                       f'font-size="9" fill="#8b949e" font-family="{FONT}">{gi}</text>')

        # Stacked bars - one segment per center per x position
        for i, row in enumerate(_hb_df.itertuples()):
            bx      = hb_xf(i) - bar_w / 2
            bottom  = 0.0
            for center in _hb_centers:
                val = float(getattr(row, center.replace(" ", "_"), 0) or 0)
                if val <= 0:
                    continue
                seg_h = val / hb_max * hb_ih
                seg_y = hb_yf(bottom + val)
                col   = CENTER_COLORS.get(center, TEXT_SEC)
                hb_svg += (f'<rect x="{bx:.1f}" y="{seg_y:.1f}" '
                           f'width="{bar_w:.1f}" height="{seg_h:.1f}" '
                           f'fill="{col}" opacity="0.85"/>')
                bottom += val

        # Quarter boundary dotted lines + x-axis labels (every other tick if crowded)
        prev_q = None
        for i, row in enumerate(_hb_df.itertuples()):
            cur_q = (row.Year, row.Quarter)
            if prev_q and cur_q != prev_q:
                lx2 = hb_pad["l"] + i * gap
                hb_svg += (f'<line x1="{lx2:.1f}" y1="{hb_pad["t"]}" '
                           f'x2="{lx2:.1f}" y2="{hb_pad["t"]+hb_ih}" '
                           f'stroke="#30363d" stroke-width="1" stroke-dasharray="4,2"/>')
            if n_hb <= 16 or i % 2 == 0:
                tx = hb_xf(i)
                ty = hb_h - 4
                hb_svg += (f'<text x="{tx:.1f}" y="{ty}" text-anchor="end" '
                           f'font-size="8" fill="#8b949e" font-family="{FONT}" '
                           f'transform="rotate(-45,{tx:.1f},{ty})">'
                           f'{hb_labels[i]}</text>')
            prev_q = cur_q

        # NA FNC amber overlay line with data labels
        line_pts = " ".join(f"{hb_xf(i):.1f},{hb_yf(v):.1f}" for i, v in enumerate(hb_totals))
        hb_svg += (f'<polyline points="{line_pts}" fill="none" stroke="{ACCENT2}" '
                   f'stroke-width="2.2" stroke-linejoin="round"/>')
        for i, v in enumerate(hb_totals):
            cx, cy = hb_xf(i), hb_yf(v)
            hb_svg += (f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="3" fill="{ACCENT2}" '
                       f'stroke="#161b22" stroke-width="1.2"/>')
            # Show label on every other point to avoid crowding
            if n_hb <= 16 or i % 2 == 0:
                hb_svg += (f'<text x="{cx:.1f}" y="{cy-7:.1f}" text-anchor="middle" '
                           f'font-size="8" fill="{ACCENT2}" font-weight="bold" '
                           f'font-family="{FONT}">{int(v)}</text>')

        # Legend - centers + amber total line
        _leg_items = [(c, CENTER_COLORS.get(c, TEXT_SEC), "rect") for c in _hb_centers]
        _leg_items += [("NA FNC Total", ACCENT2, "line")]
        _leg_cols  = 4
        for li, (name, col, shape) in enumerate(_leg_items):
            lx3 = hb_pad["l"] + (li % _leg_cols) * 128
            ly3 = hb_pad["t"] + hb_ih + 30 + (li // _leg_cols) * 14
            if shape == "rect":
                hb_svg += (f'<rect x="{lx3}" y="{ly3-7}" width="10" height="8" '
                           f'fill="{col}" rx="1" opacity="0.85"/>')
            else:
                hb_svg += (f'<line x1="{lx3}" y1="{ly3-3}" x2="{lx3+10}" y2="{ly3-3}" '
                           f'stroke="{col}" stroke-width="2"/>'
                           f'<circle cx="{lx3+5}" cy="{ly3-3}" r="2.5" fill="{col}"/>')
            hb_svg += (f'<text x="{lx3+14}" y="{ly3}" font-size="8" fill="{col}" '
                       f'font-family="{FONT}">{name}</text>')

    hb_svg += '</svg>'
    hist_bench_svg = hb_svg

    # -- KPI pill HTML -----------------------------------------------------
    def _pill(label, val, color):
        return (
            f'<div style="display:inline-block;background:#21262d;border:1px solid #30363d;'
            f'border-radius:20px;padding:5px 14px;font-size:12px;margin:3px;">'
            f'{label}: <strong style="color:{color}">{val}</strong></div>'
        )

    pills_html = (
        _pill("Bench " + report_week, f"{_gt_w1} ({_w1_pct:.1f}%)", w1_color) +
        _pill("Peak", f"{_peak_val} @ {_peak_wk}", ACCENT) +
        _pill("Target", f"{bench_target:.1f}%", ACCENT) +
        _pill("YoY W1 &#916;", f"{_q3_delta:+d}", '#f87171' if _q3_delta > 0 else '#34d399') +
        _pill("Avg Rate", f"{_avg_pct:.1f}%", avg_color)
    )

    # -- Center snapshot table rows ----------------------------------------
    snap_rows = ""
    for c in CENTERS:
        bg, fg = CENTER_ROW_COLORS.get(c, (SURFACE, TEXT_PRI))
        row_s  = _work[_work["Center"] == c]
        v1     = int(row_s[WEEKS[0]].values[0])  if not row_s.empty else 0
        v13    = int(row_s[WEEKS[-1]].values[0]) if not row_s.empty else 0
        arrow  = '&#9650; Up'   if v13 > v1 else ('&#9660; Down' if v13 < v1 else '&#8212; Flat')
        ac     = '#34d399' if v13 < v1 else ('#fb923c' if v13 > v1 else '#8b949e')
        snap_rows += (
            f'<tr>'
            f'<td style="padding:6px 10px;background:{bg};color:{fg};font-weight:600;border:1px solid #30363d">{c}</td>'
            f'<td style="padding:6px 10px;text-align:center;background:{bg};color:{fg};border:1px solid #30363d">{v1}</td>'
            f'<td style="padding:6px 10px;text-align:center;background:{bg};color:{fg};border:1px solid #30363d">{v13}</td>'
            f'<td style="padding:6px 10px;text-align:center;background:{bg};border:1px solid #30363d">'
            f'<span style="color:{ac}">{arrow}</span></td>'
            f'</tr>'
        )

    # -- Assemble the full HTML email --------------------------------------
    para_blocks = "".join(
        f'<p style="margin:0 0 12px;line-height:1.75;color:#e6edf3">{p}</p>'
        for p in plain_paras
    )

    th_style = (
        'style="padding:7px 10px;background:#0f2437;color:#38bdf8;'
        'font-size:11px;font-weight:700;text-transform:uppercase;'
        'letter-spacing:0.05em;border:1px solid #30363d;text-align:center"'
    )
    th_left = th_style.replace("text-align:center", "text-align:left")

    # Shared styles injected once into the <head> for collapsible sections
    details_css = (
        'details{border-top:1px solid #30363d;}'
        'details[open] summary{border-bottom:1px solid #21262d;}'
        'summary{'
        '  display:flex;align-items:center;justify-content:space-between;'
        '  padding:13px 28px;cursor:pointer;list-style:none;'
        '  font-size:11px;font-weight:700;text-transform:uppercase;'
        '  letter-spacing:0.08em;color:#38bdf8;'
        '  border-left:3px solid #38bdf8;margin-left:0;'
        '  user-select:none;'
        '}'
        'summary::-webkit-details-marker{display:none}'
        'summary::after{content:"\\203a";font-size:16px;transition:transform .2s;margin-left:auto;padding-left:12px}'
        'details[open] summary::after{transform:rotate(90deg)}'
        '.sec-body{padding:0 28px 16px}'
    )

    def _section(title: str, content: str, open_by_default: bool = True) -> str:
        open_attr = " open" if open_by_default else ""
        return (
            f'<details{open_attr}>'
            f'<summary>{title}</summary>'
            f'<div class="sec-body">{content}</div>'
            f'</details>'
        )

    full_html = (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        f'<style>'
        f'body{{margin:0;padding:0;background:#0d1117;font-family:{FONT}}}'
        'h2,h3,p,td,th,div,span,summary,input,button{{font-family:inherit}}'
        f'{details_css}'
        f'</style></head>'
        f'<body style="background:#0d1117;padding:24px">'
        f'<div style="max-width:640px;margin:0 auto;background:#161b22;border:1px solid #30363d;border-radius:10px;overflow:hidden">'

        # Header - always visible, no collapse
        f'<div style="background:linear-gradient(135deg,#161b22 0%,#0f2437 100%);border-bottom:2px solid #38bdf8;padding:24px 28px 18px">'
        f'<div style="font-size:11px;text-transform:uppercase;letter-spacing:0.1em;color:#8b949e;margin-bottom:6px">{region_label} Workforce Communication</div>'
        f'<div style="font-size:20px;font-weight:700;color:#e6edf3">{quarter_label} Bench Forecast Update &#8212; {report_week}</div>'
        f'<div style="font-size:12px;color:#8b949e;margin-top:4px">From: {sender_name} &nbsp;&middot;&nbsp; Auto-generated from Dashboard</div>'
        f'</div>'

        # KPI pills - always visible
        f'<div style="padding:16px 24px 12px;border-bottom:1px solid #30363d">{pills_html}</div>'

        # Collapsible: Commentary
        + _section("Commentary", para_blocks, open_by_default=True)

        # Collapsible: Q3 Comparison chart
        + _section(
            "Q3 Comparison &#8212; Forecast vs Actuals",
            q3_chart_svg,
            open_by_default=True,
        )

        # Collapsible: Historic Bench trend chart
        + _section(
            "Historic Bench Trend (Avg NA FNC by Quarter)",
            hist_bench_svg,
            open_by_default=True,
        )

        # Collapsible: Center snapshot table
        + _section(
            "Center Snapshot",
            (
                f'<table style="border-collapse:collapse;width:100%;font-size:13px">'
                f'<thead><tr>'
                f'<th {th_left}>Center</th>'
                f'<th {th_style}>Bench Wk01</th>'
                f'<th {th_style}>Bench Wk13</th>'
                f'<th {th_style}>Trend</th>'
                f'</tr></thead><tbody>{snap_rows}</tbody></table>'
            ),
            open_by_default=True,
        )

        # Collapsible: Summary
        + _section(
            "Summary",
            (
                f'<table style="border-collapse:collapse;width:100%;font-size:13px">'
                f'<tr><td style="padding:4px 8px;color:#8b949e;width:160px">Wk01 Bench</td>'
                f'<td style="padding:4px 8px;color:{w1_color};font-weight:700">{_gt_w1} ({_w1_pct:.1f}%) '
                f'{"&#9888; ABOVE TARGET" if _w1_pct > bench_target else "&#10003; ON TARGET"}</td></tr>'
                f'<tr><td style="padding:4px 8px;color:#8b949e">Wk13 Bench</td>'
                f'<td style="padding:4px 8px;color:{w13_color};font-weight:700">{_gt_w13} ({_w13_pct:.1f}%)</td></tr>'
                f'<tr><td style="padding:4px 8px;color:#8b949e">Peak Week</td>'
                f'<td style="padding:4px 8px;color:#38bdf8;font-weight:700">{_peak_val} @ {_peak_wk}</td></tr>'
                f'<tr><td style="padding:4px 8px;color:#8b949e">Avg Bench Rate</td>'
                f'<td style="padding:4px 8px;color:{avg_color};font-weight:700">{_avg_pct:.1f}% '
                f'{"&#9888; ABOVE TARGET" if _avg_pct > bench_target else "&#10003; ON TARGET"}</td></tr>'
                f'<tr><td style="padding:4px 8px;color:#8b949e">YoY W1 Delta</td>'
                f'<td style="padding:4px 8px;color:{"#f87171" if _q3_delta > 0 else "#34d399"};font-weight:700">'
                f'{_q3_delta:+d} vs Q3 2025 ({_q3_ac_w1} actual &#8594; {_q3_fc_w1} forecast)</td></tr>'
                f'</table>'
            ),
            open_by_default=True,
        )

        # Footer - always visible
        + f'<div style="padding:12px 28px;background:#21262d;border-top:1px solid #30363d;'
        f'font-size:11px;color:#8b949e">Auto-generated from the NA Bench Forecast Dashboard &nbsp;&middot;&nbsp; {sender_name}</div>'
        f'</div></body></html>'
    )

    # -- Render with toolbar + copy button ---------------------------------
    escaped_html = full_html.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")

    report_component = f"""
<style>
  body {{ margin:0; background:{BG}; font-family:Inter,system-ui,sans-serif; }}
  #toolbar {{
    background:{SURFACE2}; border:1px solid {BORDER};
    border-bottom:none; border-radius:8px 8px 0 0;
    padding:8px 14px; display:flex; align-items:center; justify-content:space-between;
  }}
  #toolbar span {{ font-size:0.73rem; color:{TEXT_SEC}; text-transform:uppercase; letter-spacing:0.07em; font-weight:600; }}
  .tbtn {{
    background:transparent; border:1px solid {ACCENT}; color:{ACCENT};
    border-radius:5px; padding:4px 14px; font-size:0.78rem; font-weight:600;
    cursor:pointer; font-family:inherit; margin-left:6px;
  }}
  .tbtn:hover {{ background:rgba(56,189,248,0.12); }}
  .tbtn.done  {{ border-color:{GOOD}; color:{GOOD}; }}
  #preview {{
    border:1px solid {BORDER}; border-radius:0 0 8px 8px;
    overflow:auto; background:#0d1117;
  }}
</style>
<div id="toolbar">
  <span>&#9993; HTML Email Report</span>
  <button class="tbtn" id="copy-btn" onclick="copyHTML()">&#128203; Copy HTML</button>
</div>
<div id="preview">
  <iframe id="frame" srcdoc="" style="width:100%;border:none;display:block" scrolling="yes"></iframe>
</div>
<script>
const htmlContent = `{escaped_html}`;
const frame = document.getElementById('frame');
frame.srcdoc = htmlContent;
frame.onload = function() {{
  const h = frame.contentDocument.body.scrollHeight;
  frame.style.height = (h + 32) + 'px';
}};

function copyHTML() {{
  navigator.clipboard.writeText(htmlContent).then(() => {{
    const btn = document.getElementById('copy-btn');
    btn.textContent = '&#10003; Copied!';
    btn.classList.add('done');
    setTimeout(() => {{ btn.textContent = '&#128203; Copy HTML'; btn.classList.remove('done'); }}, 2000);
  }});
}}
</script>
"""
    with st.expander("HTML Email Report - with Charts", expanded=True):
        st.caption("Click Copy HTML - paste into Outlook (Insert HTML) or open in browser.")
        components.html(report_component, height=2200, scrolling=True)


# -----------------------------------------------------------------------------
# Footer
# -----------------------------------------------------------------------------
st.markdown(f"""
<div style="text-align:center;color:{TEXT_SEC};font-size:0.75rem;
            border-top:1px solid {BORDER};margin-top:48px;padding-top:14px;">
    Made with IBM Bob
</div>
""", unsafe_allow_html=True)
