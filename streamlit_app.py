"""
India Bond Maturity Tracker — Streamlit app.

Data is pulled directly from the NSDL public API and cached for 24 hours.
MF yield enrichment (Yield / As of / Holders) comes from data/mf_yields.csv
in this repo, built fortnightly from the debt-scheme portfolio disclosures
of HDFC, SBI, ICICI Prudential, Aditya Birla SL, Axis, Nippon, Kotak, UTI
and Tata mutual funds (see build_mf_yields.py). Holders shows each AMC's
market value in INR crore, e.g. "HDFC (50), SBI (100)". Where a more recent BSE/NSE
secondary-market trade exists (data/secondary_yields.csv, built by the
secondary-trades-refresh workflow), it overrides the MF mark.

Deploy: push this repo to GitHub, then connect it at share.streamlit.io
"""

import io
import json
import re
import time
import urllib.parse
import warnings
from datetime import date, datetime

import openpyxl
import pandas as pd
import requests
import streamlit as st

# ------------------------------------------------------------------ #
# Page config
# ------------------------------------------------------------------ #
st.set_page_config(
    page_title="India Bond Maturity Tracker",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------------------------------------------ #
# Design system (mirrors ratings-tool app.py — Snazzy Indigo)
# ------------------------------------------------------------------ #
_CSS = """
<style>
/* ═══════════════════════════════════════════════
   Snazzy Indigo design system — DM Sans · Indigo · Frosted Glass
   ═══════════════════════════════════════════════ */
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&display=swap');

/* ── Design tokens ── */
:root {
    --bg:           #F8F9FF;
    --surface:      #FFFFFF;
    --surface-glass:rgba(255,255,255,0.72);
    --border:       #E0E2EF;
    --border-soft:  #ECEEFF;
    --text:         #111827;
    --text-muted:   #6B7280;
    --accent:       #6366F1;
    --accent-hov:   #4F46E5;
    --accent-light: #EEF2FF;
    --accent-dim:   rgba(99,102,241,0.12);
    --secondary:    #F3F4F6;
    --secondary-hov:#E5E7EB;
    --shadow-xs:    0 1px 3px rgba(99,102,241,0.08);
    --shadow-sm:    0 4px 16px rgba(99,102,241,0.12);
    --shadow-md:    0 8px 32px rgba(99,102,241,0.16);
    --radius-sm:    8px;
    --radius-md:    12px;
    --radius-lg:    16px;
    --font:         'DM Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

/* ── Hide Streamlit chrome ── */
#MainMenu, footer, .stDeployButton { visibility: hidden; }

/* ── Global background & font ── */
html, body {
    background-color: var(--bg) !important;
    font-family: var(--font) !important;
    color: var(--text);
}
[data-testid="stApp"],
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
[data-testid="stMainBlockContainer"],
[data-testid="stVerticalBlock"],
.main, .main .block-container,
section.main > div,
.stMainBlockContainer {
    background-color: var(--bg) !important;
    font-family: var(--font) !important;
}
/* Markdown text */
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] span {
    color: var(--text) !important;
    font-family: var(--font) !important;
}

/* ── Top header bar ── */
[data-testid="stHeader"] {
    background-color: var(--bg) !important;
    border-bottom: 1px solid var(--border);
}

/* ═══════════════ SIDEBAR — FROSTED GLASS ═══════════════ */
[data-testid="stSidebar"],
[data-testid="stSidebarContent"],
section[data-testid="stSidebar"] > div {
    background: var(--surface-glass) !important;
    backdrop-filter: blur(24px) saturate(180%) !important;
    -webkit-backdrop-filter: blur(24px) saturate(180%) !important;
    border-right: 1px solid rgba(99,102,241,0.15) !important;
}
/* Sidebar section header */
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: var(--accent) !important;
    font-family: var(--font) !important;
    font-size: 0.68rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
}
/* All sidebar text */
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label {
    color: var(--text) !important;
    font-family: var(--font) !important;
}
/* Sidebar widget labels */
[data-testid="stSidebar"] .stTextInput label,
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stSlider label,
[data-testid="stSidebar"] .stNumberInput label,
[data-testid="stSidebar"] .stRadio label,
[data-testid="stSidebar"] .stCheckbox label {
    color: var(--text-muted) !important;
    font-size: 0.73rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.02em;
}
/* Sidebar buttons */
[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"],
[data-testid="stSidebar"] button {
    border-radius: var(--radius-sm) !important;
    font-size: 0.8rem !important;
    background: rgba(255,255,255,0.8) !important;
    border: 1px solid var(--border) !important;
    color: var(--text) !important;
    font-weight: 500 !important;
    font-family: var(--font) !important;
    transition: all 0.15s ease;
}
[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"]:hover,
[data-testid="stSidebar"] button:hover {
    background: white !important;
    border-color: var(--accent) !important;
    color: var(--accent) !important;
    box-shadow: 0 0 0 3px var(--accent-dim) !important;
}
/* Sidebar dividers */
[data-testid="stSidebar"] hr { border-color: var(--border-soft) !important; }

/* ═══════════════ METRIC CARDS ═══════════════ */
[data-testid="stMetric"] {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-lg) !important;
    padding: 1.1rem 1.4rem !important;
    box-shadow: var(--shadow-xs) !important;
    transition: box-shadow 0.2s ease, transform 0.2s ease;
    position: relative;
    overflow: hidden;
}
[data-testid="stMetric"]:hover {
    box-shadow: var(--shadow-sm) !important;
    transform: translateY(-1px);
}
[data-testid="stMetric"]::before {
    content: '';
    position: absolute;
    top: 0; left: 0;
    width: 3px; height: 100%;
    background: linear-gradient(180deg, var(--accent), #818CF8);
    border-radius: 2px 0 0 2px;
}
[data-testid="stMetricLabel"] p {
    font-size: 0.67rem !important;
    color: var(--text-muted) !important;
    font-weight: 600 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
    font-family: var(--font) !important;
}
[data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    font-weight: 700 !important;
    color: var(--text) !important;
    letter-spacing: -0.025em !important;
    font-family: var(--font) !important;
}
[data-testid="stMetricDelta"] { font-size: 0.78rem; }

/* ═══════════════ EXPANDERS ═══════════════ */
[data-testid="stExpander"],
[data-testid="stExpanderDetails"] {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-md) !important;
    background: var(--surface) !important;
    overflow: hidden;
}
[data-testid="stExpander"] details {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-md) !important;
    background: var(--surface) !important;
    overflow: hidden;
}
[data-testid="stExpanderToggle"],
[data-testid="stExpander"] summary {
    color: var(--text) !important;
    font-weight: 500 !important;
    font-size: 0.83rem !important;
    font-family: var(--font) !important;
    padding: 0.65rem 1rem !important;
    background: var(--surface) !important;
}
[data-testid="stExpanderToggle"]:hover,
[data-testid="stExpander"] summary:hover { background: var(--secondary) !important; }
[data-testid="stExpanderToggle"] p,
[data-testid="stExpander"] summary > span,
[data-testid="stExpander"] summary p { color: var(--text) !important; }

/* ═══════════════ DIVIDERS ═══════════════ */
hr { border-color: var(--border) !important; opacity: 1 !important; }
[data-testid="stDivider"] hr { border-color: var(--border) !important; }

/* ═══════════════ BUTTONS ═══════════════ */
/* Primary — indigo fill */
[data-testid="stBaseButton-primary"],
[data-testid="stDownloadButton"] button {
    border-radius: var(--radius-sm) !important;
    background: linear-gradient(135deg, var(--accent) 0%, #818CF8 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    font-weight: 600 !important;
    font-family: var(--font) !important;
    letter-spacing: 0.01em !important;
    box-shadow: 0 2px 8px var(--accent-dim) !important;
    transition: all 0.2s ease !important;
}
[data-testid="stBaseButton-primary"]:hover,
[data-testid="stDownloadButton"] button:hover {
    background: linear-gradient(135deg, var(--accent-hov) 0%, var(--accent) 100%) !important;
    box-shadow: 0 4px 16px rgba(99,102,241,0.35) !important;
    transform: translateY(-1px);
}
/* Secondary */
[data-testid="stBaseButton-secondary"] {
    border-radius: var(--radius-sm) !important;
    background: var(--accent-light) !important;
    color: var(--accent) !important;
    border: 1px solid rgba(99,102,241,0.25) !important;
    font-weight: 600 !important;
    font-family: var(--font) !important;
    transition: all 0.15s ease !important;
}
[data-testid="stBaseButton-secondary"]:hover {
    background: #E0E7FF !important;
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 3px var(--accent-dim) !important;
}
/* Minimal */
[data-testid="stBaseButton-minimal"],
[data-testid="stBaseButton-borderless"] {
    border-radius: var(--radius-sm) !important;
    background: transparent !important;
    color: var(--accent) !important;
    border: none !important;
    font-weight: 500 !important;
    font-family: var(--font) !important;
}
[data-testid="stBaseButton-minimal"]:hover,
[data-testid="stBaseButton-borderless"]:hover {
    background: var(--accent-light) !important;
}

/* ═══════════════ DATA TABLE ═══════════════ */
[data-testid="stDataEditor"],
[data-testid="stDataFrame"] {
    border-radius: var(--radius-lg) !important;
    border: 1px solid rgba(99,102,241,0.2) !important;
    box-shadow: var(--shadow-xs) !important;
    overflow: hidden;
    background: var(--surface) !important;
}

/* ═══════════════ INPUTS ═══════════════ */
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-sm) !important;
    color: var(--text) !important;
    font-family: var(--font) !important;
    font-size: 0.85rem !important;
    transition: border-color 0.15s ease, box-shadow 0.15s ease;
}
[data-testid="stTextInput"] input:focus,
[data-testid="stNumberInput"] input:focus {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 3px var(--accent-dim) !important;
    outline: none;
}
/* Selectbox */
[data-testid="stSelectbox"] [data-baseweb="select"] > div:first-child,
[data-testid="stMultiSelect"] [data-baseweb="select"] > div:first-child {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-sm) !important;
    font-family: var(--font) !important;
    font-size: 0.85rem !important;
}
/* Dropdown menus */
[data-baseweb="popover"] ul,
[data-baseweb="menu"],
[data-baseweb="popover"] [data-baseweb="menu-item"] {
    background: var(--surface) !important;
    font-family: var(--font) !important;
    font-size: 0.83rem !important;
    color: var(--text) !important;
}
[data-baseweb="option"]:hover,
[role="option"]:hover {
    background: var(--secondary) !important;
}
/* Multiselect tags — indigo pill */
[data-baseweb="tag"] {
    background: linear-gradient(135deg, var(--accent), #818CF8) !important;
    border-radius: 20px !important;
    border: none !important;
}
[data-baseweb="tag"] span { color: #FFFFFF !important; font-size: 0.78rem !important; font-weight: 500 !important; }

/* ═══════════════ SLIDERS ═══════════════ */
[data-testid="stSlider"] label,
[data-testid="stSlider"] p {
    color: var(--text-muted) !important;
    font-size: 0.75rem !important;
    font-weight: 500 !important;
}
[data-testid="stSlider"] [data-testid="stTickBarMin"],
[data-testid="stSlider"] [data-testid="stTickBarMax"] {
    color: var(--text-muted) !important;
    font-size: 0.7rem !important;
}

/* ═══════════════ CHECKBOXES & RADIOS ═══════════════ */
[data-testid="stCheckbox"] label p,
[data-testid="stCheckbox"] span {
    color: var(--text) !important;
    font-size: 0.83rem !important;
    font-family: var(--font) !important;
}
[data-testid="stRadio"] label p,
[data-testid="stRadio"] > div > label {
    color: var(--text) !important;
    font-size: 0.83rem !important;
    font-family: var(--font) !important;
}
[data-testid="stRadio"] > div > label > div {
    color: var(--text) !important;
}

/* ═══════════════ TOGGLE ═══════════════ */
[data-testid="stToggle"] label,
[data-testid="stToggle"] p {
    color: var(--text) !important;
    font-size: 0.83rem !important;
    font-family: var(--font) !important;
}

/* ═══════════════ TYPOGRAPHY ═══════════════ */
h1, h2, h3, h4 {
    color: var(--text) !important;
    font-family: var(--font) !important;
    font-weight: 700 !important;
    letter-spacing: -0.025em !important;
}
h1 { font-size: 1.65rem !important; line-height: 1.25; }
h2 { font-size: 1.15rem !important; }
h3 { font-size: 0.97rem !important; }
p, li { font-family: var(--font) !important; }

/* Caption / muted text */
[data-testid="stCaptionContainer"] p,
.stCaption p, small {
    color: var(--text-muted) !important;
    font-size: 0.76rem !important;
    font-family: var(--font) !important;
}

/* ═══════════════ ALERTS / INFO ═══════════════ */
[data-testid="stAlert"],
[data-testid="stNotification"],
[data-testid="stAlertContentInfo"],
[data-testid="stAlertContentWarning"],
[data-testid="stAlertContentError"],
[data-testid="stAlertContentSuccess"] {
    border-radius: var(--radius-md) !important;
    font-family: var(--font) !important;
}
[data-testid="stAlert"] p,
[data-testid="stAlertContentInfo"] p { color: var(--text) !important; }

/* ═══════════════ SUBHEADER ═══════════════ */
[data-testid="stSubheader"] h2,
[data-testid="stSubheader"] p {
    color: var(--text) !important;
    font-family: var(--font) !important;
    font-weight: 600 !important;
    font-size: 1.0rem !important;
    letter-spacing: -0.015em !important;
}

/* ═══════════════ SPINNER ═══════════════ */
.stSpinner > div { border-top-color: var(--primary) !important; }

/* ═══════════════ TOAST ═══════════════ */
[data-testid="stToast"] {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-md) !important;
    color: var(--text) !important;
    box-shadow: var(--shadow-sm) !important;
    font-family: var(--font) !important;
}

/* ═══════════════ MATERIAL ICON GLYPHS ═══════════════
   Streamlit renders expander arrows / widget icons via the Material Symbols
   ligature font. The DM Sans overrides above must NOT apply to them, or the
   ligature text ("keyboard_arrow_right", …) renders as literal characters
   overlapping the labels. */
[data-testid="stIconMaterial"],
span[data-testid="stIconMaterial"],
[data-testid="stExpanderToggleIcon"],
span[class*="material-symbols"],
i[class*="material-symbols"] {
    font-family: "Material Symbols Rounded" !important;
    font-weight: normal !important;
    letter-spacing: normal !important;
    text-transform: none !important;
}
</style>
"""

# ------------------------------------------------------------------ #
# Constants
# ------------------------------------------------------------------ #
BASE_API = "https://www.indiabondinfo.nsdl.com/bds-service/v1/public/bdsinfo"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.indiabondinfo.nsdl.com/CBDServices/",
    "Accept": "*/*",
}

MF_YIELDS_URL = (
    "https://raw.githubusercontent.com/k13-spec/bond-tracker"
    "/main/data/mf_yields.csv"
)
MF_HISTORY_URL = (
    "https://raw.githubusercontent.com/k13-spec/bond-tracker"
    "/main/data/mf_yields_history.csv"
)
SECONDARY_YIELDS_URL = (
    "https://raw.githubusercontent.com/k13-spec/bond-tracker"
    "/main/data/secondary_yields.csv"
)
SECONDARY_HISTORY_URL = (
    "https://raw.githubusercontent.com/k13-spec/bond-tracker"
    "/main/data/secondary_yields_history.csv"
)
SECONDARY_META_URL = (
    "https://raw.githubusercontent.com/k13-spec/bond-tracker"
    "/main/data/secondary_meta.json"
)
# GitHub Actions workflow dispatched by the sidebar "Refresh Secondary Trades" button
GH_DISPATCH_URL = (
    "https://api.github.com/repos/k13-spec/bond-tracker"
    "/actions/workflows/secondary-trades-refresh.yml/dispatches"
)
# Self-URL used to make ISINs clickable (deep link to the yield-history view)
APP_BASE_URL = "https://creditnexus-bonds.streamlit.app"

RATING_GRADES = [
    "AAA", "AA+", "AA", "AA-",
    "A+", "A", "A-",
    "BBB+", "BBB", "BBB-",
    "BB+", "BB", "BB-",
    "B+", "B", "B-",
    "CCC+", "CCC", "CCC-", "CC", "C", "D",
    "A1+", "A1", "A2+", "A2", "A3", "A4",
]
# Longest-first alternation + custom boundaries so "AA-" / "AA+" / "BBB-"
# match as themselves instead of collapsing to "AA" / "BBB" (\b treats
# +/- as a boundary, so the old pattern could never return modifier grades).
GRADE_PATTERN = re.compile(
    r"(?<![A-Za-z0-9])("
    + "|".join(re.escape(g)
               for g in sorted(RATING_GRADES, key=len, reverse=True))
    + r")(?![A-Za-z0-9])"
)

# grade integer (ratings DB) -> long-term symbol, most reliable for cross-ref
GRADE_INT2SYM = {
    1: "AAA", 2: "AA+", 3: "AA", 4: "AA-", 5: "A+", 6: "A", 7: "A-",
    8: "BBB+", 9: "BBB", 10: "BBB-", 11: "BB+", 12: "BB", 13: "BB-",
    14: "B+", 15: "B", 16: "B-", 17: "CCC+", 18: "CCC", 19: "CCC-", 20: "D",
}


# ------------------------------------------------------------------ #
# Data loading (cached 24 hours)
# ------------------------------------------------------------------ #

def _parse_date(s) -> str | None:
    if not s:
        return None
    s = str(s).strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def _parse_size_cr(raw) -> float | None:
    """NSDL 'Issue Size(in Rs.)' -> INR crore.

    The field is rupees, but a small tail of filings uses other units
    (units/lakhs/crores) with no way to tell reliably. Values below
    Rs. 1 lakh are treated as unreliable and dropped rather than shown
    as absurd crore figures (the old heuristic let e.g. a raw 500000
    through as "500,000 Cr", wrecking any size-based sort).
    """
    if raw is None or raw == "":
        return None
    try:
        val = float(str(raw).replace(",", "").strip())
        if val < 1e5:          # < Rs. 1 lakh: zero, garbage, or unknown unit
            return None
        return round(val / 1e7, 2)
    except (ValueError, TypeError):
        return None


_COUPON_ZERO_RE = re.compile(r"^(zero\s*coupon|zero|nil)$", re.I)


def _parse_coupon(raw) -> tuple[float | None, str]:
    """NSDL 'Coupon Rate (%)' -> (numeric rate or None, detail text).

    NSDL mixes plain numbers with text: zero-coupon markers, floating-rate
    formulas ("SBI BASE RATE+300 BASIS POINT", "RESET RATE"), and
    market-linked debenture underlyings ("NIFTY 50 INDEX LINKED", "SENSEX").
    The numeric column stays sortable; the text moves to Coupon Detail.
    """
    if not raw:
        return None, ""
    s = str(raw).strip()
    num = s.rstrip("%").strip().replace(",", "")
    try:
        val = float(num)
        if val == 0:
            return 0.0, "Zero Coupon"
        return round(val, 4), ""
    except ValueError:
        pass
    if _COUPON_ZERO_RE.match(s):
        return 0.0, "Zero Coupon"
    if s.upper() in ("N.A", "N.A.", "NA", "-"):
        return None, ""
    return None, re.sub(r"\s+", " ", s)


def _primary_rating(raw: str | None) -> str | None:
    if not raw:
        return None
    grades = GRADE_PATTERN.findall(str(raw))
    return grades[0] if grades else None


def _listing(mode: str | None) -> str:
    if not mode:
        return "Unknown"
    m = str(mode).lower()
    if "public" in m:
        return "Listed"
    if "private" in m:
        return "Unlisted"
    return "Unknown"


@st.cache_data(ttl=86400, show_spinner=False)   # cache 24 hours
def load_bonds() -> pd.DataFrame:
    """Download and parse the active bond list from NSDL. Cached 24 h."""
    session = requests.Session()
    session.headers.update(HEADERS)
    session.get("https://www.indiabondinfo.nsdl.com/CBDServices/", timeout=15)
    time.sleep(0.3)

    resp = session.get(
        f"{BASE_API}/listofsecurities?type=Active",
        timeout=120,
        stream=True,
    )
    resp.raise_for_status()
    data = resp.content

    warnings.filterwarnings("ignore", category=UserWarning)
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]

    col = {h: i for i, h in enumerate(headers) if h}

    def g(row, *names):
        for n in names:
            idx = col.get(n)
            if idx is not None and row[idx] is not None:
                v = str(row[idx]).strip()
                if v and v not in ("-", "N.A.", "NA"):
                    return v
        return None

    today = date.today().isoformat()
    records = []
    for row in rows[1:]:
        if not row or not row[1]:
            continue
        isin = str(row[col.get("ISIN", 1)] or "").strip()
        if not isin or not isin.startswith("IN"):
            continue

        maturity_date = _parse_date(g(row, "Date of Redemption/Conversion"))
        if not maturity_date or maturity_date < today:
            continue

        issue_date    = _parse_date(g(row, "Date of Allotment"))
        raw_rating    = g(row, "Credit Rating")
        mode_of_issue = g(row, "Mode of Issue") or ""
        coupon_val, coupon_detail = _parse_coupon(g(row, "Coupon Rate (%)"))

        records.append({
            "ISIN":            isin,
            "Issuer":          g(row, "Name of Issuer"),
            "Type":            g(row, "Type of Instrument"),
            "Series":          g(row, "Series"),
            "Description":     g(row, "Security Description"),
            "Mode of Issue":   mode_of_issue,
            "Coupon Rate (%)": coupon_val,
            "Coupon Detail":   coupon_detail,
            "Coupon Type":     g(row, "Coupon Type"),
            "Coupon Freq":     g(row, "Frequency of Interest Payment"),
            "Issue Date":      issue_date,
            "Maturity Date":   maturity_date,
            "Issue Size (Cr)": _parse_size_cr(g(row, "Issue Size(in Rs.)")),
            "Rating":          _primary_rating(raw_rating),
            "Rating (Full)":   raw_rating,
            "Listing Status":  _listing(mode_of_issue),
            "Sector":          (g(row, "Business Sector") or "").split("(")[0].strip(),
            "Issuer Type":     g(row, "Type of Issuer-Ownership"),
        })

    df = pd.DataFrame(records)
    if not df.empty:
        df["Issue Date"]    = pd.to_datetime(df["Issue Date"], errors="coerce")
        df["Maturity Date"] = pd.to_datetime(df["Maturity Date"], errors="coerce")
        df["Coupon Rate (%)"] = pd.to_numeric(df["Coupon Rate (%)"], errors="coerce")
        df["Issue Size (Cr)"] = pd.to_numeric(df["Issue Size (Cr)"], errors="coerce")
        today_dt = pd.Timestamp(date.today())
        df["Days to Maturity"] = (df["Maturity Date"] - today_dt).dt.days.astype("Int64")
        # Renewable-energy issuers (power producers, not their lenders) get
        # their own sub-sector under Infrastructure — NSDL lumps them into
        # Electric Utilities / Energy / Power etc.
        _renew = (df["Issuer"].str.contains(_RENEW_RE, na=False)
                  & (df["Sector"].map(_sector_group) != "Financial Institutions"))
        df.loc[_renew, "Sector"] = "Renewable Energy"
        # Effective PSU flag: NSDL's Type of Issuer-Ownership field where
        # filled, else name-based detection (the field is blank for ~10%
        # of rows and NSDL sometimes misfiles well-known PSUs).
        _own_psu = df["Issuer Type"].fillna("").str.startswith("Public Sector")
        df["PSU"] = (_own_psu | df["Issuer"].apply(_is_psu)).map(
            {True: "PSU", False: "Non-PSU"})
    return df


# ------------------------------------------------------------------ #
# MF fortnightly yield enrichment (cached 1 hour)
# ------------------------------------------------------------------ #

@st.cache_data(ttl=3600, show_spinner=False)
def _load_mf_latest_raw() -> pd.DataFrame:
    """Download mf_yields.csv (latest fortnightly snapshot), as-is. Cached 1 h."""
    try:
        mf = pd.read_csv(MF_YIELDS_URL)
        mf["isin"] = mf["isin"].astype(str).str.strip()
        return mf.drop_duplicates(subset="isin", keep="first")
    except Exception:
        return pd.DataFrame(
            columns=["isin", "yield", "as_of", "holders", "source",
                     "mf_rating", "instrument_name"])


def _load_mf_yields() -> pd.DataFrame:
    """Latest MF yields shaped for the table merge (ISIN / Yield / As of / Src / Holders)."""
    mf = _load_mf_latest_raw().copy()
    if mf.empty:
        return pd.DataFrame(columns=["ISIN", "Yield (%)", "As of", "Src", "Holders"])
    mf["as_of"] = pd.to_datetime(mf["as_of"], errors="coerce").dt.strftime("%d/%m/%Y")
    return mf.rename(columns={
        "isin":    "ISIN",
        "yield":   "Yield (%)",
        "as_of":   "As of",
        "source":  "Src",
        "holders": "Holders",
    })[["ISIN", "Yield (%)", "As of", "Src", "Holders"]]


@st.cache_data(ttl=3600, show_spinner=False)
def _load_mf_history() -> pd.DataFrame:
    """Download mf_yields_history.csv — every (isin, as_of) yield ever collected.

    One row per ISIN per fortnightly disclosure; grows by one date each refresh.
    Cached 1 h. Columns: isin, as_of (datetime), yield, source, holders.
    """
    try:
        h = pd.read_csv(MF_HISTORY_URL)
        h["isin"] = h["isin"].astype(str).str.strip()
        h["as_of"] = pd.to_datetime(h["as_of"], errors="coerce")
        h = h.dropna(subset=["as_of", "yield"])
        return h.sort_values(["isin", "as_of"])
    except Exception:
        return pd.DataFrame(columns=["isin", "as_of", "yield", "source", "holders"])


# ------------------------------------------------------------------ #
# Secondary-market trade enrichment (BSE/NSE, cached 10 minutes)
# ------------------------------------------------------------------ #

@st.cache_data(ttl=600, show_spinner=False)
def _load_secondary_yields() -> pd.DataFrame:
    """Download secondary_yields.csv — one chosen best BSE/NSE trade per ISIN.

    Cached 10 min (shorter than MF, so a pushed refresh shows up quickly).
    Columns: isin, yield, as_of (datetime), source ("BSE"/"NSE"), trade_value_cr.
    """
    try:
        sec = pd.read_csv(SECONDARY_YIELDS_URL)
        sec["isin"] = sec["isin"].astype(str).str.strip()
        sec["as_of"] = pd.to_datetime(sec["as_of"], errors="coerce")
        sec = sec.dropna(subset=["as_of", "yield"])
        return sec.drop_duplicates(subset="isin", keep="first")
    except Exception:
        return pd.DataFrame(
            columns=["isin", "yield", "as_of", "source", "trade_value_cr"])


@st.cache_data(ttl=600, show_spinner=False)
def _load_secondary_history() -> pd.DataFrame:
    """Download secondary_yields_history.csv — every captured trade, append-only.

    Cached 10 min. Columns: isin, yield, as_of (datetime), source, trade_value_cr.
    """
    try:
        h = pd.read_csv(SECONDARY_HISTORY_URL)
        h["isin"] = h["isin"].astype(str).str.strip()
        h["as_of"] = pd.to_datetime(h["as_of"], errors="coerce")
        h = h.dropna(subset=["as_of", "yield"])
        return h.sort_values(["isin", "as_of"])
    except Exception:
        return pd.DataFrame(
            columns=["isin", "yield", "as_of", "source", "trade_value_cr"])


@st.cache_data(ttl=600, show_spinner=False)
def _load_secondary_meta() -> dict:
    """Download secondary_meta.json ({"last_refresh": ISO datetime}). Cached 10 min.

    Also accepts a local file path in SECONDARY_META_URL (used by tests).
    """
    try:
        if SECONDARY_META_URL.startswith("http"):
            r = requests.get(SECONDARY_META_URL, timeout=15)
            r.raise_for_status()
            return json.loads(r.text)
        with open(SECONDARY_META_URL, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


# ------------------------------------------------------------------ #
# Sector grouping (mirrors ratings tool)
# ------------------------------------------------------------------ #
_SECTOR_GROUPS = {
    "Corporate": [
        "Auto", "Auto Components", "Automobiles", "2/3 Wheelers",
        "Passenger Cars & Utility Vehicles", "Commercial Vehicles", "Tractors",
        "Tyres & Rubber Products", "Batteries - Automobile",
        "Aerospace & Defense", "Bearings", "Castings & Forgings",
        "Cement & Cement Products", "Chemicals", "Chemicals & Petrochemicals",
        "Specialty Chemicals", "Commodity Chemicals", "Carbon Black",
        "Dyes And Pigments", "Explosives", "Fertilizers  ",
        "Petrochemicals", "Petroleum Products", "Consumable Fuels",
        "Consumer Durables", "Consumer Electronics", "Household Appliances",
        "Household Products", "Personal Care", "Personal Products",
        "Gems, Jewellery And Watches", "Diversified Consumer Products",
        "Diversified FMCG", "FMCG", "Cigarettes & Tobacco Products",
        "Food Products", "Dairy Products", "Packaged Foods",
        "Other Food Products", "Edible Oil", "Sugar", "Tea & Coffee",
        "Other Beverages", "Breweries & Distilleries", "Agriculture",
        "Animal Feed", "Other Agricultural Products",
        "Healthcare", "Pharmaceuticals", "Pharmaceuticals ",
        "Biotechnology", "Hospital", "Healthcare Service Provider",
        "Healthcare Research, Analytics & Technology",
        "Medical Equipment & Supplies",
        "Technology", "Information Technology",
        "Computers - Software & Consulting", "Computers Hardware & Equipments",
        "IT Enabled Services", "Software Products", "Data Processing Services",
        "E-Learning", "Business Process Outsourcing (BPO) / Knowledge Process Outsourcing (KPO)",
        "Digital Entertainment", "Media & Entertainment",
        "Film Production, Distribution & Exhibition",
        "TV Broadcasting & Software Production", "Electronic Media",
        "Advertising & Media Agencies",
        "Retail", "Diversified Retail", "Speciality Retail",
        "E-Retail/ E- Commerce", "Distributors", "Trading & Distributors",
        "Education", "Hotels & Resorts", "Restaurants",
        "Amusement Parks/ Other Recreation", "Tour, Travel Related Services",
        "Construction", "Civil Construction", "Real Estate",
        "Residential, Commercial Projects", "Real Estate Investment Trusts (REITs)",
        "Real Estate related services",
        "Textiles & Apparels", "Garments & Apparels",
        "Cotton Textiles - Composite", "Other Textile Products",
        "Iron & Steel", "Iron & Steel Products", "Ferrous Metals",
        "Diversified Metals", "Aluminium, Copper & Zinc Products",
        "Minerals & Mining", "Coal",
        "Cement & Cement Products", "Other Construction Materials",
        "Paper & Paper Products", "Forest Products", "Printing & Publication",
        "Packaging", "Plastic Products - Consumer",
        "Industrial Manufacturing", "Industrial Products",
        "Heavy Electrical Equipment", "Electrical Equipment",
        "Industrial Electronics", "Industrial Machinery",
        "Engineering, Designing & Construction",
        "Diversified", "Multi-Product Companies", "Holding Company",
        "Trading - Chemicals", "Trading - Metals", "Trading - Minerals",
        "Consumer Services", "Commercial Services & Supplies",
        "Diversified Commercial Services", "Consulting Services",
        "Wellness", "Other Consumer Services",
    ],
    "Infrastructure": [
        "Infrastructure", "Civil Construction",
        "Renewable Energy",   # derived sub-sector (issuer-name based)
        "Energy", "Power", "Other Utilities",
        "Electric Utilities", "Electricity Generation", "Power Trading",
        "Power - Transmission", "Power - Distribution", "Multi Utilities", "Utilities",
        "Oil Exploration & Production", "Refineries & Marketing",
        "Oil Storage & Transportation", "Gas Transmission/ Marketing",
        "LPG/CNG/PN G/LNG Supplier", "Industrial Gas",
        "Airport & Airport services", "Airline",
        "Port & Port services", "Dredging", "Shipping",
        "Railways", "Road Transport", "Toll bridge operator",
        "Logistics Solution Provider", "Transport Related Services",
        "Waste Management", "Water Supply & Management",
        "Telecom - Cellular & Fixed line services",
        "Telecom - Equipment & Accessories", "Telecom - Infrastructure",
        "Other Telecom Services",
    ],
    "Financial Institutions": [
        "Finance", "Financial Institution",
        # NSDL sector strings arrive with a "(...)" suffix that the app
        # strips for display — both variants must be listed, otherwise
        # e.g. NBFCs fall through to the default (Corporate) bucket.
        "Non-Banking Financial Company (NBFC)", "Non-Banking Financial Company",
        "NBFC", "Housing Finance Company",
        "Private Sector Bank", "Public Sector Bank", "Other Bank",
        "Asset Management Company", "Investment Company",
        "Life Insurance", "General Insurance", "Other Insurance Companies",
        "Insurance Distributors",
        "Financial Technology (Fintech)", "Financial Technology",
        "Stockbroking & Allied",
        "Depositories, Clearing Houses and Other Intermediaries",
        "Other Capital Market related Services", "Other Financial Services",
    ],
}

# Renewable-energy issuer detection (power producers / IPPs). Lenders to the
# sector (IREDA, cleantech NBFCs) stay under Financial Institutions — the
# override in load_bonds() skips issuers whose NSDL sector is financial.
_RENEW_RE = re.compile(
    r"renewab|solar|\bwind\b|windpower|wind power|windfarm|green energy|"
    r"green power|greenko|clean ?energy|cleantech|clean ?max|hydro ?power|"
    r"hydroelectric|photovolta|\brenew\b|suzlon|inox wind|avaada|"
    r"azure power|juniper green|serentica|ayana renewable|amp energy|"
    r"fourth partner|o2 power|radiance renew|virescent|adani green|"
    r"tata power renewable|continuum green|vena energy|sael\b",
    re.I,
)

_PSU_FRAGMENTS = [
    "ntpc ", "bhel ", " sail ", "ongc", "iocl", "gail ",
    "nalco", "nmdc", "nhpc", "npcil", "powergrid", "power grid",
    "irfc", "nhai ", "hudco", "sidbi", "nabard",
    "coal india", "indian oil", "bharat petroleum",
    "hindustan petroleum", "oil and natural gas",
    "gas authority", "steel authority",
    "national aluminium", "national mineral development",
    "national thermal power", "national highways authority",
    "bharat heavy electricals", "bharat electronics",
    "bharat dynamics", "hindustan aeronautics",
    "state bank of india", "punjab national bank",
    "bank of baroda", "bank of india", "bank of maharashtra",
    "canara bank", "union bank of india", "central bank of india",
    "indian bank ", "uco bank", "jammu and kashmir bank",
    "life insurance corporation",
    "power finance corp", "rural electrification corp",
    "housing and urban development", "national bank for agriculture",
    "export import bank", "exim bank",
    "rec limited", "pfc limited",
    "food corporation of india", "oil india", "mrpl", "bpcl", "hpcl",
    # verified against NSDL's Type of Issuer-Ownership field, 2026-08-05
    "indian railway finance", "india infrastructure finance",
    "indian renewable energy development", "nuclear power corporation",
    "national housing bank", "financing infrastructure and development",
    "thdc ", "mtnl", "bsnl", "bharat sanchar", "mahanagar telephone",
    "pnb housing", "ircon", "rites ", "nbcc ", "moil ",
    "mazagon dock", "garden reach ship", "goa shipyard",
    "engineers india", "rashtriya chemicals", "national fertilizers",
    "seci ", "solar energy corporation",
    # state-government entities: discoms/transcos, state FIs, civic bodies
    "power corporation", "energy corporation", "electricity board",
    "state electricity", "rajya vidyut", "prasaran nigam", "vidyut nigam",
    "power distribution company", "power generation co",
    "municipal corporation", "nagar nigam",
    "metropolitan development authority", "capital region development",
    "infrastructure development board", "state beverages",
    "kerala financial corporation", "kerala infrastructure investment",
    "mineral development corporation", "industrial infrastructure corporation",
]


def _is_psu(name) -> bool:
    # NaN / None / non-string issuer names must not crash the filter
    if not isinstance(name, str):
        return False
    n = (" " + name.lower() + " ")
    return any(f in n for f in _PSU_FRAGMENTS)


def _sector_group(sector: str) -> str:
    for grp, members in _SECTOR_GROUPS.items():
        if sector in members:
            return grp
    return "Corporate"


# ------------------------------------------------------------------
# Ratings-DB enrichment (optional, cached)
# ------------------------------------------------------------------
_CO_STOP_RT = {"limited","ltd","private","pvt","llp","corporation",
              "corp","inc","co","bank","finance","financial"}


def _normalize_co_rt(name: str) -> str:
    words = [w.strip('.,&') for w in name.lower().split()]
    return " ".join(w for w in words if w not in _CO_STOP_RT).strip()


@st.cache_data(ttl=3600, show_spinner=False)
def _load_ratings_lookup() -> dict:
    """Download ratings_current.csv from ratings-tool repo. Cached 1h."""
    url = ("https://raw.githubusercontent.com/k13-spec/ratings-tool"
           "/master/data/ratings_current.csv")
    try:
        df = pd.read_csv(url)
        lookup = {}
        for _, row in df.iterrows():
            key = _normalize_co_rt(str(row.get("company_name", "")))
            if key:
                lookup[key] = {
                    "rating":  str(row.get("rating", "") or ""),
                    "grade":   int(row["grade"]) if pd.notna(row.get("grade")) else None,
                    "agency":  str(row.get("agency", "") or ""),
                    "outlook": str(row.get("outlook", "") or ""),
                    "date":    str(row.get("rating_date", "") or ""),
                }
        return lookup
    except Exception:
        return {}


@st.cache_data(ttl=3600, show_spinner=False)
def _load_prefix_map() -> dict:
    """ISIN issuer-prefix (first 7 chars) -> ratings-DB company name.

    Built by ratings-tool/build_isin_map.py from NSDL issuer codes — turns the
    fuzzy name join into an exact ISIN lookup. Cached 1 h; empty dict if the
    map isn't published yet.
    """
    url = ("https://raw.githubusercontent.com/k13-spec/ratings-tool"
           "/master/data/issuer_prefix_map.csv")
    try:
        df = pd.read_csv(url)
        return {str(p): str(c) for p, c in
                zip(df["isin_prefix"], df["company_name"])}
    except Exception:
        return {}



def _sector_checkbox_panel(available_sectors: list) -> list:
    grouped: dict[str, list] = {"Corporate": [], "Infrastructure": [],
                                "Financial Institutions": []}
    for s in available_sectors:
        grouped[_sector_group(s)].append(s)

    for s in available_sectors:
        wkey = f"bchk_{s}"
        if wkey not in st.session_state:
            st.session_state[wkey] = True

    st.markdown("**Sectors**")
    qc1, qc2, qc3 = st.columns(3)
    if qc1.button("All",  key="bsec_all",  use_container_width=True):
        for s in available_sectors:
            st.session_state[f"bchk_{s}"] = True
        st.rerun()
    if qc2.button("None", key="bsec_none", use_container_width=True):
        for s in available_sectors:
            st.session_state[f"bchk_{s}"] = False
        st.rerun()
    if qc3.button("Corp", key="bsec_corp", use_container_width=True,
                  help="Corporate sectors only"):
        for s in available_sectors:
            st.session_state[f"bchk_{s}"] = (_sector_group(s) == "Corporate")
        st.rerun()

    selected = []
    for grp in ["Corporate", "Infrastructure", "Financial Institutions"]:
        members = grouped.get(grp, [])
        if not members:
            continue
        with st.expander(grp, expanded=(grp == "Corporate")):
            ga1, ga2 = st.columns(2)
            if ga1.button("All",  key=f"bgrp_all_{grp}",  use_container_width=True):
                for s in members:
                    st.session_state[f"bchk_{s}"] = True
                st.rerun()
            if ga2.button("None", key=f"bgrp_none_{grp}", use_container_width=True):
                for s in members:
                    st.session_state[f"bchk_{s}"] = False
                st.rerun()
            for sector in sorted(members):
                if st.checkbox(sector or "(unclassified)", key=f"bchk_{sector}"):
                    selected.append(sector)
    return selected


# ------------------------------------------------------------------ #
# App UI
# ------------------------------------------------------------------ #

st.markdown(_CSS, unsafe_allow_html=True)


# ------------------------------------------------------------------ #
# Yield-history view (deep link: ?isin=INE...)
# Clicking an ISIN in the main table lands here. Rendered INSTEAD of the
# main table, and before the NSDL load so it opens instantly.
# ------------------------------------------------------------------ #

def _render_yield_history(isin: str) -> None:
    import altair as alt

    meta = _load_mf_latest_raw()
    row = meta[meta["isin"] == isin]
    name = str(row["instrument_name"].iloc[0]) if not row.empty and pd.notna(row["instrument_name"].iloc[0]) else ""
    rating = str(row["mf_rating"].iloc[0]) if not row.empty and pd.notna(row["mf_rating"].iloc[0]) else ""

    st.markdown(
        f"""
        <div style="background:linear-gradient(135deg,#6366F1 0%,#818CF8 50%,#A5B4FC 100%);
                    border-radius:16px;padding:24px 32px;margin-bottom:16px;
                    box-shadow:0 8px 32px rgba(99,102,241,0.22);">
            <div style="font-family:'DM Sans',sans-serif;font-size:0.78rem;font-weight:600;
                        color:rgba(255,255,255,0.75);letter-spacing:0.08em;margin-bottom:4px;">
                YIELD HISTORY</div>
            <div style="font-family:'DM Sans',sans-serif;font-size:1.45rem;font-weight:700;
                        color:#FFFFFF;letter-spacing:-0.02em;line-height:1.25;">{isin}</div>
            <div style="font-family:'DM Sans',sans-serif;font-size:0.9rem;
                        color:rgba(255,255,255,0.85);margin-top:2px;">
                {name}{(" &middot; " + rating) if rating else ""}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(f'<a href="{APP_BASE_URL}" target="_self" style="font-family:\'DM Sans\','
                'sans-serif;font-size:13px;font-weight:600;color:#6366F1;'
                'text-decoration:none;">← Back to all bonds</a>',
                unsafe_allow_html=True)

    mf_hist = _load_mf_history()
    sec_hist = _load_secondary_history()
    hm = mf_hist[mf_hist["isin"] == isin].copy()
    hs = sec_hist[sec_hist["isin"] == isin].copy()
    if not hm.empty:
        hm["kind"] = hm["source"].astype(str)                # AMC name
    if not hs.empty:
        hs["kind"] = hs["source"].astype(str) + " trade"     # "BSE trade" / "NSE trade"
    h = pd.concat([hm, hs], ignore_index=True).sort_values("as_of")
    for _c in ("holders", "trade_value_cr", "kind"):
        if _c not in h.columns:
            h[_c] = float("nan")
    if h.empty:
        st.info("No yield history for this ISIN yet. History builds up one "
                "point per fortnightly MF disclosure (refreshed on the 5th and "
                "20th of each month) plus any captured BSE/NSE trades.")
        return

    latest, first = h.iloc[-1], h.iloc[0]
    prev = h.iloc[-2] if len(h) > 1 else None
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Latest yield", f"{latest['yield']:.2f}%",
              help=f"As of {latest['as_of']:%d %b %Y} ({latest['source']})")
    c2.metric("vs previous fortnight",
              f"{(latest['yield'] - prev['yield']) * 100:+.0f} bps"
              if prev is not None else "—",
              help=f"Previous: {prev['yield']:.2f}% as of {prev['as_of']:%d %b %Y}"
              if prev is not None else "Only one data point so far")
    c3.metric("Data points", f"{len(h)}",
              help=f"{first['as_of']:%d %b %Y} → {latest['as_of']:%d %b %Y}")
    c4.metric("Holders (latest)",
              f"{len(str(latest['holders']).split(';'))}" if pd.notna(latest["holders"]) else "0",
              help=str(latest["holders"]))

    chart = (
        alt.Chart(h.assign(**{"Yield (%)": h["yield"]}))
        .mark_line(color="#6366F1", strokeWidth=2.5,
                   point=alt.OverlayMarkDef(size=90, filled=True, color="#4F46E5"))
        .encode(
            x=alt.X("as_of:T", title="Date",
                    axis=alt.Axis(format="%d %b %y", labelAngle=0)),
            y=alt.Y("Yield (%):Q", scale=alt.Scale(zero=False)),
            tooltip=[
                alt.Tooltip("as_of:T", title="As of", format="%d %b %Y"),
                alt.Tooltip("Yield (%):Q", format=".2f"),
                alt.Tooltip("kind:N", title="Source"),
            ],
        )
        .properties(height=380)
    )
    st.altair_chart(chart, use_container_width=True)
    if len(h) == 1:
        st.caption("One data point so far — the trendline builds up with every "
                   "fortnightly disclosure (5th and 20th of each month).")

    shown = h.rename(columns={"as_of": "As of", "yield": "Yield (%)",
                              "kind": "Source", "holders": "Holders"})
    shown["As of"] = shown["As of"].dt.strftime("%d/%m/%Y")
    shown["Trade Value (Cr)"] = shown["trade_value_cr"].apply(
        lambda x: f"{x:,.2f}" if pd.notna(x) else "—")
    shown["Holders"] = shown["Holders"].fillna("—")
    st.dataframe(shown[["As of", "Yield (%)", "Source", "Trade Value (Cr)", "Holders"]]
                 .iloc[::-1], hide_index=True, use_container_width=True)


_qp_isin = st.query_params.get("isin", "").strip().upper()
if _qp_isin:
    _render_yield_history(_qp_isin)
    st.stop()
st.markdown(
    """
    <div style="
        background: linear-gradient(135deg, #6366F1 0%, #818CF8 50%, #A5B4FC 100%);
        border-radius: 16px;
        padding: 28px 32px 22px 32px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 8px 32px rgba(99,102,241,0.22);
    ">
        <div>
            <div style="font-family:'DM Sans',sans-serif;font-size:1.65rem;font-weight:700;
                        color:#FFFFFF;letter-spacing:-0.025em;line-height:1.2;margin-bottom:6px;">
                India Bond Maturity Tracker
            </div>
            <div style="font-family:'DM Sans',sans-serif;font-size:0.83rem;font-weight:400;
                        color:rgba(255,255,255,0.8);letter-spacing:0.01em;">
                NSDL India Bond Info &middot; Refreshed every 24 hours &nbsp;&nbsp;|&nbsp;&nbsp;
                Yields from MF fortnightly disclosures and BSE/NSE trades
            </div>
        </div>
        <div>
            <a href="https://creditnexus.streamlit.app/" target="_blank"
               rel="noopener noreferrer"
               style="font-family:'DM Sans',sans-serif;font-size:13px;font-weight:600;
                      color:#6366F1;background:#FFFFFF;border-radius:8px;
                      padding:7px 16px;text-decoration:none;white-space:nowrap;
                      box-shadow:0 2px 8px rgba(0,0,0,0.12);display:inline-block;">
                ↗ Credit Ratings
            </a>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Load data
with st.spinner("Loading bond data from NSDL… (first load takes ~30 seconds)"):
    try:
        df = load_bonds()
    except Exception as e:
        st.error(f"Failed to load data from NSDL: {e}")
        st.stop()

if df.empty:
    st.warning("No upcoming bond maturities found.")
    st.stop()

# MF yield enrichment (Yield / As of / Src / Holders), merged by ISIN
_mf_yields = _load_mf_yields()
if not _mf_yields.empty:
    df = df.merge(_mf_yields, on="ISIN", how="left")
else:
    df["Yield (%)"] = None
    df["As of"] = None
    df["Src"] = None
    df["Holders"] = None

# Secondary-trade override (BSE/NSE): where an exchange trade exists and is
# strictly more recent than the MF mark (or the bond has no MF mark at all),
# show the traded yield / trade date instead, with Src = exchange name.
_secondary = _load_secondary_yields()
if not _secondary.empty:
    _sec = _secondary.rename(columns={
        "isin":   "ISIN",
        "yield":  "_sec_yield",
        "as_of":  "_sec_as_of",
        "source": "_sec_src",
    })[["ISIN", "_sec_yield", "_sec_as_of", "_sec_src"]]
    df = df.merge(_sec, on="ISIN", how="left")
    _mf_as_of = pd.to_datetime(df["As of"], format="%d/%m/%Y", errors="coerce")
    _override = df["_sec_yield"].notna() & (
        df["Yield (%)"].isna() | (df["_sec_as_of"] > _mf_as_of)
    )
    df.loc[_override, "Yield (%)"] = df.loc[_override, "_sec_yield"]
    df.loc[_override, "As of"] = df.loc[_override, "_sec_as_of"].dt.strftime("%d/%m/%Y")
    df.loc[_override, "Src"] = df.loc[_override, "_sec_src"]
    df = df.drop(columns=["_sec_yield", "_sec_as_of", "_sec_src"])

# ------------------------------------------------------------------ #
# Latest-rating cross-reference (ratings tool DB)
# NSDL's Credit Rating field is as filed at issuance and often stale.
# Where the issuer is found in the ratings DB (creditnexus ratings
# tracker), Rating shows the current rating and Rating Src the agency;
# otherwise it falls back to the NSDL at-issuance rating.
# ------------------------------------------------------------------ #
df["Rating (NSDL)"] = df["Rating"]
df["Rating Src"] = "NSDL"
df["Rated On"] = pd.NaT
_rt_lookup = _load_ratings_lookup()
if _rt_lookup:
    _pfx_map = _load_prefix_map()
    _db_syms, _db_agencies, _db_dates = [], [], []
    for _isin, _issuer in zip(df["ISIN"], df["Issuer"]):
        # 1) exact join via the ISIN issuer-prefix map (immune to name quirks)
        _rec = None
        _cname = _pfx_map.get(str(_isin)[:7])
        if _cname:
            _rec = _rt_lookup.get(_normalize_co_rt(_cname))
        # 2) fall back to the normalized-name join
        if _rec is None:
            _rec = _rt_lookup.get(
                _normalize_co_rt(_issuer) if isinstance(_issuer, str) else "")
        # Prefer the DB's parsed grade integer (exact, incl. +/- modifiers);
        # fall back to parsing the rating string.
        _sym = (GRADE_INT2SYM.get(_rec.get("grade"))
                or _primary_rating(_rec.get("rating"))) if _rec else None
        _db_syms.append(_sym)
        _db_agencies.append((_rec.get("agency") or "").strip() if _rec else "")
        _db_dates.append((_rec.get("date") or "") if _rec else "")
    df["_db_rating"] = _db_syms
    _has_db = df["_db_rating"].notna()
    df.loc[_has_db, "Rating"] = df.loc[_has_db, "_db_rating"]
    df["_db_agency"] = _db_agencies
    df.loc[_has_db, "Rating Src"] = df.loc[_has_db, "_db_agency"]
    # rating_date strings mix formats ("2026-01-29", "July 04, 2025",
    # "2026-03-30T00:00:00+05:30", …); parse per unique value and strip any
    # timezone — pandas 3.x raises on mixed formats in one vectorised
    # to_datetime call, and mixing tz-aware with naive Timestamps degrades
    # the Series to object dtype, which .loc/assignment then rejects.
    _dcache = {}
    def _pdate(s):
        if not s or s in ("nan", "None"):
            return pd.NaT
        if s not in _dcache:
            try:
                _ts = pd.to_datetime(s, errors="coerce")
                if _ts is not pd.NaT and getattr(_ts, "tzinfo", None) is not None:
                    _ts = _ts.tz_localize(None)
                _dcache[s] = _ts
            except (ValueError, TypeError):
                _dcache[s] = pd.NaT
        return _dcache[s]
    _dates = pd.to_datetime(
        pd.Series([_pdate(d) for d in _db_dates], index=df.index),
        errors="coerce")
    df["Rated On"] = _dates.where(_has_db)
    df = df.drop(columns=["_db_rating", "_db_agency"])

# Market-linked debentures: identified from the Coupon Detail text (index /
# equity / commodity / G-sec linked underlyings, as filed with NSDL)
_MLD_RE = re.compile(r"nifty|sensex|index|leap|gold|silver|mcx|basket|"
                     r"underlying|linked|revenue of", re.I)
df["MLD"] = df["Coupon Detail"].fillna("").str.contains(_MLD_RE).map(
    {True: "Yes", False: "No"})

# Stats row
today = date.today()
col1, col2, col3, col4 = st.columns(4)
col1.metric("Upcoming Bonds", f"{len(df):,}")
col2.metric("Maturing in 30 days",  f"{(df['Days to Maturity'] <= 30).sum():,}")
col3.metric("With Yield", f"{df['Yield (%)'].notna().sum():,}")
col4.metric("Data as of", today.strftime("%d %b %Y"))

st.divider()

# ------------------------------------------------------------------ #
# Sidebar filters
# ------------------------------------------------------------------ #
with st.sidebar:
    st.header("Filters")

    if st.button("↺  Refresh Data", type="secondary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # ---- Secondary trades refresh (BSE/NSE) ----
    _sec_meta = _load_secondary_meta()
    _last_ref = str(_sec_meta.get("last_refresh") or "")
    if _last_ref:
        try:
            _last_ref = (datetime.fromisoformat(_last_ref)
                         .strftime("%d %b %Y, %H:%M IST"))
        except ValueError:
            pass
    st.caption(f"Secondary trades last refreshed: {_last_ref or 'never'}")

    if st.button("⟳  Refresh Secondary Trades (BSE/NSE)", type="secondary",
                 use_container_width=True):
        try:
            from gh_app_auth import get_installation_token
            _gh_token = get_installation_token()
            _gh_err = ""
        except Exception as e:
            _gh_token = ""
            _gh_err = str(e)
        if not _gh_token:
            st.error("Couldn't get a GitHub App token to start the refresh. "
                     "Check that GH_APP_PRIVATE_KEY is set in Streamlit secrets. "
                     f"({_gh_err})")
        else:
            try:
                _resp = requests.post(
                    GH_DISPATCH_URL,
                    json={"ref": "main"},
                    headers={
                        "Authorization": f"Bearer {_gh_token}",
                        "Accept": "application/vnd.github+json",
                        "X-GitHub-Api-Version": "2022-11-28",
                        "User-Agent": "creditnexus-bonds-app",
                    },
                    timeout=20,
                )
                if _resp.status_code == 204:
                    st.success("Refresh started — new trades will appear in "
                               "~2 minutes. Click '↺ Refresh Data' after that.")
                else:
                    st.error(f"Workflow dispatch failed "
                             f"({_resp.status_code}): {_resp.text[:200]}")
            except Exception as e:
                st.error(f"Workflow dispatch failed: {e}")

    st.divider()
    # ---- Search ----
    # Pre-fill from deep-link: ?issuer=HDFC
    _qp_issuer = st.query_params.get("issuer", "")
    search_query = st.text_input(
        "Search by ISIN or Issuer", placeholder="e.g. INE002A07HF3, Tata, HDFC",
        value=_qp_issuer,
    ).strip()

    maturity_years = sorted(df["Maturity Date"].dt.year.dropna().unique().astype(int).tolist())
    selected_years = st.multiselect("Maturity Year (blank = all)", maturity_years)

    st.divider()

    # ---- Rating ----
    grade_options = {
        "AAA only":       ["AAA"],
        "AA+ or better":  ["AAA", "AA+"],
        "AA or better":   ["AAA", "AA+", "AA"],
        "AA- or better":  ["AAA", "AA+", "AA", "AA-"],
        "A+ or better":   ["AAA", "AA+", "AA", "AA-", "A+"],
        "A or better":    ["AAA", "AA+", "AA", "AA-", "A+", "A"],
        "A- or better":   ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-"],
        "BBB+ or better": ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+"],
        "All rated":      RATING_GRADES,
        "All (inc. unrated)": None,
    }
    grade_choice = st.selectbox(
        "Minimum Rating",
        options=list(grade_options.keys()),
        index=list(grade_options.keys()).index("All (inc. unrated)"),
    )
    selected_grade_set = grade_options[grade_choice]

    exact_ratings = st.multiselect(
        "Exact Ratings (tick to include only these)",
        options=RATING_GRADES,
        default=[],
        placeholder="e.g. AA-",
        help="Show only bonds whose current rating is one of the ticked "
             "grades (e.g. tick just AA-). Overrides Minimum Rating. "
             "Ratings are cross-referenced against the ratings tracker DB "
             "where the issuer is covered; otherwise the NSDL rating is used.",
    )

    st.divider()

    # ---- MF holdings ----
    only_mf_held = st.checkbox(
        "Only MF-held bonds",
        help="Show only bonds appearing in the latest fortnightly debt-scheme "
             "disclosures of HDFC / SBI / ICICI Pru / Aditya Birla / Axis / "
             "Nippon / Kotak / UTI / Tata MF",
    )

    st.divider()

    # ---- Issue size ----
    st.markdown("**Issue Size (₹ Crores)**")
    sc1, sc2 = st.columns(2)
    min_size = sc1.number_input("Min", min_value=0.0, value=0.0, step=50.0, format="%.0f", label_visibility="collapsed")
    max_size = sc2.number_input("Max (0=no limit)", min_value=0.0, value=0.0, step=500.0, format="%.0f", label_visibility="collapsed")
    sc1.caption("Min Cr")
    sc2.caption("Max Cr (0=no limit)")

    st.divider()

    # ---- Listing status ----
    listed_choice = st.radio(
        "Listing Status",
        options=["All", "Listed only", "Unlisted only"],
        index=0,
        horizontal=True,
    )

    # ---- PSU / sovereign ----
    exclude_psu = st.checkbox(
        "Exclude PSU / Sovereign",
        help="Hide government-owned / public sector issuers. Uses NSDL's "
             "Type of Issuer-Ownership field, backed up by name-based "
             "detection where NSDL leaves it blank or misfiles it.",
    )

    # ---- Market-linked debentures ----
    mld_choice = st.radio(
        "Market-Linked Debentures",
        options=["All", "Exclude MLDs", "Only MLDs"],
        index=0,
        horizontal=True,
        help="MLDs are detected from the coupon text filed with NSDL "
             "(Nifty/Sensex/index/gold/G-sec-linked underlyings — see the "
             "Coupon Detail column).",
    )

    st.divider()

    # ---- Instrument type ----
    with st.expander("Instrument Type", expanded=False):
        all_types = sorted(df["Type"].dropna().unique().tolist())
        selected_types = st.multiselect("Types (blank = all)", all_types, label_visibility="collapsed")

    # ---- Maturity window ----
    with st.expander("Maturity Window", expanded=False):
        max_days = st.slider(
            "Maturing within N days (0 = no limit)",
            min_value=0, max_value=3650, value=0, step=30,
        )

    st.divider()

    # ---- Sectors ----
    available_sectors = sorted(df["Sector"].dropna().unique().tolist())
    available_sectors = [s for s in available_sectors if s]
    selected_sectors = _sector_checkbox_panel(available_sectors)

    st.divider()

    # ---- Sort ----
    st.markdown("**Sort**")
    sort_col = st.selectbox("Sort by", [
        "Maturity Date", "Issue Date", "Issue Size (Cr)", "Coupon Rate (%)",
        "Rating", "Issuer", "Days to Maturity", "Yield (%)"
    ], label_visibility="collapsed")
    sort_asc = st.radio("Order", ["Ascending", "Descending"], horizontal=True) == "Ascending"


# ------------------------------------------------------------------ #
# Apply filters
# ------------------------------------------------------------------ #
filtered = df.copy()

# ISIN / Issuer search
if search_query:
    filtered = filtered[
        filtered["ISIN"].str.contains(search_query, case=False, na=False) |
        filtered["Issuer"].str.contains(search_query, case=False, na=False)
    ]

# Maturity year
if selected_years:
    filtered = filtered[filtered["Maturity Date"].dt.year.isin(selected_years)]

# Rating filter — exact ticked grades take precedence over the preset
if exact_ratings:
    filtered = filtered[filtered["Rating"].isin(set(exact_ratings))]
elif selected_grade_set is not None:
    grade_set = set(selected_grade_set)
    filtered = filtered[filtered["Rating"].isin(grade_set)]

# MF-held only
if only_mf_held and "Holders" in filtered.columns:
    filtered = filtered[filtered["Holders"].notna()]

# Issue size
if min_size > 0:
    filtered = filtered[filtered["Issue Size (Cr)"].fillna(0) >= min_size]
if max_size > 0:
    filtered = filtered[filtered["Issue Size (Cr)"].fillna(0) <= max_size]

# Listing
if listed_choice == "Listed only":
    filtered = filtered[filtered["Listing Status"] == "Listed"]
elif listed_choice == "Unlisted only":
    filtered = filtered[filtered["Listing Status"] == "Unlisted"]

# PSU exclude — effective flag from NSDL ownership field + name heuristic
if exclude_psu:
    filtered = filtered[filtered["PSU"] != "PSU"]

# Market-linked debentures
if mld_choice == "Exclude MLDs":
    filtered = filtered[filtered["MLD"] != "Yes"]
elif mld_choice == "Only MLDs":
    filtered = filtered[filtered["MLD"] == "Yes"]

# Instrument type
if selected_types:
    filtered = filtered[filtered["Type"].isin(selected_types)]

# Maturity window
if max_days > 0:
    filtered = filtered[filtered["Days to Maturity"] <= max_days]

# Sectors — only apply if not all selected
if selected_sectors and len(selected_sectors) < len(available_sectors):
    filtered = filtered[filtered["Sector"].isin(selected_sectors)]

# Sort
filtered = filtered.sort_values(sort_col, ascending=sort_asc, na_position="last")

# ------------------------------------------------------------------ #
# Display table
# ------------------------------------------------------------------ #
st.subheader(f"Upcoming Maturities — {len(filtered):,} bonds")

DISPLAY_COLS = [
    "ISIN", "Issuer", "Coupon Rate (%)", "Coupon Detail", "Issue Date",
    "Maturity Date", "Days to Maturity", "Issue Size (Cr)", "Rating",
    "Rating Src", "Rated On", "Yield (%)", "As of", "Src", "Holders", "Sector"
]
display_df = filtered[DISPLAY_COLS].copy()
# ISIN becomes a link into the yield-history view (?isin=...)
display_df["ISIN"] = display_df["ISIN"].apply(
    lambda i: f"{APP_BASE_URL}/?isin={i}"
)
# Dates and sizes keep their native dtypes (datetime64 / float) so that
# clicking a column header sorts chronologically / numerically — the old
# code formatted them into dd/mm/yyyy and "1,234.00" strings, which made
# header-click sorting lexicographic and effectively random.
display_df["As of"] = pd.to_datetime(display_df["As of"],
                                     format="%d/%m/%Y", errors="coerce")
display_df["Holders"] = display_df["Holders"].str.replace("; ", ", ", regex=False)

st.dataframe(
    display_df,
    use_container_width=True,
    height=580,
    hide_index=True,
    column_config={
        "ISIN":           st.column_config.LinkColumn(
                              width="small",
                              display_text=r"isin=(.*)$",
                              help="Click to see this bond's MF yield trendline over time"),
        "Issuer":         st.column_config.TextColumn(width="large"),
        "Coupon Rate (%)": st.column_config.NumberColumn("Coupon (%)", format="%.2f", width="small",
                                                         help="Fixed coupon rate. Blank for floating-rate / market-linked bonds — see Coupon Detail. 0.00 = zero-coupon."),
        "Coupon Detail":  st.column_config.TextColumn("Coupon Detail", width="small",
                                                      help="Zero Coupon, floating-rate formula, or market-linked underlying (from NSDL's coupon field)"),
        "Issue Date":     st.column_config.DateColumn("Issue Date", format="DD/MM/YYYY", width="small"),
        "Maturity Date":  st.column_config.DateColumn("Maturity Date", format="DD/MM/YYYY", width="small"),
        "Days to Maturity": st.column_config.NumberColumn("Days Left", format="%d", width="small"),
        "Issue Size (Cr)": st.column_config.NumberColumn("Size (Cr)", format="%.2f", width="small",
                                                         help="Issue size in INR crore (from NSDL, filed in rupees). Blank where NSDL's figure is unreliable."),
        "Rating":         st.column_config.TextColumn(width="small",
                                                      help="Current rating — from the ratings tracker DB where the issuer is covered, else NSDL at-issuance"),
        "Rating Src":     st.column_config.TextColumn("Rating Src", width="small",
                                                      help="Agency of the latest rating from the ratings tracker DB; NSDL = at-issuance rating from NSDL"),
        "Rated On":       st.column_config.DateColumn("Rated On", format="DD/MM/YYYY", width="small",
                                                      help="Date of the rating shown — blank where the rating is NSDL's at-issuance value (date unknown, possibly years old)"),
        "Yield (%)":      st.column_config.NumberColumn("Yield (%)", format="%.2f", width="small",
                                                        help="YTM as marked in the latest MF fortnightly debt-scheme disclosure"),
        "As of":          st.column_config.DateColumn("As of", format="DD/MM/YYYY", width="small",
                                                      help="Date of the MF disclosure or BSE/NSE trade the yield is taken from"),
        "Src":            st.column_config.TextColumn("Src", width="small",
                                                      help="Where the yield comes from: AMC name = MF valuation mark; BSE/NSE = actual exchange trade"),
        "Holders":        st.column_config.TextColumn("Holders", width="medium",
                                                      help="Mutual funds holding this bond, with each AMC's total market value in INR crore, e.g. HDFC (50), SBI (100)"),
        "Sector":         st.column_config.TextColumn(width="medium"),
    },
)

# ------------------------------------------------------------------ #
# CSV export
# ------------------------------------------------------------------ #
EXPORT_COLS = [
    "ISIN", "Issuer", "Type", "Series", "Coupon Rate (%)", "Coupon Detail",
    "MLD", "Coupon Type", "Coupon Freq", "Issue Date", "Maturity Date",
    "Days to Maturity",
    "Issue Size (Cr)", "Rating", "Rating Src", "Rated On", "Rating (NSDL)",
    "Rating (Full)", "Yield (%)", "As of",
    "Src", "Holders", "Listing Status", "Sector", "PSU", "Issuer Type",
    "Mode of Issue",
]
export_df = filtered[[c for c in EXPORT_COLS if c in filtered.columns]].copy()
export_df["Issue Date"]    = export_df["Issue Date"].dt.strftime("%d/%m/%Y")
export_df["Maturity Date"] = export_df["Maturity Date"].dt.strftime("%d/%m/%Y")
export_df["Rated On"]      = export_df["Rated On"].dt.strftime("%d/%m/%Y")

csv_bytes = export_df.to_csv(index=False).encode("utf-8")
st.download_button(
    label=f"Download filtered list ({len(filtered):,} bonds) as CSV",
    data=csv_bytes,
    file_name=f"bond_maturities_{today.isoformat()}.csv",
    mime="text/csv",
    use_container_width=True,
)

st.caption(
    "⚠️ Rating shows the latest rating from the ratings tracker DB (Rating Src "
    "= agency) where the issuer is covered; otherwise the NSDL at-issuance "
    "rating (Rating Src = NSDL), which can be dated. "
    "Yield / As of / Holders come from the fortnightly debt-scheme portfolio "
    "disclosures of HDFC, SBI, ICICI Prudential, Aditya Birla SL, Axis, Nippon, "
    "Kotak, UTI and Tata mutual funds; the figure in brackets after each holder "
    "is that AMC's total market value across its schemes in INR crore. Yields "
    "are as per each fund's valuation methodology and are not exchange-traded "
    "levels. Where Src shows BSE/NSE, the yield is a recent exchange trade "
    "that post-dates the MF mark."
)

# ---- Contact footer ----
st.markdown(
    '<div style="text-align:center;margin-top:40px;padding-bottom:16px;'
    'font-size:12px;font-family:\'DM Sans\',sans-serif;color:var(--text-muted,#6B7280)">'
    '<a href="https://www.linkedin.com/in/saxenakriti/" target="_blank"'
    ' style="color:#6366F1;text-decoration:none;font-weight:500">Contact</a></div>',
    unsafe_allow_html=True,
)
