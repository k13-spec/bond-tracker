"""
India Bond Maturity Tracker — Streamlit app.

Data is pulled directly from the NSDL public API and cached for 24 hours.
No database or persistent storage required — works out of the box on
Streamlit Community Cloud.

Deploy: push this repo to GitHub, then connect it at share.streamlit.io
"""

import io
import re
import time
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

RATING_GRADES = [
    "AAA", "AA+", "AA", "AA-",
    "A+", "A", "A-",
    "BBB+", "BBB", "BBB-",
    "BB+", "BB", "BB-",
    "B+", "B", "B-",
    "CCC+", "CCC", "CCC-", "CC", "C", "D",
    "A1+", "A1", "A2+", "A2", "A3", "A4",
]
GRADE_PATTERN = re.compile(
    r"\b(" + "|".join(re.escape(g) for g in RATING_GRADES) + r")\b"
)


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
    if raw is None or raw == "":
        return None
    try:
        val = float(str(raw).replace(",", "").strip())
        if val == 0:
            return None
        return round(val / 1e7, 2) if val > 1e6 else round(val, 2)
    except (ValueError, TypeError):
        return None


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
        coupon_raw    = g(row, "Coupon Rate (%)") or ""
        coupon_val    = coupon_raw.rstrip("%").strip()

        records.append({
            "ISIN":            isin,
            "Issuer":          g(row, "Name of Issuer"),
            "Type":            g(row, "Type of Instrument"),
            "Series":          g(row, "Series"),
            "Description":     g(row, "Security Description"),
            "Mode of Issue":   mode_of_issue,
            "Coupon Rate (%)": coupon_val,
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
        today_dt = pd.Timestamp(date.today())
        df["Days to Maturity"] = (df["Maturity Date"] - today_dt).dt.days.astype("Int64")
    return df


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
        "Electric Utilities", "Electricity Generation", "Power Trading",
        "Power - Transmission", "Multi Utilities", "Utilities",
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
    "Financial": [
        "Finance", "Financial Institution",
        "Non-Banking Financial Company (NBFC)", "Housing Finance Company",
        "Private Sector Bank", "Public Sector Bank", "Other Bank",
        "Asset Management Company", "Investment Company",
        "Life Insurance", "General Insurance", "Other Insurance Companies",
        "Insurance Distributors", "Financial Technology (Fintech)",
        "Stockbroking & Allied",
        "Depositories, Clearing Houses and Other Intermediaries",
        "Other Capital Market related Services", "Other Financial Services",
    ],
}

_PSU_FRAGMENTS = [
    "ntpc ", "bhel ", " sail ", "ongc", "iocl", "gail ",
    "nalco", "nmdc", "nhpc", "npcil", "powergrid",
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
    "indian bank ", "uco bank",
    "life insurance corporation",
    "power finance corp", "rural electrification corp",
    "housing and urban development", "national bank for agriculture",
    "export import bank", "exim bank",
    "rec limited", "pfc limited",
    "food corporation of india", "oil india", "mrpl", "bpcl", "hpcl",
]


def _is_psu(name: str) -> bool:
    n = (" " + (name or "").lower() + " ")
    return any(f in n for f in _PSU_FRAGMENTS)


def _sector_group(sector: str) -> str:
    for grp, members in _SECTOR_GROUPS.items():
        if sector in members:
            return grp
    return "Corporate"


def _sector_checkbox_panel(available_sectors: list) -> list:
    grouped: dict[str, list] = {"Corporate": [], "Infrastructure": [], "Financial": []}
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
    for grp in ["Corporate", "Infrastructure", "Financial"]:
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

st.title("📊 India Bond Maturity Tracker")
st.caption("Data: NSDL India Bond Info · Refreshed every 24 hours · Source: indiabondinfo.nsdl.com")

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

# Stats row
today = date.today()
col1, col2, col3, col4 = st.columns(4)
col1.metric("Upcoming Bonds", f"{len(df):,}")
col2.metric("Maturing in 30 days",  f"{(df['Days to Maturity'] <= 30).sum():,}")
col3.metric("Maturing in 90 days",  f"{(df['Days to Maturity'] <= 90).sum():,}")
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

    st.divider()

    # ---- Issuer search ----
    issuer_search = st.text_input(
        "Search Issuer Name", placeholder="e.g. Tata, Reliance, HDFC…"
    ).strip()

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
        help="Hide government-owned / public sector issuers (name-based detection)",
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
        "Maturity Date", "Issue Date", "Issue Size (Cr)", "Rating", "Issuer", "Days to Maturity"
    ], label_visibility="collapsed")
    sort_asc = st.radio("Order", ["Ascending", "Descending"], horizontal=True) == "Ascending"

# ------------------------------------------------------------------ #
# Apply filters
# ------------------------------------------------------------------ #
filtered = df.copy()

# Issuer search
if issuer_search:
    filtered = filtered[filtered["Issuer"].str.contains(issuer_search, case=False, na=False)]

# Rating filter
if selected_grade_set is not None:
    grade_set = set(selected_grade_set)
    filtered = filtered[filtered["Rating"].isin(grade_set)]

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

# PSU exclude
if exclude_psu:
    filtered = filtered[~filtered["Issuer"].apply(_is_psu)]

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
    "ISIN", "Issuer", "Type", "Coupon Rate (%)", "Issue Date", "Maturity Date",
    "Days to Maturity", "Issue Size (Cr)", "Rating", "Listing Status", "Sector"
]
display_df = filtered[DISPLAY_COLS].copy()
display_df["Issue Date"]    = display_df["Issue Date"].dt.strftime("%d/%m/%Y")
display_df["Maturity Date"] = display_df["Maturity Date"].dt.strftime("%d/%m/%Y")
display_df["Issue Size (Cr)"] = display_df["Issue Size (Cr)"].apply(
    lambda x: f"{x:,.2f}" if pd.notna(x) else "—"
)

st.dataframe(
    display_df,
    use_container_width=True,
    height=580,
    hide_index=True,
    column_config={
        "ISIN":           st.column_config.TextColumn(width="small"),
        "Issuer":         st.column_config.TextColumn(width="large"),
        "Days to Maturity": st.column_config.NumberColumn("Days Left", format="%d", width="small"),
        "Issue Size (Cr)": st.column_config.TextColumn("Size (Cr)", width="small"),
        "Rating":         st.column_config.TextColumn(width="small"),
        "Listing Status": st.column_config.TextColumn("Listing", width="small"),
        "Sector":         st.column_config.TextColumn(width="medium"),
    },
)

# ------------------------------------------------------------------ #
# CSV export
# ------------------------------------------------------------------ #
EXPORT_COLS = [
    "ISIN", "Issuer", "Type", "Series", "Coupon Rate (%)", "Coupon Type",
    "Coupon Freq", "Issue Date", "Maturity Date", "Days to Maturity",
    "Issue Size (Cr)", "Rating", "Rating (Full)", "Listing Status",
    "Sector", "Issuer Type", "Mode of Issue",
]
export_df = filtered[[c for c in EXPORT_COLS if c in filtered.columns]].copy()
export_df["Issue Date"]    = export_df["Issue Date"].dt.strftime("%d/%m/%Y")
export_df["Maturity Date"] = export_df["Maturity Date"].dt.strftime("%d/%m/%Y")

csv_bytes = export_df.to_csv(index=False).encode("utf-8")
st.download_button(
    label=f"Download filtered list ({len(filtered):,} bonds) as CSV",
    data=csv_bytes,
    file_name=f"bond_maturities_{today.isoformat()}.csv",
    mime="text/csv",
    use_container_width=True,
)

st.caption(
    "⚠️ Listing Status is derived from Mode of Issue (Public Issue → Listed; "
    "Private Placement → Unlisted). For definitive listing status, check BSE/NSE."
)
