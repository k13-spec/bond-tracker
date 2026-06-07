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
col2.metric("Maturing in 30 days",
            f"{(df['Days to Maturity'] <= 30).sum():,}")
col3.metric("Maturing in 90 days",
            f"{(df['Days to Maturity'] <= 90).sum():,}")
col4.metric("Data as of", today.strftime("%d %b %Y"))

st.divider()

# ------------------------------------------------------------------ #
# Sidebar filters
# ------------------------------------------------------------------ #
with st.sidebar:
    st.header("Filters")

    # Issue size
    size_vals = df["Issue Size (Cr)"].dropna()
    if not size_vals.empty:
        size_min_val = float(size_vals.min())
        size_max_val = float(size_vals.max())
        st.subheader("Issue Size (Crores)")
        min_size = st.number_input("Minimum", min_value=0.0, value=0.0, step=50.0, format="%.0f")
        max_size = st.number_input("Maximum", min_value=0.0, value=size_max_val, step=500.0, format="%.0f")
    else:
        min_size, max_size = 0.0, 1e12

    st.subheader("Rating")
    all_ratings = sorted(df["Rating"].dropna().unique().tolist(), key=lambda r: RATING_GRADES.index(r) if r in RATING_GRADES else 99)
    selected_ratings = st.multiselect("Select ratings (leave blank = all)", all_ratings)

    st.subheader("Listing Status")
    listing_opts = ["All"] + sorted(df["Listing Status"].dropna().unique().tolist())
    selected_listing = st.selectbox("", listing_opts, index=0)

    st.subheader("Instrument Type")
    all_types = sorted(df["Type"].dropna().unique().tolist())
    selected_types = st.multiselect("Select types (leave blank = all)", all_types)

    st.subheader("Maturity Window")
    max_days = st.slider(
        "Show bonds maturing within N days (0 = no limit)",
        min_value=0, max_value=3650, value=0, step=30
    )

    st.divider()
    st.subheader("Sort")
    sort_col = st.selectbox("Sort by", [
        "Maturity Date", "Issue Date", "Issue Size (Cr)", "Rating", "Issuer", "Days to Maturity"
    ])
    sort_asc = st.radio("Order", ["Ascending", "Descending"]) == "Ascending"

    st.divider()
    if st.button("Force Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ------------------------------------------------------------------ #
# Apply filters
# ------------------------------------------------------------------ #
filtered = df.copy()

if min_size > 0:
    filtered = filtered[filtered["Issue Size (Cr)"].fillna(0) >= min_size]
if max_size < float(df["Issue Size (Cr)"].max() or 1e12):
    filtered = filtered[filtered["Issue Size (Cr)"].fillna(0) <= max_size]
if selected_ratings:
    filtered = filtered[filtered["Rating"].isin(selected_ratings)]
if selected_listing != "All":
    filtered = filtered[filtered["Listing Status"] == selected_listing]
if selected_types:
    filtered = filtered[filtered["Type"].isin(selected_types)]
if max_days > 0:
    filtered = filtered[filtered["Days to Maturity"] <= max_days]

# Sort
sort_col_map = {
    "Maturity Date": "Maturity Date",
    "Issue Date": "Issue Date",
    "Issue Size (Cr)": "Issue Size (Cr)",
    "Rating": "Rating",
    "Issuer": "Issuer",
    "Days to Maturity": "Days to Maturity",
}
filtered = filtered.sort_values(
    sort_col_map[sort_col],
    ascending=sort_asc,
    na_position="last",
)

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
        "ISIN": st.column_config.TextColumn(width="small"),
        "Issuer": st.column_config.TextColumn(width="large"),
        "Days to Maturity": st.column_config.NumberColumn(
            "Days Left", format="%d", width="small"
        ),
        "Issue Size (Cr)": st.column_config.TextColumn("Size (Cr)", width="small"),
        "Rating": st.column_config.TextColumn(width="small"),
        "Listing Status": st.column_config.TextColumn("Listing", width="small"),
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
export_df["Issue Date"]    = export_df["Issue Date"].dt.strftime("%d/%m/%Y") if "Issue Date" in export_df else ""
export_df["Maturity Date"] = export_df["Maturity Date"].dt.strftime("%d/%m/%Y") if "Maturity Date" in export_df else ""

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
