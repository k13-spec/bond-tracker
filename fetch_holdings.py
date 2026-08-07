"""
Fetch NSDL's quarterly investor-category holdings for covered bonds.

NSDL publishes, per ISIN, a quarterly "Statement showing pattern of
Debenture/Bond Holders" with units held by investor category (endpoints
reverse-engineered from the CBDServices Angular bundle, 2026-08-06):

    GET {BASE}/quarterenddate?isin={isin}
        -> [{"year": "2026-2027", "quarter": "Jun 26"}, ...]  (latest first)
    GET {BASE}/isin/{isin}/quarterdate/{Month YY}/publicshareholderdetails
        -> {"reportGenerationDate": "30/06/2026", "shareholdingData": [
             B1 Institutions (Domestic) / B2 (Foreign) / B3 Govt /
             B4 Non-Institutions / Summary Table, each with per-sub-category
             holder counts and units held]}

Scope (maintainer decisions 2026-08-06/07):
  1. bonds carrying a yield in the app — ISINs present in data/mf_yields.csv
     or data/secondary_yields.csv (any sector, incl. FIs/PSUs), plus
  2. every active non-PSU Corporate / Infrastructure NCD (sector-group and
     PSU classification mirror streamlit_app.py — keep the constants below
     IN SYNC with the app when they change there).
Financial-institution and PSU bonds outside the yield-covered set are
deliberately excluded (~14k + ~0.6k), as are convertibles/preference-style
instruments (same purge regexes as the app). Both rules apply automatically
to newly issued bonds on every run.

Output: data/holdings_mix.csv with columns
    isin, as_of, mix     e.g.  "MF 88%, Banks 8%, FPI 4%"
(non-zero categories only, sorted by share, % of the public-holding total).
Side file: data/holdings_misses.csv — ISINs attempted but with no statement
on NSDL, so --new-only backfills skip them instead of re-fetching every
batch. Full refreshes rebuild it (all misses re-checked each quarter).

Usage:
    python fetch_holdings.py                  # refresh ALL targets (quarterly)
    python fetch_holdings.py 30               # smoke test: first 30 targets
    python fetch_holdings.py --new-only 800   # backfill: up to 800 ISINs not
                                              # yet in the CSV (exit 3 = none left)
Run from the bond-tracker repo root. Existing CSV rows are kept for ISINs
that fail to fetch this run (best-effort refresh, no data loss).
"""
import csv
import io
import re
import sys
import time
import warnings
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd
import requests

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "data" / "holdings_mix.csv"
MISSES = ROOT / "data" / "holdings_misses.csv"
BASE = "https://www.indiabondinfo.nsdl.com/bds-service/v1/public/bdsinfo"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Referer": "https://www.indiabondinfo.nsdl.com/CBDServices/",
      "Accept": "application/json"}
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]

# One shared session: connection + TLS reuse. The initial population ran at
# ~1.8 s/ISIN with bare requests.get (fresh TLS context per call); a session
# roughly halves that.
SESSION = requests.Session()
SESSION.headers.update(UA)

# sub-category -> short label (keyword match, first hit wins)
LABELS = [
    ("mutual fund", "MF"),
    ("foreign portfolio", "FPI"),
    ("foreign direct", "FDI"),
    ("foreign venture", "VC"),
    ("venture capital", "VC"),
    ("alternate investment", "AIF"),
    ("bank", "Banks"),
    ("insurance", "Insurance"),
    ("nbfc", "NBFC"),
    ("asset reconstruction", "ARC"),
    ("sovereign wealth", "SWF"),
    ("provident", "PF/Pension"),
    ("pension", "PF/Pension"),
    ("gratuity", "PF/Pension"),
    ("superannuation", "PF/Pension"),
    ("national investment fund", "Govt"),
    ("central government", "Govt"),
    ("state government", "Govt"),
    ("state industrial development", "Govt"),
    ("overseas depositor", "DR holders"),
    ("depositor", "DR holders"),
    ("bodies corporate", "Corporates"),
    ("overseas corporate", "Corporates"),
    ("partnership", "Corporates"),
    ("llp", "Corporates"),
    ("individual", "Individuals"),
    ("huf", "Individuals"),
    ("hindu undivided", "Individuals"),
    ("nri", "NRIs"),
    ("non-resident", "NRIs"),
    ("trust", "Trusts"),
    ("societ", "Trusts"),
    ("employee welfare", "Trusts"),
    ("clearing member", "Clearing members"),
    ("financial institution", "Other FIs"),
]

# ------------------------------------------------------------------ #
# Universe classification — MIRRORS streamlit_app.py (_SECTOR_GROUPS,
# _PSU_FRAGMENTS, _is_psu, _CONVERTIBLE_RE/_PREF_RE). If those change in
# the app, change them here too.
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
        "Renewable Energy",
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
    "indian railway finance", "india infrastructure finance",
    "indian renewable energy development", "nuclear power corporation",
    "national housing bank", "financing infrastructure and development",
    "thdc ", "mtnl", "bsnl", "bharat sanchar", "mahanagar telephone",
    "pnb housing", "ircon", "rites ", "nbcc ", "moil ",
    "mazagon dock", "garden reach ship", "goa shipyard",
    "engineers india", "rashtriya chemicals", "national fertilizers",
    "seci ", "solar energy corporation",
    "power corporation", "energy corporation", "electricity board",
    "state electricity", "rajya vidyut", "prasaran nigam", "vidyut nigam",
    "power distribution company", "power generation co",
    "municipal corporation", "nagar nigam",
    "metropolitan development authority", "capital region development",
    "infrastructure development board", "state beverages",
    "kerala financial corporation", "kerala infrastructure investment",
    "mineral development corporation", "industrial infrastructure corporation",
]

_CONVERTIBLE_RE = re.compile(
    r"\b(CCD|OCD|PCD|FCD)S?\b|COMPULSOR\w* CONVERT|OPTIONAL\w* CONVERT|"
    r"PARTLY CONVERT|FULLY CONVERT", re.I)
_PREF_RE = re.compile(
    r"\bNCRPS\b|\bCRPS\b|\bOCRPS\b|\bRPS\b|\bNCPS\b|PREFERENCE SHARE|PREF\.? SHARE",
    re.I)


def _is_psu(name) -> bool:
    if not isinstance(name, str):
        return False
    n = " " + name.lower() + " "
    return any(f in n for f in _PSU_FRAGMENTS)


def _sector_group(sector: str) -> str:
    for grp, members in _SECTOR_GROUPS.items():
        if sector in members:
            return grp
    return "Corporate"


def short_label(name: str) -> str:
    n = name.lower()
    for kw, lab in LABELS:
        if kw in n:
            return lab
    # fallback: strip the trailing " - B1a"-style code
    return re.sub(r"\s*-\s*B\d\w*$", "", name).strip() or "Others"


def fmt_pct(p: float) -> str:
    return f"{p:.0f}%" if p >= 9.5 else (f"{p:.1f}%".replace(".0%", "%"))


def get(url, **kw):
    last = None
    for attempt in range(3):
        try:
            r = SESSION.get(url, timeout=45, **kw)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r
        except Exception as e:
            last = e
            time.sleep(3 * (attempt + 1))
    raise last


def fetch_one(session_get, isin: str):
    """Return (as_of, mix_string) or None."""
    r = get(f"{BASE}/quarterenddate", params={"isin": isin})
    if r is None or not r.text.startswith("["):
        return None
    quarters = r.json()
    if not quarters:
        return None
    q = (quarters[0].get("quarter") or "").strip()   # latest, e.g. "Jun 26"
    if " " not in q:
        return None
    mon, yy = q.split(" ", 1)
    full = next((m for m in MONTHS if m.startswith(mon[:3])), None)
    if not full:
        return None
    r2 = get(f"{BASE}/isin/{isin}/quarterdate/{full} {yy}/publicshareholderdetails")
    if r2 is None:
        return None
    d = r2.json()
    data = d.get("shareholdingData") or []
    summary = next((h for h in data if "summary" in (h.get("headerName") or "").lower()), None)
    total = (summary or {}).get("totalShareCountHeader")
    if not total:
        total = sum((h.get("totalShareCountHeader") or 0) for h in data
                    if (h.get("headerName") or "").strip().upper().startswith("B"))
    if not total:
        return None
    agg = {}
    for h in data:
        hn = (h.get("headerName") or "").lower()
        if "summary" in hn:
            continue
        for sub in (h.get("listOfSubHeader") or []):
            units = sub.get("totalShareCountSubHeader")
            if units:
                lab = short_label(sub.get("subHeaderName") or "")
                agg[lab] = agg.get(lab, 0) + units
    if not agg:
        return None
    # drop categories that would display as 0% (user wants non-zero only)
    parts = [(lab, u) for lab, u in sorted(agg.items(), key=lambda kv: -kv[1])
             if u / total * 100 >= 0.05]
    if not parts:
        return None
    shown = [f"{lab} {fmt_pct(units / total * 100)}" for lab, units in parts[:6]]
    rest = sum(u for _, u in parts[6:])
    if rest:
        shown.append(f"others {fmt_pct(rest / total * 100)}")
    as_of = d.get("reportGenerationDate") or f"{full} {yy}"
    return as_of, ", ".join(shown)


def target_isins() -> list:
    """Covered universe: yield-carrying ISINs (any sector) plus every active
    non-PSU Corporate/Infrastructure NCD. All from NSDL's active list;
    INE (corporate-issuer) 12-char ISINs only."""
    covered = set()
    for f, col in (("mf_yields.csv", "isin"), ("secondary_yields.csv", "isin")):
        p = ROOT / "data" / f
        if p.exists():
            covered |= set(pd.read_csv(p)[col].astype(str).str.strip())
    covered = {i for i in covered if i.startswith("INE") and len(i) == 12}

    # The active-list download is a ~3 MB xlsx and runs once per invocation —
    # retry transient failures (IncompleteRead killed backfill batch 4 on
    # 2026-08-07) instead of crashing the whole batch.
    last = None
    for attempt in range(3):
        try:
            SESSION.get("https://www.indiabondinfo.nsdl.com/CBDServices/", timeout=20)
            time.sleep(0.3)
            r = SESSION.get(f"{BASE}/listofsecurities?type=Active", timeout=180,
                            headers={"Accept": "*/*"})   # xlsx download, not JSON
            r.raise_for_status()
            break
        except Exception as e:
            last = e
            print(f"  active-list download failed (attempt {attempt + 1}): {e}",
                  file=sys.stderr)
            time.sleep(10 * (attempt + 1))
    else:
        raise last
    warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(io.BytesIO(r.content), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr = {h: i for i, h in enumerate(rows[0]) if h}
    today = date.today()

    def g(row, col):
        i = hdr.get(col)
        v = row[i] if i is not None and i < len(row) else None
        return str(v).strip() if v is not None else ""

    def pdte(v):
        try:
            return datetime.strptime(str(v).strip(), "%d-%m-%Y").date()
        except Exception:
            return None

    targets = set()
    for row in rows[1:]:
        if not row or not row[1]:
            continue
        isin = g(row, "ISIN")
        if not isin.startswith("INE") or len(isin) != 12:
            continue
        md = pdte(g(row, "Date of Redemption/Conversion"))
        if not md or md < today:
            continue
        if isin in covered:
            targets.add(isin)
            continue
        # non-yield bonds qualify only as non-PSU Corporate/Infrastructure NCDs
        desc = g(row, "Security Description")
        if _CONVERTIBLE_RE.search(desc) or _PREF_RE.search(desc):
            continue
        sector = g(row, "Business Sector").split("(")[0].strip()
        if _sector_group(sector) not in ("Corporate", "Infrastructure"):
            continue
        issuer = g(row, "Name of Issuer")
        if g(row, "Type of Issuer-Ownership").startswith("Public Sector") or _is_psu(issuer):
            continue
        targets.add(isin)
    return sorted(targets)


def main():
    args = list(sys.argv[1:])
    new_only = "--new-only" in args
    if new_only:
        args.remove("--new-only")
    limit = int(args[0]) if args else None

    isins = target_isins()
    print(f"total target ISINs: {len(isins)}")

    existing = {}
    if OUT.exists():
        for r in csv.DictReader(open(OUT, encoding="utf-8")):
            existing[r["isin"]] = (r["as_of"], r["mix"])

    # ISINs already attempted with no statement on NSDL. Without this,
    # --new-only batches re-fetch earlier batches' misses before reaching
    # new ground and the tail of the candidate list is never attempted
    # (the bug in the 2026-08-07 initial backfill). Full refreshes rebuild
    # this file from scratch so misses are re-checked every quarter.
    known_misses = set()
    if MISSES.exists():
        for r in csv.DictReader(open(MISSES, encoding="utf-8")):
            known_misses.add(r["isin"])

    if new_only:
        isins = [i for i in isins if i not in existing and i not in known_misses]
        print(f"not yet attempted: {len(isins)}")
        if not isins:
            print("nothing left to backfill.")
            return 3
    if limit:
        isins = isins[:limit]
    print(f"fetching this run: {len(isins)}", flush=True)

    results, ok, miss = dict(existing), 0, 0
    missed_now = []
    for n, isin in enumerate(isins, 1):
        try:
            res = fetch_one(get, isin)
            if res:
                results[isin] = res
                ok += 1
            else:
                miss += 1
                missed_now.append(isin)
        except Exception as e:
            # transient error, NOT a confirmed "no statement" — don't record
            print(f"  ERR {isin}: {e}", file=sys.stderr)
            miss += 1
        if n % 100 == 0:
            print(f"  {n}/{len(isins)} done ({ok} with data)", flush=True)
        time.sleep(0.25)

    # Systemic-failure guard: merge semantics never lose rows, but refuse to
    # claim success if almost nothing came back on a sizable run. Misses
    # don't count against the guard when they're the expected outcome
    # (backfill of never-attempted names has a ~45% hit rate).
    if existing and (ok + miss) >= 50 and ok < 50 and not new_only:
        print(f"ABORT: only {ok} fresh results — keeping existing CSV.", file=sys.stderr)
        return 1
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["isin", "as_of", "mix"])
        for isin in sorted(results):
            as_of, mix = results[isin]
            w.writerow([isin, as_of, mix])
    # Misses file: full refresh rebuilds it (re-check everything quarterly);
    # --new-only appends this run's confirmed no-statement ISINs.
    today_s = date.today().isoformat()
    all_misses = (known_misses | set(missed_now)) if new_only else set(missed_now)
    with open(MISSES, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["isin", "checked_on"])
        for isin in sorted(all_misses):
            w.writerow([isin, today_s])
    print(f"wrote {OUT}: {len(results)} ISINs ({ok} refreshed, {miss} without data this run)")
    print(f"wrote {MISSES}: {len(all_misses)} known no-statement ISINs")
    return 0


if __name__ == "__main__":
    sys.exit(main())
