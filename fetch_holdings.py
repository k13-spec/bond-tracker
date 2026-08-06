"""
Fetch NSDL's quarterly investor-category holdings for yield-covered bonds.

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

Scope: bonds carrying a yield in the app — ISINs present in data/mf_yields.csv
or data/secondary_yields.csv, intersected with NSDL's active list (~2.2k).
The full NCD universe would be ~42k requests; deliberately not fetched.

Output: data/holdings_mix.csv with columns
    isin, as_of, mix     e.g.  "MF 88%, Banks 8%, FPI 4%"
(non-zero categories only, sorted by share, % of the public-holding total).

Usage:  python fetch_holdings.py [limit]
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
BASE = "https://www.indiabondinfo.nsdl.com/bds-service/v1/public/bdsinfo"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Referer": "https://www.indiabondinfo.nsdl.com/CBDServices/",
      "Accept": "application/json"}
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]

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
            r = requests.get(url, headers=UA, timeout=45, **kw)
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
    """Yield-covered ISINs (MF-held or secondary-traded) that are active NSDL debt."""
    covered = set()
    for f, col in (("mf_yields.csv", "isin"), ("secondary_yields.csv", "isin")):
        p = ROOT / "data" / f
        if p.exists():
            covered |= set(pd.read_csv(p)[col].astype(str).str.strip())
    covered = {i for i in covered if i.startswith("INE") and len(i) == 12}

    s = requests.Session()
    s.headers.update({**UA, "Accept": "*/*"})   # xlsx download, not JSON
    s.get("https://www.indiabondinfo.nsdl.com/CBDServices/", timeout=20)
    time.sleep(0.3)
    r = s.get(f"{BASE}/listofsecurities?type=Active", timeout=180)
    r.raise_for_status()
    warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(io.BytesIO(r.content), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr = {h: i for i, h in enumerate(rows[0]) if h}
    today = date.today()

    def pdte(v):
        try:
            return datetime.strptime(str(v).strip(), "%d-%m-%Y").date()
        except Exception:
            return None

    active = set()
    for row in rows[1:]:
        if not row or not row[1]:
            continue
        md = pdte(row[hdr["Date of Redemption/Conversion"]])
        if md and md >= today:
            active.add(str(row[hdr["ISIN"]]).strip())
    return sorted(covered & active)


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else None
    isins = target_isins()
    if limit:
        isins = isins[:limit]
    print(f"target ISINs: {len(isins)}")

    existing = {}
    if OUT.exists():
        for r in csv.DictReader(open(OUT, encoding="utf-8")):
            existing[r["isin"]] = (r["as_of"], r["mix"])

    results, ok, miss = dict(existing), 0, 0
    for n, isin in enumerate(isins, 1):
        try:
            res = fetch_one(get, isin)
            if res:
                results[isin] = res
                ok += 1
            else:
                miss += 1
        except Exception as e:
            print(f"  ERR {isin}: {e}", file=sys.stderr)
            miss += 1
        if n % 100 == 0:
            print(f"  {n}/{len(isins)} done ({ok} with data)")
        time.sleep(0.25)

    if ok < 50 and existing:
        print(f"ABORT: only {ok} fresh results — keeping existing CSV.", file=sys.stderr)
        return 1
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["isin", "as_of", "mix"])
        for isin in sorted(results):
            as_of, mix = results[isin]
            w.writerow([isin, as_of, mix])
    print(f"wrote {OUT}: {len(results)} ISINs ({ok} refreshed, {miss} without data this run)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
