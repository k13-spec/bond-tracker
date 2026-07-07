"""
Fetch corporate-bond secondary-market trades (BSE + NSE) from BSE's
Central Trade Repository API and write them as a normalized CSV.

    python fetch_secondary_trades.py <from YYYY-MM-DD> <to YYYY-MM-DD> <out.csv>

Endpoint (discovered from the trade_repository page component, 2026-07-07):
  https://api.bseindia.com/BseIndiaAPI/api/Mkt_Debt_Trade_SecondaryMarket_beta/w
      ?EXCHANGE_FLAG=&ISIN=&IsSearch=&FromDate=&ToDate=
  - EXCHANGE_FLAG empty  -> BOTH exchanges (the page's "BSE + NSE" choice)
  - FromDate/ToDate empty -> latest trade date only
  - One JSON row per (trade date, ISIN, exchange): AVGWEIGHTEDYIELD is the
    "Yeild% (WAY)#" column, TRADEVALUE is "Total Trade Value* in Rs. Lacs".

Output columns: isin, yield, as_of (YYYY-MM-DD), source (BSE/NSE),
trade_value_cr — one row per (isin, as_of): the row with the largest
trade value that day (so a bigger BSE aggregate beats a smaller NSE one).
Rows below MIN_TRADE_CR (odd-lot retail noise) are dropped.
"""
import sys
import time
from datetime import datetime

import pandas as pd
import requests

API = ("https://api.bseindia.com/BseIndiaAPI/api/"
       "Mkt_Debt_Trade_SecondaryMarket_beta/w")
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36"),
    "Referer": "https://www.bseindia.com/",
    "Accept": "application/json",
}
MIN_TRADE_CR = 1.0        # ignore trades below ₹1 crore (= 100 lacs)


def get_json(params):
    last = None
    for attempt in range(3):
        try:
            r = requests.get(API, params=params, headers=HEADERS, timeout=90)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            last = e
            time.sleep(5 * (attempt + 1))
    raise last


def fetch(fm: str, to: str) -> list:
    """fm/to: YYYY-MM-DD. Tries the date formats the API might accept."""
    fmd = datetime.strptime(fm, "%Y-%m-%d")
    tod = datetime.strptime(to, "%Y-%m-%d")
    candidates = [
        {"EXCHANGE_FLAG": "", "ISIN": "", "IsSearch": "1",
         "FromDate": fmd.strftime("%d/%m/%Y"), "ToDate": tod.strftime("%d/%m/%Y")},
        {"EXCHANGE_FLAG": "", "ISIN": "", "IsSearch": "",
         "FromDate": fmd.strftime("%d/%m/%Y"), "ToDate": tod.strftime("%d/%m/%Y")},
        {"EXCHANGE_FLAG": "", "ISIN": "", "IsSearch": "1",
         "FromDate": fmd.strftime("%d-%b-%Y"), "ToDate": tod.strftime("%d-%b-%Y")},
        # last resort: no dates -> latest trade date only
        {"EXCHANGE_FLAG": "", "ISIN": "", "IsSearch": "",
         "FromDate": "", "ToDate": ""},
    ]
    best_rows, best_dates = [], 0
    for i, params in enumerate(candidates):
        try:
            js = get_json(params)
        except Exception as e:
            print(f"  combo {i}: request failed: {e}", file=sys.stderr)
            continue
        rows = js.get("Table", []) if isinstance(js, dict) else []
        dates = {r.get("TRADE_DATE", "").upper() for r in rows}
        print(f"  combo {i} ({params['FromDate'] or 'default'}): "
              f"{len(rows)} rows, {len(dates)} distinct dates")
        if len(dates) > best_dates or (len(dates) == best_dates and len(rows) > len(best_rows)):
            best_rows, best_dates = rows, len(dates)
        # a multi-day window answered with >1 distinct dates = format worked
        if (tod - fmd).days >= 1 and len(dates) > 1:
            break
        if (tod - fmd).days == 0 and len(rows) > 0:
            break
    return best_rows


def normalize(rows: list) -> pd.DataFrame:
    recs = []
    for r in rows:
        isin = str(r.get("ISIN") or "").strip()
        y = r.get("AVGWEIGHTEDYIELD")
        tv_lacs = r.get("TRADEVALUE")
        dt = str(r.get("TRADE_DATE") or "").strip()
        src = str(r.get("EXCHANGE_FLAG") or "").strip().upper()
        if not isin.startswith("IN") or y is None or tv_lacs is None or not dt:
            continue
        try:
            y = float(y)
            tv_cr = float(tv_lacs) / 100.0
            as_of = datetime.strptime(dt.title(), "%d-%b-%Y").strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            continue
        if not (0 < y <= 60) or tv_cr < MIN_TRADE_CR:
            continue
        recs.append({"isin": isin, "yield": round(y, 4), "as_of": as_of,
                     "source": src or "BSE", "trade_value_cr": round(tv_cr, 2)})
    df = pd.DataFrame(recs, columns=["isin", "yield", "as_of", "source", "trade_value_cr"])
    if df.empty:
        return df
    # one row per (isin, as_of): keep the largest trade value that day
    df = df.sort_values(["isin", "as_of", "trade_value_cr"],
                        ascending=[True, True, False])
    df = df.drop_duplicates(subset=["isin", "as_of"], keep="first")
    return df.sort_values(["isin", "as_of"])


if __name__ == "__main__":
    fm, to, out = sys.argv[1], sys.argv[2], sys.argv[3]
    print(f"fetching BSE+NSE secondary trades {fm} -> {to}")
    rows = fetch(fm, to)
    df = normalize(rows)
    df.to_csv(out, index=False)
    print(f"wrote {out}: {len(df)} (isin, date) rows >= ₹{MIN_TRADE_CR:.0f}cr "
          f"across {df['as_of'].nunique() if not df.empty else 0} dates")
