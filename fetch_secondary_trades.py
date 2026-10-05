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

Transport (2026-10-05): Akamai in front of api.bseindia.com started 403-ing
the GitHub runner's requests on ~23 Sep; the 2026-10-01 browser-header fix
did not help, so the block is on the TLS fingerprint (python-requests') or
IP reputation, not headers. get_json() now tries, in order, remembering the
first transport that works for subsequent calls:
  1. plain requests  (works locally / wherever we are not bot-flagged)
  2. curl_cffi with Chrome TLS impersonation (defeats TLS fingerprinting;
     installed at runtime if the workflow's install step predates this fix)
  3. curl_cffi with an Akamai cookie-priming visit to bseindia.com first
Every failure is recorded in the diagnostics that update_secondary.py folds
into data/secondary_meta.json["last_fetch"], so if all three 403 (= hard IP
block), the meta file says so and the remaining options are a self-hosted /
local run or a different data source.
"""
import subprocess
import sys
import time
from datetime import datetime

import pandas as pd
import requests

API = ("https://api.bseindia.com/BseIndiaAPI/api/"
       "Mkt_Debt_Trade_SecondaryMarket_beta/w")
HOMEPAGE = "https://www.bseindia.com/"
# Full browser-like header set for the plain-requests transport.
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/140.0.0.0 Safari/537.36"),
    "Referer": "https://www.bseindia.com/",
    "Origin": "https://www.bseindia.com",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate",
    "Connection": "keep-alive",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-site",
    "sec-ch-ua": '"Chromium";v="140", "Google Chrome";v="140", "Not;A=Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
}
# For curl_cffi, impersonation supplies UA/sec-ch/TLS itself — only add what
# the page's own XHR adds. Overriding the UA would desync it from the TLS hello.
CFFI_HEADERS = {
    "Referer": "https://www.bseindia.com/",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
}
MIN_TRADE_CR = 1.0        # ignore trades below ₹1 crore (= 100 lacs)
CHUNK_DAYS = 31           # one API call per ~month when the window has grown
DIAG = []                 # per-attempt diagnostics, folded into secondary_meta.json
DIAG_FILE = "fetch_diag.json"


def _note(**kw):
    """Record one request attempt (status / error / body head) for the meta file."""
    kw["at"] = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    DIAG.append(kw)
    print(f"  diag: {kw}", file=sys.stderr)


# --------------------------- transports -------------------------------------

_curl_cffi_ready = None      # None = not yet checked
_cffi_session = None         # cookie-primed curl_cffi session, built lazily
_preferred = None            # first transport that worked this run


def _ensure_curl_cffi():
    """Import curl_cffi, pip-installing it at runtime if the environment's
    install step predates this fix. Fails soft: requests transport still runs."""
    global _curl_cffi_ready
    if _curl_cffi_ready is not None:
        return _curl_cffi_ready
    try:
        import curl_cffi  # noqa: F401
        _curl_cffi_ready = True
        return True
    except ImportError:
        pass
    try:
        print("  installing curl_cffi (runtime bootstrap)...", file=sys.stderr)
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q",
                               "curl_cffi"])
        import curl_cffi  # noqa: F401
        _curl_cffi_ready = True
    except Exception as e:
        _note(transport="curl_cffi", error=f"install failed: "
              f"{type(e).__name__}: {str(e)[:150]}")
        _curl_cffi_ready = False
    return _curl_cffi_ready


def _check(r, transport):
    """Shared response handling: note non-200s and non-JSON, return parsed JSON."""
    if r.status_code != 200:
        _note(transport=transport, status=r.status_code,
              server=r.headers.get("Server", ""),
              body=r.text[:200].replace("\n", " "))
        r.raise_for_status()
    try:
        return r.json()
    except ValueError:
        _note(transport=transport, status=r.status_code, error="non-JSON body",
              ctype=r.headers.get("Content-Type", ""),
              body=r.text[:200].replace("\n", " "))
        raise


def _requests_get(params):
    r = requests.get(API, params=params, headers=HEADERS, timeout=90)
    return _check(r, "requests")


def _cffi_get(params):
    from curl_cffi import requests as creq
    r = creq.get(API, params=params, headers=CFFI_HEADERS,
                 impersonate="chrome", timeout=90)
    return _check(r, "curl_cffi")


def _cffi_primed_get(params):
    """Akamai sets bot-manager cookies on the site page; carry them to the API."""
    global _cffi_session
    from curl_cffi import requests as creq
    if _cffi_session is None:
        s = creq.Session(impersonate="chrome")
        home = s.get(HOMEPAGE, timeout=60)
        _note(transport="curl_cffi_primed", step="homepage",
              status=home.status_code,
              cookies=",".join(sorted(s.cookies.keys())[:6]))
        _cffi_session = s
    r = _cffi_session.get(API, params=params, headers=CFFI_HEADERS, timeout=90)
    return _check(r, "curl_cffi_primed")


def get_json(params):
    """Try each transport (preferred-first) twice around, 3s between failures."""
    global _preferred
    transports = [("requests", _requests_get)]
    if _ensure_curl_cffi():
        transports += [("curl_cffi", _cffi_get),
                       ("curl_cffi_primed", _cffi_primed_get)]
    if _preferred:
        transports.sort(key=lambda t: t[0] != _preferred)
    last = None
    for _round in range(2):
        for name, fn in transports:
            try:
                js = fn(params)
                if _preferred != name:
                    _note(transport=name, ok=True)
                    _preferred = name
                return js
            except Exception as e:
                last = e
                time.sleep(3)
    raise last


# ------------------------- fetch + normalize --------------------------------

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


def last_captured_trade_date(history_csv="data/secondary_yields_history.csv"):
    """Latest as_of already in history (YYYY-MM-DD), or None."""
    try:
        h = pd.read_csv(history_csv, usecols=["as_of"])
        return str(h["as_of"].dropna().max())
    except Exception:
        return None


if __name__ == "__main__":
    import json
    from datetime import timedelta

    fm, to, out = sys.argv[1], sys.argv[2], sys.argv[3]

    # Self-healing window (2026-10-01): the workflow derives `fm` from the last
    # RUN time, so after an API outage the window would slide past the days
    # that were never captured. Pull `fm` back to the last trade date we hold.
    last_trade = last_captured_trade_date()
    if last_trade and last_trade < fm:
        print(f"window start {fm} -> {last_trade} (last captured trade date)")
        fm = last_trade

    print(f"fetching BSE+NSE secondary trades {fm} -> {to}")
    fmd = datetime.strptime(fm, "%Y-%m-%d").date()
    tod = datetime.strptime(to, "%Y-%m-%d").date()
    rows, failed_chunks, cur = [], 0, fmd
    while cur <= tod:
        nxt = min(cur + timedelta(days=CHUNK_DAYS - 1), tod)
        try:
            rows += fetch(cur.isoformat(), nxt.isoformat())
        except Exception as e:
            failed_chunks += 1
            print(f"  chunk {cur} -> {nxt} failed: {e}", file=sys.stderr)
        cur = nxt + timedelta(days=1)
    df = normalize(rows)
    df.to_csv(out, index=False)
    print(f"wrote {out}: {len(df)} (isin, date) rows >= ₹{MIN_TRADE_CR:.0f}cr "
          f"across {df['as_of'].nunique() if not df.empty else 0} dates")

    # Diagnostics for update_secondary.py -> secondary_meta.json["last_fetch"]
    diag = {
        "window": [fm, to],
        "raw_rows": len(rows),
        "rows": int(len(df)),
        "failed_chunks": failed_chunks,
        "ok": bool(len(rows) > 0),
        "transport_used": _preferred,
        "attempts": DIAG[:12],
    }
    with open(DIAG_FILE, "w") as fh:
        json.dump(diag, fh, default=str)
