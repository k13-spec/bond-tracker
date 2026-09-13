"""
Backfill BSE/NSE secondary-market trade yields into
data/secondary_yields_history.csv for a past date range, WITHOUT touching
data/secondary_yields.csv (governing trade per ISIN) or
data/secondary_meta.json (last-refresh stamp shown on the dashboard).

    python backfill_secondary.py [--start YYYY-MM-DD] [--end YYYY-MM-DD]
                                 [--chunk-days N] [--root .]

Defaults (maintainer request 2026-09-13: "yields from 1 Jan 2026, without
redoing dates the database already has"):
  --start  2026-01-01
  --end    the day before the earliest as_of already in history
           (i.e. only the gap in front of the existing data is fetched)
  --chunk-days 31   one API call per ~month; the CTR API answered a 30-day
                    window with ~19k raw rows in 2026-07, so a month is safe

Skip logic (so reruns never redo work):
  * chunks whose weekdays are already >= 80 % present in history are skipped
    (the 20 % slack covers exchange holidays without a holiday calendar);
  * merged rows never overwrite an existing (isin, as_of) row: the daily
    refresh's picks stay authoritative, backfill only fills holes.

Reuses fetch()/normalize() from fetch_secondary_trades.py (>= Rs 1cr filter,
largest trade per (isin, date)).
"""
import argparse
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

from fetch_secondary_trades import fetch, normalize, MIN_TRADE_CR

COLS = ["isin", "yield", "as_of", "source", "trade_value_cr"]


def weekdays(a: date, b: date) -> set:
    d, out = a, set()
    while d <= b:
        if d.weekday() < 5:
            out.add(d.isoformat())
        d += timedelta(days=1)
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", default="2026-01-01")
    ap.add_argument("--end", default="")
    ap.add_argument("--chunk-days", type=int, default=31)
    ap.add_argument("--root", default=".")
    a = ap.parse_args()

    history_path = Path(a.root) / "data" / "secondary_yields_history.csv"
    if history_path.exists():
        hist = pd.read_csv(history_path).reindex(columns=COLS)
        hist["isin"] = hist["isin"].astype(str).str.strip()
    else:
        hist = pd.DataFrame(columns=COLS)
    have_dates = set(hist["as_of"].dropna().astype(str))
    print(f"history: {len(hist)} rows, {len(have_dates)} dates"
          + (f" ({min(have_dates)} -> {max(have_dates)})" if have_dates else ""))

    start = datetime.strptime(a.start, "%Y-%m-%d").date()
    if a.end:
        end = datetime.strptime(a.end, "%Y-%m-%d").date()
    elif have_dates:
        end = datetime.strptime(min(have_dates), "%Y-%m-%d").date() - timedelta(days=1)
    else:
        end = date.today() - timedelta(days=1)
    if end < start:
        print(f"nothing to do: history already starts on/before {a.start}")
        return 0
    print(f"backfill window: {start} -> {end}")

    # ---- chunked fetch ----------------------------------------------------
    new_frames, cur = [], start
    while cur <= end:
        nxt = min(cur + timedelta(days=a.chunk_days - 1), end)
        wd = weekdays(cur, nxt)
        covered = len(wd & have_dates) / len(wd) if wd else 1.0
        if covered >= 0.8:
            print(f"[{cur} -> {nxt}] skip: {covered:.0%} of weekdays already in history")
            cur = nxt + timedelta(days=1)
            continue
        print(f"[{cur} -> {nxt}] fetching ({covered:.0%} covered) ...")
        try:
            rows = fetch(cur.isoformat(), nxt.isoformat())
        except Exception as e:                       # keep going; rerun fills the hole
            print(f"[{cur} -> {nxt}] FAILED: {e}", file=sys.stderr)
            cur = nxt + timedelta(days=1)
            continue
        df = normalize(rows)
        # defensive: the API's no-date fallback returns the latest day only;
        # drop anything outside the requested chunk
        df = df[(df["as_of"] >= cur.isoformat()) & (df["as_of"] <= nxt.isoformat())]
        print(f"[{cur} -> {nxt}] {len(rows)} raw rows -> {len(df)} (isin, date) rows "
              f">= Rs {MIN_TRADE_CR:.0f}cr across {df['as_of'].nunique()} dates")
        new_frames.append(df)
        cur = nxt + timedelta(days=1)

    new = pd.concat(new_frames, ignore_index=True) if new_frames else pd.DataFrame(columns=COLS)

    # ---- merge: existing rows win, backfill only fills holes ---------------
    before = len(hist)
    merged = pd.concat([hist, new], ignore_index=True)
    merged = merged.drop_duplicates(subset=["isin", "as_of"], keep="first")
    merged = merged.sort_values(["isin", "as_of"])
    history_path.parent.mkdir(parents=True, exist_ok=True)
    merged.to_csv(history_path, index=False)
    added = len(merged) - before
    print(f"history: {before} -> {len(merged)} rows (+{added}); "
          f"{merged['as_of'].nunique()} dates, {merged['isin'].nunique()} ISINs")

    # ---- coverage report (trading days captured per month) -----------------
    m = pd.to_datetime(merged["as_of"]).dt.to_period("M")
    per_month = merged.groupby(m)["as_of"].nunique()
    print("trading days with trades per month:")
    for k, v in per_month.items():
        print(f"  {k}: {v}")
    # exported for the workflow's commit message
    Path("backfill_added.txt").write_text(str(added))
    return 0


if __name__ == "__main__":
    sys.exit(main())
