"""
Fold freshly fetched BSE/NSE secondary trades into the repo's data files:

  data/secondary_yields.csv          — one row per ISIN: the governing trade
                                       (latest trade date; largest Total Trade
                                       Value within that date)
  data/secondary_yields_history.csv  — append-only, one row per (isin, as_of)
  data/secondary_meta.json           — {"last_refresh": <ISO datetime +05:30>}

Usage:  python update_secondary.py <new_trades.csv> [repo_root]

<new_trades.csv> columns: isin, yield, as_of (YYYY-MM-DD), source (BSE/NSE),
trade_value_cr — already filtered to trades >= MIN_TRADE_CR by the fetcher.

Retention rule (per spec): a bond with no MF mark shows a secondary yield only
if traded in the last 30 days, so rows older than 30 days are dropped from the
LATEST file (they remain in history). Bonds WITH MF marks only display the
secondary yield when it is newer than the MF as-of date — the app enforces
that; rows older than 30 days are useless there too (MF marks are at most
~20 days old), so one shared 30-day retention keeps the latest file clean.

An empty trades file is fine (no trades in window / weekend): latest is still
re-pruned and last_refresh still advances.
"""
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd

IST = timezone(timedelta(hours=5, minutes=30))
RETENTION_DAYS = 30
COLS = ["isin", "yield", "as_of", "source", "trade_value_cr"]


def pick_governing(trades: pd.DataFrame) -> pd.DataFrame:
    """Per ISIN: latest as_of wins; within that day, largest trade value."""
    trades = trades.sort_values(["isin", "as_of", "trade_value_cr"],
                                ascending=[True, False, False])
    return trades.drop_duplicates(subset="isin", keep="first")


def main() -> int:
    trades_path = Path(sys.argv[1])
    root = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(".")
    latest_path = root / "data" / "secondary_yields.csv"
    history_path = root / "data" / "secondary_yields_history.csv"
    meta_path = root / "data" / "secondary_meta.json"

    new = pd.read_csv(trades_path) if trades_path.exists() else pd.DataFrame(columns=COLS)
    new = new.reindex(columns=COLS)
    new["isin"] = new["isin"].astype(str).str.strip()
    new = new.dropna(subset=["isin", "yield", "as_of"])
    print(f"new trades: {len(new)} rows across {new['as_of'].nunique()} dates")

    # ---- history: append + dedupe on (isin, as_of) keep newest run --------
    if history_path.exists():
        hist = pd.read_csv(history_path).reindex(columns=COLS)
        before = len(hist)
        hist = pd.concat([hist, new], ignore_index=True)
    else:
        before = 0
        hist = new.copy()
    hist = hist.drop_duplicates(subset=["isin", "as_of"], keep="last")
    hist = hist.sort_values(["isin", "as_of"])
    history_path.parent.mkdir(parents=True, exist_ok=True)
    hist.to_csv(history_path, index=False)
    print(f"history: {before} -> {len(hist)} rows ({len(hist) - before:+d})")

    # ---- latest: merge old latest + new, pick governing trade, prune ------
    if latest_path.exists():
        old = pd.read_csv(latest_path).reindex(columns=COLS)
    else:
        old = pd.DataFrame(columns=COLS)
    combined = pd.concat([old, new], ignore_index=True).dropna(subset=["isin", "yield", "as_of"])
    latest = pick_governing(combined)
    cutoff = (datetime.now(IST) - timedelta(days=RETENTION_DAYS)).strftime("%Y-%m-%d")
    n_before_prune = len(latest)
    latest = latest[latest["as_of"] >= cutoff]
    latest = latest.sort_values("isin")
    latest.to_csv(latest_path, index=False)
    print(f"latest: {len(latest)} ISINs (pruned {n_before_prune - len(latest)} older than {cutoff})")

    # ---- meta ---------------------------------------------------------------
    meta_path.write_text(json.dumps(
        {"last_refresh": datetime.now(IST).isoformat(timespec="seconds")}) + "\n")
    print(f"meta: last_refresh updated -> {meta_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
