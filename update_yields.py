"""
Fold a freshly built mf_yields snapshot into the repo's two data files:

  data/mf_yields.csv          — latest snapshot only (what the table shows)
  data/mf_yields_history.csv  — every (isin, as_of) data point ever collected
                                (what the per-ISIN trendline is drawn from)

Usage:  python update_yields.py <new_snapshot.csv> [repo_root]

Safety: if the new snapshot looks broken (fewer than MIN_YIELDS ISINs with a
yield — e.g. every AMC download failed), NOTHING is written and the script
exits 1 so the CI run is marked failed instead of silently wiping good data.
"""
import sys
from pathlib import Path

import pandas as pd

MIN_YIELDS = 500          # a healthy fortnight has ~3,000+ ISINs with yield
HISTORY_COLS = ["isin", "as_of", "yield", "source", "holders"]


def main() -> int:
    snapshot_path = Path(sys.argv[1])
    root = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(".")
    latest_path = root / "data" / "mf_yields.csv"
    history_path = root / "data" / "mf_yields_history.csv"

    new = pd.read_csv(snapshot_path)
    n_yields = int(new["yield"].notna().sum())
    as_of_vals = new.loc[new["yield"].notna(), "as_of"].dropna().unique()
    print(f"snapshot: {len(new)} ISINs, {n_yields} with yield, as_of={list(as_of_vals)}")

    if n_yields < MIN_YIELDS:
        print(f"ABORT: only {n_yields} ISINs with yield (< {MIN_YIELDS}). "
              "Snapshot looks broken — leaving repo data untouched.", file=sys.stderr)
        return 1

    # ---- history: append + dedupe on (isin, as_of), keep newest run ----
    # only yield-bearing rows are historical data points (no-yield rows
    # have no as_of date and can't be placed on a trendline)
    new_hist = new[new["yield"].notna()].reindex(columns=HISTORY_COLS).copy()
    new_hist["isin"] = new_hist["isin"].astype(str).str.strip()
    if history_path.exists():
        hist = pd.read_csv(history_path)
        hist = hist.reindex(columns=HISTORY_COLS)
        before = len(hist)
        hist = pd.concat([hist, new_hist], ignore_index=True)
    else:
        before = 0
        hist = new_hist
    hist = hist.drop_duplicates(subset=["isin", "as_of"], keep="last")
    hist = hist.sort_values(["isin", "as_of"])
    history_path.parent.mkdir(parents=True, exist_ok=True)
    hist.to_csv(history_path, index=False)
    print(f"history: {before} -> {len(hist)} rows "
          f"({len(hist) - before:+d}) across "
          f"{hist['as_of'].nunique()} dates -> {history_path}")

    # ---- latest: snapshot replaces the table's data file ----------------
    new.to_csv(latest_path, index=False)
    print(f"latest:  wrote {latest_path} ({len(new)} ISINs)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
