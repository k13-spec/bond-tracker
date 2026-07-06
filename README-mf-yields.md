# MF Yield Enrichment — how it works and how it refreshes itself

## What the dashboard shows

Three columns in the table (and CSV export), merged by **ISIN**:

- **Yield (%)** — YTM of the bond as marked in a mutual fund's fortnightly debt-scheme
  portfolio disclosure (per the fund's valuation methodology).
- **As of** — date of the disclosure the yield was taken from (15th or month-end).
- **Holders** — every one of the 8 tracked AMCs whose disclosure contains the ISIN.

**Clicking an ISIN** opens that bond's yield-history view (`?isin=INE...`): a
trendline of every fortnightly yield mark collected so far, with latest yield,
change vs the previous fortnight, and the full data-point table. The trendline
grows by one point per fortnight as history accumulates.

## Two data files

| File | Contents |
|---|---|
| `data/mf_yields.csv` | **Latest fortnight only** — what the main table shows |
| `data/mf_yields_history.csv` | **Append-only archive** — one row per (ISIN, as_of) per fortnight; feeds the trendline. Seeded 30-Jun-2026. |

## Fully automated refresh (no human intervention)

The GitHub Actions workflow `.github/workflows/mf-yields-refresh.yml` runs on
GitHub's servers at **10:00 IST on the 5th and 20th** of each month (SEBI requires
disclosure within 5 days of each fortnight end, so the 5th captures month-end
marks and the 20th captures 15th marks). Catch-up runs on the **6th and 21st**
re-try any AMC files that were published late — if nothing changed, they commit
nothing.

Each run:

1. `fetch_disclosures.py <fortnight-end> disclosures/` — downloads the disclosures
   of HDFC, SBI, ICICI Pru, Aditya Birla SL, Axis, Nippon, Kotak, UTI
   (transient network errors are retried; per-AMC misses are logged, not fatal).
2. `build_mf_yields.py disclosures/ <fortnight-end>` — parses all workbooks into a
   snapshot CSV. Yields come from the **first** AMC that discloses the ISIN, in
   priority order HDFC → SBI → ICICI Pru → Aditya Birla → Axis → Nippon → Kotak →
   UTI; later AMCs only append to Holders. Fraction-style yields (0.073 = 7.3%)
   are auto-detected per sheet and scaled.
3. `update_yields.py disclosures/mf_yields.csv .` — replaces `data/mf_yields.csv`
   and appends the snapshot's yield rows to `data/mf_yields_history.csv`
   (deduped on ISIN + as_of, so re-runs are harmless).
   **Safety valve:** if the snapshot has fewer than 500 yields (e.g. every AMC
   download failed), nothing is written and the job fails loudly instead of
   wiping good data.
4. Commits both CSVs as `github-actions[bot]`. The live app picks the change up
   within its 1-hour cache.

Monitoring: GitHub emails the repo owner on workflow failure, and a Cowork
scheduled task double-checks each run at 14:00 IST and alerts only if something
went wrong. To run a refresh manually: repo → Actions → "MF yields fortnightly
refresh" → Run workflow (optionally passing an explicit fortnight-end date).

## Caveats

- MF yields are valuation marks, not traded levels; different AMCs can mark the
  same bond differently. The source AMC of each mark is recorded in the `source`
  column of both CSVs and shown in the trendline tooltips.
- Per-AMC URL/API quirks live in `fetch_disclosures.py` (Kotak & UTI have JSON
  APIs; Axis has no stable URL — the last-known pattern is tried, else that
  fortnight simply lacks Axis data; HDFC is 39 per-scheme files with predictable
  names — update `HDFC_SCHEMES` if HDFC launches/merges debt schemes; Nippon's
  file naming is inconsistent, two variants are tried).
- If AMC websites ever start blocking GitHub's IP ranges, the workflow's job
  summary will show the misses; the fallback is running the same three scripts
  on any normal machine and committing the two CSVs.
