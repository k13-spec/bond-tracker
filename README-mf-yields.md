# MF Yield Enrichment — how it works and how to refresh

## What was added to the Bond Tracker

Three new columns in the dashboard table (and CSV export), merged by **ISIN**:

- **Yield (%)** — YTM of the bond as marked in a mutual fund's fortnightly debt-scheme
  portfolio disclosure (per the fund's valuation methodology).
- **As of** — date of the disclosure the yield was taken from (15th or month-end).
- **Holders** — every one of the 8 tracked AMCs whose disclosure contains the ISIN.

Also: the "↗ Ratings" hyperlink column was removed — the NSDL rating now just reads
as plain text in the **Rating** column. New sidebar filter "Only MF-held bonds",
new stat "With MF Yield", and "Yield (%)" added as a sort option.

## Data pipeline

1. `fetch_disclosures.py 2026-06-30 disclosures/` — downloads the latest fortnightly
   disclosures from HDFC, SBI, ICICI Pru, Aditya Birla SL, Axis, Nippon, Kotak, UTI
   into `disclosures/`. (Run on a normal machine with open internet. Axis sometimes
   needs a manual download — the script tells you when.)
2. `build_mf_yields.py disclosures/` — parses all workbooks and writes
   `disclosures/mf_yields.csv`.
   - Yields come from the **first** AMC that discloses the ISIN, in this priority
     order: HDFC → SBI → ICICI Pru → Aditya Birla → Axis → Nippon → Kotak → UTI.
     Later AMCs never overwrite the yield; they only append to Holders.
   - Sheets that store yield as a decimal fraction (0.073 = 7.3%) are auto-detected
     and scaled.
3. Commit the refreshed CSV to `data/mf_yields.csv` in the `bond-tracker` repo.
   The live app reloads it within an hour (1-hour cache).

## Refresh cadence

SEBI requires debt-scheme portfolio disclosure within 5 days of each fortnight end,
so refresh on the **5th** (captures month-end marks) and **20th** (captures 15th
marks) of each month. A Cowork scheduled task is set up to do this automatically;
it re-downloads the disclosures, rebuilds the CSV, and pushes it to GitHub.

## Files in this folder

| File | Goes where |
|---|---|
| `streamlit_app.py` | replaces `streamlit_app.py` in `k13-spec/bond-tracker` |
| `data/mf_yields.csv` | new file `data/mf_yields.csv` in `k13-spec/bond-tracker` |
| `build_mf_yields.py` | new file in `k13-spec/bond-tracker` (parser) |
| `fetch_disclosures.py` | new file in `k13-spec/bond-tracker` (downloader) |

## Caveats

- MF yields are valuation marks, not traded levels; different AMCs can mark the
  same bond a few bps apart. The dashboard shows the priority-order AMC's mark.
- Only bonds held by at least one of the 8 AMCs get a yield (~3,800 ISINs as of
  30-Jun-2026, of which the corporate NCD subset overlaps the NSDL active list).
- HDFC publishes per-scheme files; its scheme list is hardcoded in
  `fetch_disclosures.py` and needs updating if HDFC adds/merges debt schemes.
