"""
Download the latest fortnightly debt-scheme portfolio disclosures from
9 AMCs into the folder layout expected by build_mf_yields.py:

    <out>/hdfc/*.xlsx   <out>/sbi/sbi.xlsx     <out>/icici/*.xlsx (unzipped)
    <out>/absl/*.xls*   <out>/axis/axis.xlsx   <out>/nippon/nippon.xls
    <out>/kotak/kotak.xlsx                     <out>/uti/Sebi Exposure*.xlsx
    <out>/tata/tata.xlsx

Usage:  python fetch_disclosures.py 2026-06-30 [out_dir]
        (date = the fortnight-end being fetched: 15th or month-end)

Per-AMC quirks (re-verified 2026-08-20):
  - HDFC: one xlsx per debt scheme, predictable S3 URLs; update HDFC_SCHEMES
    when schemes launch/merge/mature. Month in filename switched from full
    ("15-August-2026") to abbreviated ("15-Aug-2026") with the Aug-2026
    fortnight — both formats are probed.
  - ABSL: filename convention changes almost every fortnight — several
    variants are probed (zip and xls).
  - Axis: CMS API (POST /cms/get-scheme-documents) with a static public
    Bearer token returns the exact document URL; legacy numeric-path probe
    kept as fallback.
  - Kotak: their listing API sits behind Radware bot-protection, but the
    S3 file host is open and the path is predictable — hit it directly.
  - UTI: JSON API needs browser-ish Accept/Referer headers and can 502
    transiently; retried.
  - Tata: Drupal CMS, predictable "Fortnightly Portfolio as on ..." URL in
    the publish-month folder (file for the 31st lands in next month's folder).
"""
import io
import json
import re
import sys
import time
import zipfile
from datetime import date, datetime
from pathlib import Path

import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}

# Static public token baked into axismf.com's frontend bundle (captured
# 2026-08-05); works without cookies. If Axis rotates it, re-capture from
# the Authorization header of any /cms/ call on axismf.com.
AXIS_CMS_TOKEN = ("c060dc4235de5fefc8fe5da8ef2b64d59fdf4f46c8ebeddb394a47daeac8c67c"
                  "083d602ed9d4133d32b50ce33241fbedb6240c94cc801279292b3f301ae1ef6f"
                  "713e38c38d778f9a7ec84bd4c094c0b5fa3cd8b3c5e9d5ae43b9a47ddcfe60b6"
                  "339fe8395818d3f21ffaaaca455fe03e48b47a5079bf4a2eb86fece310b253ff")

# HDFC publishes one xlsx per debt scheme. Scheme list as of Aug-2026;
# update if HDFC launches/merges debt schemes.
HDFC_SCHEMES = [
    "HDFC Ultra Short Term Fund", "HDFC Short Term Debt Fund",
    "HDFC Retirement Savings Fund - Hybrid-Debt Plan", "HDFC Overnight Fund",
    "HDFC NIFTY SDL Plus G-Sec Jun 2027 40.60 Index Fund",
    "HDFC Nifty SDL Oct 2026 Index Fund", "HDFC Nifty G-Sec Sep 2032 Index Fund",
    "HDFC Nifty G-Sec Jun 2036 Index Fund", "HDFC Nifty G-Sec Jun 2027 Index Fund",
    "HDFC Nifty G-Sec July 2031 Index Fund", "HDFC Nifty G-Sec Dec 2026 Index Fund",
    "HDFC NIFTY G-Sec Apr 2029 Index Fund", "HDFC NIFTY 1D RATE LIQUID ETF",
    "HDFC Multi-Asset Active FOF", "HDFC Money Market Fund",
    "HDFC Medium Term Debt Fund", "HDFC Low Duration Fund",
    "HDFC Long Duration Debt Fund", "HDFC Liquid Fund",
    "HDFC Income Plus Arbitrage Omni FOF", "HDFC Income Plus Arbitrage Active FOF",
    "HDFC Income Fund", "HDFC Hybrid Debt Fund", "HDFC Gilt Fund",
    "HDFC FMP 2638D February 2023", "HDFC FMP 1876D March 2022",
    "HDFC FMP 1861D March 2022", "HDFC FMP 1269D March 2023",
    "HDFC Floating Rate Debt Fund", "HDFC Dynamic Debt Fund",
    "HDFC Diversified Equity All Cap Active FOF",
    "HDFC CRISIL-IBX Financial Services 9-12 Months Debt Index Fund",
    "HDFC CRISIL-IBX Financial Services 3-6 Months Debt Index Fund",
    "HDFC Credit Risk Debt Fund", "HDFC Corporate Bond Fund",
    "HDFC Charity Fund for Cancer Cure",
    "HDFC Banking and PSU Debt Fund",
]


def get(url, **kw):
    """GET with 3 attempts on transient failures (connection errors / 5xx).

    4xx raises immediately — a 404 won't heal on retry, and the HDFC/Nippon
    URL probing relies on fast failure for wrong candidate URLs.
    """
    last = None
    for attempt in range(3):
        try:
            r = requests.get(url, headers=UA, timeout=60, **kw)
            r.raise_for_status()
            return r
        except requests.HTTPError as e:
            if e.response is not None and 400 <= e.response.status_code < 500:
                raise
            last = e
        except Exception as e:
            last = e
        time.sleep(5 * (attempt + 1))
    raise last


def save(path: Path, content: bytes):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    print(f"  saved {path} ({len(content):,} bytes)")


def try_urls(urls):
    for u in urls:
        try:
            r = get(u)
            if len(r.content) > 5000:
                return r.content, u
        except Exception:
            continue
    return None, None


def fetch(d: date, out: Path):
    day = d.day                     # 15 or month-end
    mon_full = d.strftime("%B")     # June
    mon_abbr = d.strftime("%b")     # Jun
    yyyy = d.strftime("%Y")
    yy = d.strftime("%y")
    dd = f"{d.day:02d}"
    mm = f"{d.month:02d}"
    ordsfx = {1: "st", 2: "nd", 3: "rd", 21: "st", 22: "nd", 23: "rd", 31: "st",
              15: "th", 30: "th"}.get(day, "th")
    # publish-month folders: same month, then next month (month-end files
    # are usually uploaded in the first days of the following month)
    nxt = date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    folders = [d.strftime("%Y-%m"), nxt.strftime("%Y-%m")]

    # ---- HDFC (per-scheme files) -------------------------------------
    print("HDFC…")
    for scheme in HDFC_SCHEMES:
        content = None
        # abbreviated month first (current convention since Aug-2026),
        # full month as fallback (convention up to Jul-2026)
        for m in (mon_abbr, mon_full):
            fname = f"{scheme} - {dd}-{m}-{yyyy}.xlsx"
            urls = [f"https://files.hdfcfund.com/s3fs-public/{f}/{requests.utils.quote(fname)}"
                    for f in reversed(folders)]
            content, _ = try_urls(urls)
            if content:
                save(out / "hdfc" / fname, content)
                break
        if not content:
            print(f"  MISS {scheme} - {dd}-{mon_abbr}|{mon_full}-{yyyy}.xlsx")

    # ---- SBI ----------------------------------------------------------
    print("SBI…")
    u = (f"https://www.sbimf.com/docs/default-source/scheme-portfolios/"
         f"debt-schemes-fortnightly-portfolio---as-on-{day}{ordsfx}-{mon_full.lower()}-{yyyy}.xlsx")
    content, _ = try_urls([u])
    if content: save(out / "sbi" / "sbi.xlsx", content)
    else: print("  MISS", u)

    # ---- ICICI Prudential (zip of per-scheme xlsx) ---------------------
    print("ICICI Pru…")
    u = (f"https://www.icicipruamc.com/blob/downloads/Files/Fortnightly Portfolio "
         f"Disclosures/{yyyy}/Fortnightly Debt Scheme Portfolio - {day}{ordsfx} {mon_full} {yyyy}.zip")
    content, _ = try_urls([u])
    if content:
        zipfile.ZipFile(io.BytesIO(content)).extractall(out / "icici")
        print("  unzipped icici")
    else: print("  MISS", u)

    # ---- Aditya Birla SL (zip or xls; naming varies fortnight to
    #      fortnight, so several conventions are probed) -----------------
    print("Aditya Birla…")
    ddmmyy = f"{dd}{mm}{yy}"
    names = [
        f"absl_fortnightly_portfolio_report_{ddmmyy}.zip",   # current (since 15-Aug-2026)
        f"sebi_fortnightly_portfolio_report_{ddmmyy}.zip",
        f"sebifortnightlyportfolioreport{ddmmyy}.xls",
        f"sebifortnightlyportfolioreport{ddmmyy}.xlsx",
        f"sebi_fortnightly_portfolio_{dd}-{mon_full.lower()}-{yyyy}.zip",
        f"sebi_fortnightly_portfolio_{dd}-{mon_abbr.lower()}-{yyyy}.zip",
        f"sebi_fortnightly_portfolio-{dd}-{mon_full.lower()}-{yyyy}.zip",
        f"sebi_fortnightly_portfolio_{dd}_{mon_abbr.lower()}-{yyyy}.zip",
        f"abslmf_fortnightlydisclosure-{dd}-{mm}-{yy}.zip",
    ]
    urls = [f"https://mutualfund.adityabirlacapital.com/-/media/bsl/files/resources/"
            f"fortnightly-portfolio/{yyyy}/{n}" for n in names]
    content, matched = try_urls(urls)
    if content:
        if content[:2] == b"PK" and matched.endswith(".zip"):
            zipfile.ZipFile(io.BytesIO(content)).extractall(out / "absl")
            print("  unzipped absl:", matched.rsplit("/", 1)[-1])
        else:
            ext = ".xlsx" if content[:2] == b"PK" else ".xls"
            save(out / "absl" / f"absl{ext}", content)
    else: print("  MISS ABSL — check mutualfund.adityabirlacapital.com → "
                "Forms & Downloads → Portfolio → Fortnightly for the new filename")

    # ---- Axis (CMS API returns the exact document URL) -----------------
    print("Axis…")
    axis_content = None
    try:
        r = requests.post(
            "https://www.axismf.com/cms/get-scheme-documents",
            headers={**UA, "Content-Type": "application/json",
                     "Authorization": f"Bearer {AXIS_CMS_TOKEN}",
                     "Referer": "https://www.axismf.com/statutory-disclosures"},
            json={"sdType": "yearMonthSchemeDocs", "sdID": "sdFortnightlyPortfolio",
                  "year": yyyy, "month": mon_full, "schemeCode": "Consolidated"},
            timeout=60)
        r.raise_for_status()
        docs = (r.json().get("data") or {}).get("documentList") or []
        want = f"{dd}-{mm}-{yyyy}"
        doc = next((x for x in docs
                    if want in (x.get("documentName") or "")
                    or f"_{dd}_{mm}_{yyyy}" in (x.get("docuementURL") or "")), None)
        if doc and doc.get("docuementURL"):
            axis_content = get(doc["docuementURL"]).content
    except Exception as e:
        print("  Axis CMS API failed:", e)
    if not axis_content:
        # legacy fallback: last-known numeric path (changes over time)
        u = f"https://www.axismf.com/1/5/464/2383/3698/4524/Fortnightly_Portfolio_{dd}_{mm}_{yyyy}.xlsx"
        axis_content, _ = try_urls([u])
    if axis_content and len(axis_content) > 5000:
        save(out / "axis" / "axis.xlsx", axis_content)
    else:
        print("  MISS Axis — download manually from axismf.com > Statutory "
              "Disclosures > 8. Portfolios > Fortnightly")

    # ---- Nippon (consolidated xls; naming varies: 30-Jun-26 / 15-June-26) --
    print("Nippon…")
    variants = [f"NIMF-FORTNIGHTLY-PORTFOLIO-{d.day}-{m}-{yy}.xls"
                for m in (mon_abbr, mon_full)]
    urls = [f"https://mf.nipponindiaim.com/InvestorServices/FactsheetsDocuments/{v}"
            for v in variants]
    content, _ = try_urls(urls)
    if content: save(out / "nippon" / "nippon.xls", content)
    else: print("  MISS Nippon", urls)

    # ---- Kotak (their listing API is bot-guarded, but the S3 file host
    #      is open and the path is predictable) --------------------------
    print("Kotak…")
    kpath = (f"FAD/Portfolios/Fortnightly-Portfolio-as-on-{mon_full}-{d.day},-{yyyy}/"
             f"FortnightlyPortfolio{mon_full}{d.day}{yyyy}.xlsx")
    u = "https://vatseelabs-s3.kotakmf.com/" + requests.utils.quote(kpath)
    content, _ = try_urls([u])
    if content and content[:2] == b"PK":
        save(out / "kotak" / "kotak.xlsx", content)
    else:
        print("  MISS Kotak:", u)

    # ---- UTI (API returns CDN URL of consolidated zip; needs browser-ish
    #      headers and can 502 transiently) ------------------------------
    print("UTI…")
    uti_headers = {**UA, "Accept": "application/json",
                   "Referer": "https://www.utimf.com/forms-and-downloads/portfolio-disclosure"}
    half = "1-15" if day == 15 else f"16-{day}"
    rows = []
    for attempt in range(3):
        try:
            r = requests.get(
                "https://www.utimf.com/api/get-consolidate-debt-portfolio-disclosure",
                params={"year": yyyy, "month": f"{half} {mon_full}"},
                headers=uti_headers, timeout=60)
            r.raise_for_status()
            rows = json.loads(r.text).get("rows", [])
            break
        except Exception as e:
            print(f"  UTI API attempt {attempt + 1} failed: {e}")
            time.sleep(10 * (attempt + 1))
    if rows:
        try:
            content = get(rows[0]["url"]).content
            zipfile.ZipFile(io.BytesIO(content)).extractall(out / "uti")
            print("  unzipped uti")
        except Exception as e:
            print("  MISS UTI (download/unzip):", e)
    else:
        print("  MISS UTI: no rows for", half, mon_full)

    # ---- Tata (Drupal CMS; one consolidated workbook) -------------------
    print("Tata…")
    tname = f"Fortnightly Portfolio as on {day}{ordsfx} {mon_full} {yyyy}"
    urls = []
    for f in reversed(folders):          # try next-month folder first
        for suffix in (".xlsx", ".xls", " (1).xlsx", "_0.xlsx"):
            urls.append(f"https://betacms.tatamutualfund.com/system/files/{f}/"
                        f"{requests.utils.quote(tname + suffix)}")
    content, matched = try_urls(urls)
    if content:
        ext = ".xlsx" if content[:2] == b"PK" else ".xls"
        save(out / "tata" / f"tata{ext}", content)
    else:
        print("  MISS Tata — check tatamutualfund.com > Schemes related > "
              "Portfolio > Fortnightly")


if __name__ == "__main__":
    d = datetime.strptime(sys.argv[1], "%Y-%m-%d").date() if len(sys.argv) > 1 else date.today()
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("disclosures")
    fetch(d, out)
    print("\nNow run:  python build_mf_yields.py", out)
