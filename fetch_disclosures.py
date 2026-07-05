"""
Download the latest fortnightly debt-scheme portfolio disclosures from
8 AMCs into the folder layout expected by build_mf_yields.py:

    <out>/hdfc/*.xlsx   <out>/sbi/sbi.xlsx     <out>/icici/*.xlsx (unzipped)
    <out>/absl/*.xlsx   <out>/axis/axis.xlsx   <out>/nippon/nippon.xls
    <out>/kotak/kotak.xlsx                     <out>/uti/Sebi Exposure*.xlsx

Usage:  python fetch_disclosures.py 2026-06-30 [out_dir]
        (date = the fortnight-end being fetched: 15th or month-end)

NOTE: run this on a normal machine (it needs open internet access to the
AMC websites). Some AMCs (Axis) don't expose a stable URL — for those the
script calls the same JSON APIs their own websites use.
"""
import io
import re
import sys
import zipfile
from datetime import date, datetime
from pathlib import Path

import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}

# HDFC publishes one xlsx per debt scheme. Scheme list as of Jun-2026;
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
    "HDFC FMP 1861D March 2022", "HDFC FMP 1406D August 2022",
    "HDFC FMP 1359D September 2022", "HDFC FMP 1269D March 2023",
    "HDFC Floating Rate Debt Fund", "HDFC Dynamic Debt Fund",
    "HDFC Diversified Equity All Cap Active FOF",
    "HDFC CRISIL-IBX Financial Services 9-12 Months Debt Index Fund",
    "HDFC CRISIL-IBX Financial Services 3-6 Months Debt Index Fund",
    "HDFC Credit Risk Debt Fund", "HDFC Charity Fund for Cancer Cure",
    "HDFC Banking and PSU Debt Fund",
]


def get(url, **kw):
    r = requests.get(url, headers=UA, timeout=60, **kw)
    r.raise_for_status()
    return r


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

    # ---- HDFC (per-scheme files) -------------------------------------
    print("HDFC…")
    # files are uploaded early the following month, so folder = publish month
    folders = [d.strftime("%Y-%m")]
    nxt = date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    folders.append(nxt.strftime("%Y-%m"))
    for scheme in HDFC_SCHEMES:
        fname = f"{scheme} - {dd}-{mon_full}-{yyyy}.xlsx"
        urls = [f"https://files.hdfcfund.com/s3fs-public/{f}/{requests.utils.quote(fname)}"
                for f in reversed(folders)]
        content, _ = try_urls(urls)
        if content:
            save(out / "hdfc" / fname, content)
        else:
            print(f"  MISS {fname}")

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

    # ---- Aditya Birla SL (zip with one consolidated xlsx) --------------
    print("Aditya Birla…")
    u = (f"https://mutualfund.adityabirlacapital.com/-/media/bsl/files/resources/"
         f"fortnightly-portfolio/{yyyy}/sebi_fortnightly_portfolio_{dd}-{mon_full.lower()}-{yyyy}.zip")
    content, _ = try_urls([u])
    if content:
        zipfile.ZipFile(io.BytesIO(content)).extractall(out / "absl")
        print("  unzipped absl")
    else: print("  MISS", u)

    # ---- Axis (consolidated xlsx; numeric URL discovered via fiber/DOM;
    #      no stable pattern — try the JSON-less fallback of last known path
    #      then give up with a message) -----------------------------------
    print("Axis…")
    u = f"https://www.axismf.com/1/5/464/2383/3698/4349/Fortnightly_Portfolio_{dd}_{mm}_{yyyy}.xlsx"
    content, _ = try_urls([u])
    if content: save(out / "axis" / "axis.xlsx", content)
    else: print("  MISS Axis — download manually from axismf.com → Statutory "
                "Disclosures → 8. Portfolios → Fortnightly")

    # ---- Nippon (consolidated xls; naming varies: 30-Jun-26 / 15-June-26) --
    print("Nippon…")
    variants = [f"NIMF-FORTNIGHTLY-PORTFOLIO-{d.day}-{m}-{yy}.xls"
                for m in (mon_abbr, mon_full)]
    urls = [f"https://mf.nipponindiaim.com/InvestorServices/FactsheetsDocuments/{v}"
            for v in variants]
    content, _ = try_urls(urls)
    if content: save(out / "nippon" / "nippon.xls", content)
    else: print("  MISS Nippon", urls)

    # ---- Kotak (API lists content path on a public S3 host) -------------
    print("Kotak…")
    try:
        js = get("https://www.kotakmf.com/api/kotakapi/forms/user/v1/getsubheaderList/417").json()
        want = f"Fortnightly Portfolio as on {mon_full} {d.day}, {yyyy}"
        item = next((i for i in js.get("subHeaderList", [])
                     if i.get("subHeaderTitle", "").strip() == want), None)
        if item:
            u = "https://vatseelabs-s3.kotakmf.com/" + requests.utils.quote(item["content"])
            save(out / "kotak" / "kotak.xlsx", get(u).content)
        else:
            print("  MISS Kotak:", want)
    except Exception as e:
        print("  Kotak API error:", e)

    # ---- UTI (API returns CDN URL of consolidated zip) ------------------
    print("UTI…")
    try:
        half = "1-15" if day == 15 else f"16-{day}"
        js = get("https://www.utimf.com/api/get-consolidate-debt-portfolio-disclosure",
                 params={"year": yyyy, "month": f"{half} {mon_full}"}).json()
        rows = js.get("rows", [])
        if rows:
            u = rows[0]["url"]
            content = get(u).content
            zipfile.ZipFile(io.BytesIO(content)).extractall(out / "uti")
            print("  unzipped uti")
        else:
            print("  MISS UTI: no rows for", half, mon_full)
    except Exception as e:
        print("  UTI API error:", e)


if __name__ == "__main__":
    d = datetime.strptime(sys.argv[1], "%Y-%m-%d").date() if len(sys.argv) > 1 else date.today()
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("disclosures")
    fetch(d, out)
    print("\nNow run:  python build_mf_yields.py", out)
