"""Download SEC filings and FRED data, and snapshot the excerpts the model uses into data/."""
import json
import re
import time
import warnings
from datetime import date
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# ---------------------------- CONFIG ----------------------------
USER_AGENT = "portfolio-research contact@example.com"
REPO = Path(__file__).resolve().parent.parent
DATA_DIR = REPO / "data"
EXCERPT_DIR = DATA_DIR / "excerpts"
CACHE_DIR = REPO / ".cache"
SLEEP_SECONDS = 0.3
FRED_SERIES = ["SOFR"]
FRED_START = "2026-06-01"
CHART_CIK = "0000892553"
BKR_CIK = "0001701605"
# name -> (cik, accession, primary document)
FILINGS = {
    "bkr_8k_20250729": ("1701605", "0001193125-25-167118", "d912206d8k.htm"),
    "bkr_ex991_20250729": ("1701605", "0001193125-25-167118", "d912206dex991.htm"),
    "chart_8k_20250729": ("892553", "0001193125-25-167126", "d89051d8k.htm"),
    "chart_defm14a": ("892553", "0001193125-25-198320", "d52397ddefm14a.htm"),
    "chart_10k_fy2025": ("892553", "0000892553-26-000027", "gtls-20251231.htm"),
    "chart_10q_q1_2026": ("892553", "0000892553-26-000040", "gtls-20260331.htm"),
    "chart_10q_q2_2025": ("892553", "0000892553-25-000119", "gtls-20250630.htm"),
    "chart_ex991_20260227": ("892553", "0000892553-26-000025", "exhibit991gtls-20251231pre.htm"),
    "chart_8k_20260716": ("892553", "0001193125-26-305482", "d10289d8k.htm"),
    "bkr_10k_fy2025": ("1701605", "0001701605-26-000007", "bkr-20251231.htm"),
    "bkr_10q_q1_2026": ("1701605", "0001701605-26-000014", "bkr-20260331.htm"),
    "bkr_10q_q2_2026": ("1701605", "0001701605-26-000023", "bkr-20260630.htm"),
    "bkr_8k_20260716": ("1701605", "0001193125-26-305477", "d105425d8k.htm"),
    "bkr_ex991_20260716": ("1701605", "0001193125-26-305477", "d105425dex991.htm"),
}
WHOLE_DOCS = ["bkr_ex991_20250729", "bkr_8k_20260716", "chart_8k_20260716", "bkr_ex991_20260716"]
# excerpt name -> (doc, start regex, number of lines, optional skip of first N regex matches)
WINDOWS = {
    "chart_10k_balance_sheet": ("chart_10k_fy2025", r"^CONSOLIDATED BALANCE SHEETS", 45, 0),
    "chart_10k_income_statement": ("chart_10k_fy2025", r"^CONSOLIDATED STATEMENTS OF INCOME", 55, 0),
    "chart_10k_cash_flow": ("chart_10k_fy2025", r"^CONSOLIDATED STATEMENTS OF CASH FLOWS", 60, 0),
    "chart_10k_debt_note": ("chart_10k_fy2025", r"^Debt and Credit Arrangements", 40, 0),
    "chart_10k_options": ("chart_10k_fy2025", r"summarizes our stock option activity", 14, 0),
    "chart_10k_rsu_psu": ("chart_10k_fy2025", r"summarizes our unvested restricted stock and RSUs", 30, 0),
    "chart_10k_preferred_conversion": ("chart_10k_fy2025", r"all outstanding shares of the Mandatory Convertible Preferred", 3, 0),
    "chart_ebitda_reconciliation_fy2025": ("chart_ex991_20260227", r"^RECONCILIATION OF NET INCOME FROM CONTINUING OPERATIONS TO EBITDA", 30, 0),
    "chart_10q_q1_2026_balance_sheet": ("chart_10q_q1_2026", r"^CONDENSED CONSOLIDATED BALANCE SHEETS", 45, 0),
    "chart_10q_q1_2026_income_statement": ("chart_10q_q1_2026", r"^CONDENSED CONSOLIDATED STATEMENTS OF OPERATIONS", 45, 0),
    "chart_10q_q2_2025_balance_sheet": ("chart_10q_q2_2025", r"^Cash and cash equivalents \|", 36, 0),
    "chart_10q_q2_2025_debt_note": ("chart_10q_q2_2025", r"following table represents the components of our borrowings", 22, 0),
    "chart_defm14a_projections": ("chart_defm14a", r"The following table summarizes the Chart Projections", 14, 0),
    "chart_defm14a_market_price": ("chart_defm14a", r"the last trading day before the public announcement of Chart", 4, 0),
    "chart_defm14a_background_value": ("chart_defm14a", r"The Baker Hughes Proposal implied an equity value", 3, 0),
    "chart_defm14a_wells_fargo": ("chart_defm14a", r"Taking$", 6, 0),
    "bkr_10k_income_statement": ("bkr_10k_fy2025", r"^CONSOLIDATED STATEMENTS OF INCOME", 28, 0),
    "bkr_10k_balance_sheet": ("bkr_10k_fy2025", r"^CONSOLIDATED STATEMENTS OF FINANCIAL POSITION", 45, 0),
    "bkr_10k_cash_flow": ("bkr_10k_fy2025", r"^CONSOLIDATED STATEMENTS OF CASH FLOWS", 45, 0),
    "bkr_10k_eps_note": ("bkr_10k_fy2025", r"^EARNINGS PER SHARE", 14, 0),
    "bkr_10k_segment_ebitda": ("bkr_10k_fy2025", r"^Add: Depreciation and amortization \| 932", 40, 0),
    "bkr_10k_interest_income": ("bkr_10k_fy2025", r"Net interest expense incurred in 2025", 2, 0),
    "bkr_10k_chart_advance": ("bkr_10k_fy2025", r"This payment on behalf of Chart was recorded", 1, 0),
    "bkr_10q_q1_2026_debt_note": ("bkr_10q_q1_2026", r"^DEBT$", 32, 0),
    "bkr_10q_q2_2026_debt_and_chart": ("bkr_10q_q2_2026", r"On March 11, 2026, the Company completed an offering", 12, 0),
    "bkr_10q_q2_2026_chart_acquisition": ("bkr_10q_q2_2026", r"the Company completed the previously announced acquisition of Chart", 6, 0),
}
# ----------------------------------------------------------------


def get(url, tries=4):
    headers = {"User-Agent": USER_AGENT}
    for attempt in range(tries):
        response = requests.get(url, headers=headers, timeout=60)
        if response.status_code == 200:
            time.sleep(SLEEP_SECONDS)
            return response
        time.sleep(2 * (attempt + 1))
    response.raise_for_status()


def to_text(html_bytes):
    warnings.filterwarnings("ignore")
    soup = BeautifulSoup(html_bytes, "lxml")
    for table in soup.find_all("table"):
        rows = []
        for tr in table.find_all("tr"):
            cells = [re.sub(r"\s+", " ", c.get_text(" ")).strip() for c in tr.find_all(["td", "th"])]
            cells = [c for c in cells if c and c not in ("$", ")", "%")]
            if cells:
                rows.append(" | ".join(cells))
        table.replace_with(soup.new_string("\n[TABLE]\n" + "\n".join(rows) + "\n[/TABLE]\n"))
    return re.sub(r"\n\s*\n+", "\n", soup.get_text("\n"))


def filing_url(cik, accession, document):
    return f"https://www.sec.gov/Archives/edgar/data/{cik}/{accession.replace('-', '')}/{document}"


def load_docs():
    CACHE_DIR.mkdir(exist_ok=True)
    texts = {}
    for name, (cik, accession, document) in FILINGS.items():
        cache_file = CACHE_DIR / f"{name}.txt"
        if not cache_file.exists():
            cache_file.write_text(to_text(get(filing_url(cik, accession, document)).content), encoding="utf-8")
        texts[name] = cache_file.read_text(encoding="utf-8").splitlines()
    return texts


def window(lines, pattern, count, skip):
    matches = [i for i, line in enumerate(lines) if re.search(pattern, line)]
    if len(matches) <= skip:
        raise ValueError(f"pattern not found: {pattern}")
    start = matches[skip]
    keep = [ln for ln in lines[start:start + count] if ln.strip() not in ("[TABLE]", "[/TABLE]")]
    return "\n".join(keep)


def snapshot_excerpts(texts):
    EXCERPT_DIR.mkdir(parents=True, exist_ok=True)
    for name in WHOLE_DOCS:
        (EXCERPT_DIR / f"{name}.txt").write_text("\n".join(texts[name]), encoding="utf-8")
    for name, (doc, pattern, count, skip) in WINDOWS.items():
        (EXCERPT_DIR / f"{name}.txt").write_text(window(texts[doc], pattern, count, skip), encoding="utf-8")


def snapshot_fred():
    for series in FRED_SERIES:
        text = get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series}").text
        lines = text.strip().splitlines()
        keep = [lines[0]] + [ln for ln in lines[1:] if ln.split(",")[0] >= FRED_START]
        (DATA_DIR / f"fred_{series.lower()}.csv").write_text("\n".join(keep) + "\n", encoding="utf-8")


def snapshot_companyfacts():
    # Kept small: only the tags used in the cross-check
    tags = ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "NetIncomeLoss", "CashAndCashEquivalentsAtCarryingValue",
            "LongTermDebtNoncurrent", "Goodwill", "StockholdersEquity", "MinorityInterest"]
    out = {}
    for label, cik in (("chart", CHART_CIK), ("bkr", BKR_CIK)):
        facts = get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json").json()["facts"]["us-gaap"]
        out[label] = {}
        for tag in tags:
            if tag in facts:
                rows = facts[tag]["units"].get("USD", [])
                out[label][tag] = [r for r in rows if r.get("form") in ("10-K", "10-Q") and r["end"] >= "2025-03-01"]
    (DATA_DIR / "xbrl_companyfacts_subset.json").write_text(json.dumps(out, indent=1), encoding="utf-8")


if __name__ == "__main__":
    DATA_DIR.mkdir(exist_ok=True)
    docs = load_docs()
    snapshot_excerpts(docs)
    snapshot_fred()
    snapshot_companyfacts()
    print("snapshot complete", date.today().isoformat())
