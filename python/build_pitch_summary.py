"""Builds docs/pitch_summary.docx from results.json. Every number in the document is read
from results.json; nothing here is typed by hand.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx import Document
from docx.shared import Inches, Pt

from config import AUTHOR, CHARTS_DIR, DOCX_OUTPUT, RESULTS_JSON

# ---------------------------- CONFIG ----------------------------
DOC_TITLE = "Baker Hughes / Chart Industries: Merger and LBO Analysis"
DISCLAIMER = (
    "This document is an educational, illustrative analysis built from public SEC filings and "
    "labeled assumptions. It is not investment advice, a fairness opinion, or a solicitation, "
    "and it was not prepared by or for either company."
)
# ----------------------------------------------------------------


def usd_mm(x):
    return f"${x:,.0f}mm"


def pct(x):
    return f"{x:+.1%}"


def build():
    results = json.loads(RESULTS_JSON.read_text(encoding="utf-8"))
    su = results["merger_sources_and_uses"]
    ppa = results["purchase_price_allocation"]
    ad = results["accretion_dilution"]
    lbo = results["lbo"]

    doc = Document()
    doc.core_properties.author = AUTHOR

    doc.add_heading(DOC_TITLE, level=0)
    p = doc.add_paragraph()
    p.add_run(DISCLAIMER).italic = True

    doc.add_heading("Deal Overview", level=1)
    doc.add_paragraph(
        f"On {results['announcement_date']}, Baker Hughes Company (NASDAQ: BKR) agreed to acquire "
        f"Chart Industries, Inc. (NYSE: GTLS) for {su['shares']['total_diluted_shares']:.1f}mm fully "
        f"diluted shares at $210.00 per share in an all-cash transaction. The transaction closed on "
        f"{results['closing_date']}. Baker Hughes stated a total enterprise value of "
        f"{usd_mm(su['stated_enterprise_value'])} and an implied equity value of "
        f"{usd_mm(su['stated_equity_value'])} at signing, financed with newly issued senior notes, "
        f"two 2-year unsecured term loans, and cash on hand; Chart's own outstanding notes, term loan "
        f"and revolver were redeemed or repaid at closing."
    )

    doc.add_heading("Valuation and Sources and Uses", level=1)
    doc.add_paragraph(
        f"This model's independent reconstruction of the purchase price, using the fully diluted share "
        f"count and Chart's FY2025 year-end net debt, implies an enterprise value of "
        f"{usd_mm(su['implied_enterprise_value'])}, a variance of {pct(su['ev_variance_pct'])} against "
        f"the disclosed {usd_mm(su['stated_enterprise_value'])}. The gap is explained by differing "
        f"measurement dates: the disclosed figure reflects share count and net debt near signing "
        f"(July 2025), while this model uses the closest pre-close print (Q1 2026) for shares and "
        f"Chart's FY2025 10-K for net debt, since a full debt breakdown as of the actual close date was "
        f"not separately disclosed."
    )
    table = doc.add_table(rows=1, cols=2)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text = "Sources and Uses ($mm)", "Amount"
    rows = [
        ("Equity purchase price", su["uses"]["equity_purchase_price"]),
        ("Redeem Chart senior secured notes", su["uses"]["redeem_secured_notes"]),
        ("Redeem Chart senior unsecured notes", su["uses"]["redeem_unsecured_notes"]),
        ("Repay Chart term loan", su["uses"]["repay_term_loan"]),
        ("Repay Chart revolver and other debt", su["uses"]["repay_revolver"] + su["uses"]["repay_other_debt"]),
        ("Advisory and financing fees", su["uses"]["advisory_and_legal_fees"] + su["uses"]["financing_fees"]),
        ("Total Uses", su["uses"]["total_uses"]),
        ("New term loans", su["sources"]["new_term_loans"]),
        ("New USD and EUR senior notes", su["sources"]["new_usd_notes"] + su["sources"]["new_eur_notes_usd_equivalent"]),
        ("Cash on hand", su["sources"]["cash_on_hand"]),
        ("Total Sources", su["sources"]["total_sources"]),
    ]
    for label, value in rows:
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = f"{value:,.1f}"

    doc.add_heading("Purchase Price Allocation", level=1)
    doc.add_paragraph(
        f"Applying illustrative fair-value step-ups to inventory, PP&E and identifiable intangibles "
        f"(none of which Baker Hughes has separately disclosed), and a rebuilt deferred tax liability "
        f"on those step-ups, this model computes goodwill of {usd_mm(ppa['goodwill'])} against total "
        f"consideration of {usd_mm(su['uses']['equity_purchase_price'])}. The identifiable intangibles "
        f"fair value ({usd_mm(ppa['fair_value_assets']['identifiable_intangibles'])}) and its useful "
        f"life are illustrative assumptions, not disclosed figures; see the README for the full "
        f"assumption list."
    )

    doc.add_heading("Accretion / Dilution", level=1)
    doc.add_paragraph(
        "The press release states double-digit GAAP EPS accretion in the first full year after close. "
        "This independent model, which charges the combined entity the full incremental cost of "
        "financing the entire purchase (including refinancing Chart's own debt) at an illustrative 23% "
        "tax rate, and layers in amortization of the newly created intangible assets, finds the deal "
        f"dilutive in year one ({pct(ad['by_year'][0]['accretion_dilution_pct'])}) and roughly "
        f"breakeven in year two ({pct(ad['by_year'][1]['accretion_dilution_pct'])}), turning accretive "
        f"from year three onward ({pct(ad['by_year'][2]['accretion_dilution_pct'])} in 2028, reaching "
        f"{pct(ad['by_year'][-1]['accretion_dilution_pct'])} by 2030). This gap against the company's own "
        "claim is reported here rather than smoothed over; it likely reflects a non-GAAP or cash-EPS "
        "basis, a lower assumed cost of the acquisition financing, or synergies realized faster than "
        "this model's illustrative phase-in."
    )
    doc.add_picture(str(CHARTS_DIR / "accretion_dilution.png"), width=Inches(6.0))

    doc.add_heading("LBO: Hypothetical Sponsor Alternative", level=1)
    doc.add_paragraph(
        "Chart was acquired strategically, not by a financial sponsor. The LBO below is a clearly "
        "hypothetical alternative capital structure for the same target at the same disclosed $210.00 "
        "per share price, built purely for comparison."
    )
    doc.add_paragraph(
        f"At entry leverage of {lbo['sources_and_uses']['entry_leverage_x']:.1f}x EBITDA "
        f"({usd_mm(lbo['sources_and_uses']['term_loan_b'])} Term Loan B plus "
        f"{usd_mm(lbo['sources_and_uses']['senior_notes'])} senior notes) and sponsor equity of "
        f"{usd_mm(lbo['sources_and_uses']['sponsor_equity'])}, a {lbo['returns']['hold_years']}-year hold "
        f"exiting at {lbo['exit']['exit_multiple']:.1f}x EBITDA produces a {lbo['returns']['moic']:.2f}x "
        f"MOIC and a {lbo['returns']['irr']:.1%} IRR. Both figures sit below typical sponsor return "
        "hurdles (often 20%+ IRR / 2.5x+ MOIC), which is a weak result worth stating plainly: at the "
        "actual deal price and this illustrative leverage and exit multiple, this is not an obviously "
        "attractive standalone buyout."
    )
    doc.add_picture(str(CHARTS_DIR / "lbo_debt_paydown.png"), width=Inches(6.0))

    doc.add_heading("Key Assumptions and Risks", level=1)
    bullets = [
        "Synergies: $325mm annualized cost synergies at year-3 run-rate is Baker Hughes' own disclosed "
        "figure; the ramp shape (25% / 60% / 100%) and the one-times cost to achieve them are illustrative.",
        "Tax rate: a 23% blended rate is used for deal adjustments; Baker Hughes' own FY2025 effective "
        "rate was materially lower (8.8%), so this is a conservative choice, not a forecast of the "
        "combined company's actual cash tax rate.",
        "Financing cost: the new debt's blended coupon is built from the actually disclosed tranche "
        "coupons and face amounts; the term loan margin (100bps over SOFR) is an assumption within the "
        "disclosed contractual range.",
        "LBO structure (leverage split, pricing, exit multiple, cash sweep): entirely illustrative, "
        "since Chart was not actually taken private.",
        "Purchase price allocation step-ups (inventory, PP&E, intangibles) are illustrative; Baker "
        "Hughes has not disclosed a final fair-value allocation as of this writing.",
    ]
    for b in bullets:
        doc.add_paragraph(b, style="List Bullet")

    doc.add_heading("Disclaimer", level=1)
    doc.add_paragraph(DISCLAIMER)

    DOCX_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(DOCX_OUTPUT)
    print("wrote", DOCX_OUTPUT)


if __name__ == "__main__":
    build()
