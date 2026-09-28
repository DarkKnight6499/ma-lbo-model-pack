"""Builds excel/ma_lbo_model.xlsx as a formula-driven workbook (openpyxl writes formulas,
not values). Mirrors the independent Python model in model.py so the two can be diffed.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

from config import AUTHOR, DEAL_NAME, EXCEL_OUTPUT, FORECAST_YEARS, INPUTS

# ---------------------------- CONFIG ----------------------------
HEADER_FONT = Font(bold=True, size=12)
SUBHEADER_FONT = Font(bold=True, size=10)
BOLD = Font(bold=True)
LABEL_COL_WIDTH = 46
NUMBER_COL_WIDTH = 15
# ----------------------------------------------------------------


def new_sheet(wb, title):
    ws = wb.create_sheet(title)
    ws.column_dimensions["A"].width = LABEL_COL_WIDTH
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = NUMBER_COL_WIDTH
    return ws


def header(ws, text, row=1, col=1):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = HEADER_FONT
    return row + 2


def row_label(ws, row, label, bold=False):
    cell = ws.cell(row=row, column=1, value=label)
    if bold:
        cell.font = BOLD
    return cell


NUM_FMT = '#,##0.0'
PCT_FMT = '0.0%'
X_FMT = '0.00"x"'
USD_FMT = '$#,##0.0'


def build_inputs_sheet(wb):
    """Values live in columns C:G (C alone for a scalar, C:G for a 5-year list); metadata
    always sits at fixed columns I:L so a 5-value list can never collide with it."""
    ws = new_sheet(wb, "Inputs")
    r = header(ws, f"{DEAL_NAME}: Model Inputs (all assumptions and disclosed figures, cited)")
    ws.cell(row=r, column=1, value="Key").font = SUBHEADER_FONT
    ws.cell(row=r, column=2, value="Label").font = SUBHEADER_FONT
    ws.cell(row=r, column=3, value="Value(s)").font = SUBHEADER_FONT
    ws.cell(row=r, column=9, value="Unit").font = SUBHEADER_FONT
    ws.cell(row=r, column=10, value="Kind").font = SUBHEADER_FONT
    ws.cell(row=r, column=11, value="Source").font = SUBHEADER_FONT
    ws.cell(row=r, column=12, value="Locator").font = SUBHEADER_FONT
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["L"].width = 70
    r += 1
    cell_refs = {}
    for key in INPUTS.keys():
        row = INPUTS.row(key)
        raw = row["value"]
        ws.cell(row=r, column=1, value=key)
        ws.cell(row=r, column=2, value=row["label"])
        if "|" in raw:
            values = raw.split("|")
            for i, v in enumerate(values):
                ws.cell(row=r, column=3 + i, value=float(v))
            cell_refs[key] = [f"Inputs!${get_column_letter(3 + i)}${r}" for i in range(len(values))]
        else:
            try:
                ws.cell(row=r, column=3, value=float(raw))
            except ValueError:
                ws.cell(row=r, column=3, value=raw)
            cell_refs[key] = f"Inputs!$C${r}"
        ws.cell(row=r, column=9, value=row["unit"])
        ws.cell(row=r, column=10, value=row["kind"])
        ws.cell(row=r, column=11, value=row["source_id"])
        ws.cell(row=r, column=12, value=row["locator"])
        r += 1
    return cell_refs


def build_sources_uses_sheet(wb, ref):
    ws = new_sheet(wb, "Sources_Uses")
    r = header(ws, "Merger Sources and Uses (Baker Hughes acquisition of Chart Industries, actual terms)")

    row_label(ws, r, "Diluted Chart shares (basic + RSU + PSU + option treasury method)", bold=True)
    r += 1
    row_label(ws, r, "Basic shares outstanding (closest pre-close print)")
    ws.cell(row=r, column=2, value=47.869076)
    basic_row = r
    r += 1
    row_label(ws, r, "Unvested RSUs")
    ws.cell(row=r, column=2, value=f"={ref['chart_rsu']}")
    rsu_row = r
    r += 1
    row_label(ws, r, "Unvested PSUs (100% earned)")
    ws.cell(row=r, column=2, value=f"={ref['chart_psu']}")
    psu_row = r
    r += 1
    row_label(ws, r, "Options: treasury-method dilution")
    ws.cell(row=r, column=2, value=f"={ref['chart_options']}*(1-{ref['chart_option_strike']}/{ref['offer_price']})")
    opt_row = r
    r += 1
    row_label(ws, r, "Total diluted shares (mm)", bold=True)
    ws.cell(row=r, column=2, value=f"=B{basic_row}+B{rsu_row}+B{psu_row}+B{opt_row}").font = BOLD
    shares_row = r
    r += 2

    row_label(ws, r, "USES", bold=True)
    r += 1
    uses_start = r
    row_label(ws, r, "Equity purchase price")
    ws.cell(row=r, column=2, value=f"=B{shares_row}*{ref['offer_price']}")
    equity_row = r
    r += 1
    row_label(ws, r, "Redeem Chart 7.500% senior secured notes due 2030 (at premium)")
    ws.cell(row=r, column=2, value=f"={ref['chart_notes_secured_principal']}*(1+{ref['redemption_premium_secured_pct']})")
    r += 1
    row_label(ws, r, "Redeem Chart 9.500% senior notes due 2031 (at premium)")
    ws.cell(row=r, column=2, value=f"={ref['chart_notes_unsecured_principal']}*(1+{ref['redemption_premium_unsecured_pct']})")
    r += 1
    row_label(ws, r, "Repay Chart term loan (at par)")
    ws.cell(row=r, column=2, value=f"={ref['chart_term_loan_principal']}")
    r += 1
    row_label(ws, r, "Repay Chart revolver (at par)")
    ws.cell(row=r, column=2, value=f"={ref['chart_revolver_principal']}")
    r += 1
    row_label(ws, r, "Repay other Chart debt facilities")
    ws.cell(row=r, column=2, value=f"={ref['chart_other_debt_principal']}")
    r += 1
    row_label(ws, r, "Advisory and legal fees (0.75% of stated EV, illustrative)")
    ws.cell(row=r, column=2, value=f"={ref['advisory_fee_pct']}*{ref['stated_ev']}")
    r += 1
    row_label(ws, r, "New-debt financing fees (0.25% of gross new debt, illustrative)")
    new_debt_formula_cell = r  # placeholder, filled after sources block known; recomputed below
    r += 1
    uses_end_before_total = r - 1
    row_label(ws, r, "Total Uses", bold=True)
    total_uses_row = r
    r += 2

    row_label(ws, r, "SOURCES", bold=True)
    r += 1
    sources_start = r
    row_label(ws, r, "New term loans (Bank of America + UniCredit, $1.0bn each)")
    ws.cell(row=r, column=2, value=f"={ref['term_loan_total']}")
    tl_row = r
    r += 1
    row_label(ws, r, "New USD senior notes (March 2026 offering)")
    ws.cell(row=r, column=2, value=f"={ref['usd_notes_face_total']}")
    usd_row = r
    r += 1
    row_label(ws, r, "New EUR senior notes, USD equivalent (carrying value proxy)")
    eur_formula = "+".join(f"{c}" for c in ref["eur_tranche_carrying"])
    ws.cell(row=r, column=2, value=f"={eur_formula}")
    eur_row = r
    r += 1
    row_label(ws, r, "Cash on hand (plug)")
    cash_row = r
    r += 1
    row_label(ws, r, "Total Sources", bold=True)
    total_sources_row = r
    ws.cell(row=r, column=2, value=f"=B{tl_row}+B{usd_row}+B{eur_row}+B{cash_row}").font = BOLD

    # now fill in financing fee (0.25% of gross new debt) and cash plug
    ws.cell(row=new_debt_formula_cell, column=2, value=f"={ref['tl_fee_pct']}*(B{tl_row}+B{usd_row}+B{eur_row})")
    ws.cell(row=total_uses_row, column=2, value=f"=SUM(B{uses_start}:B{uses_end_before_total})").font = BOLD
    ws.cell(row=cash_row, column=2, value=f"=B{total_uses_row}-(B{tl_row}+B{usd_row}+B{eur_row})")

    r += 2
    row_label(ws, r, "Implied enterprise value (equity purchase price + Chart FY2025 net debt)")
    ev_row = r
    net_debt_formula = (
        f"({ref['chart_notes_secured_principal']}+{ref['chart_notes_unsecured_principal']}+"
        f"{ref['chart_term_loan_principal']}+{ref['chart_revolver_principal']}+{ref['chart_other_debt_principal']}"
        f"-{ref['chart_cash']})"
    )
    ws.cell(row=r, column=2, value=f"=B{equity_row}+{net_debt_formula}")
    r += 1
    row_label(ws, r, "Stated enterprise value (press release)")
    ws.cell(row=r, column=2, value=f"={ref['stated_ev']}")
    stated_ev_row = r
    r += 1
    row_label(ws, r, "Variance vs. stated EV")
    ws.cell(row=r, column=2, value=f"=B{ev_row}/B{stated_ev_row}-1").number_format = PCT_FMT

    for row in range(uses_start, total_sources_row + 1):
        c = ws.cell(row=row, column=2)
        if isinstance(c.value, str) and c.value.startswith("="):
            c.number_format = USD_FMT
    ws.cell(row=ev_row, column=2).number_format = USD_FMT
    ws.cell(row=stated_ev_row, column=2).number_format = USD_FMT

    refs = {
        "equity_purchase_price": f"Sources_Uses!$B${equity_row}",
        "total_uses": f"Sources_Uses!$B${total_uses_row}",
        "total_sources": f"Sources_Uses!$B${total_sources_row}",
        "financing_fees": f"Sources_Uses!$B${new_debt_formula_cell}",
        "diluted_shares": f"Sources_Uses!$B${shares_row}",
        "implied_ev": f"Sources_Uses!$B${ev_row}",
        "stated_ev": f"Sources_Uses!$B${stated_ev_row}",
    }
    return refs


def build_ppa_sheet(wb, ref, su_ref):
    ws = new_sheet(wb, "Purchase_Price_Allocation")
    r = header(ws, "Purchase Price Allocation (Chart Industries, at close)")

    row_label(ws, r, "Fair value of identifiable assets acquired", bold=True)
    r += 1
    items = [
        ("Cash and equivalents", f"={ref['chart_cash']}"),
        ("Accounts receivable, net", f"={ref['chart_ar']}"),
        ("Inventory (book x (1 + step-up %))", f"={ref['chart_inventory']}*(1+{ref['inventory_stepup_pct']})"),
        ("Unbilled contract revenue", f"={ref['chart_unbilled']}"),
        ("Other current assets", f"={ref['chart_other_current_assets']}"),
        ("PP&E (book x (1 + step-up %))", f"={ref['chart_ppe']}*(1+{ref['ppe_stepup_pct']})"),
        ("Identifiable intangible assets (fair value)", f"={ref['new_intangibles']}"),
        ("Other assets", f"={ref['chart_other_assets']}"),
    ]
    asset_rows = {}
    for label, formula in items:
        row_label(ws, r, label)
        ws.cell(row=r, column=2, value=formula).number_format = USD_FMT
        asset_rows[label] = r
        r += 1
    total_assets_row = r
    row_label(ws, r, "Total fair value of assets", bold=True)
    ws.cell(row=r, column=2, value=f"=SUM(B{min(asset_rows.values())}:B{max(asset_rows.values())})").number_format = USD_FMT
    r += 2

    ppe_stepup_row = asset_rows["PP&E (book x (1 + step-up %))"]
    inv_stepup_row = asset_rows["Inventory (book x (1 + step-up %))"]
    intang_row = asset_rows["Identifiable intangible assets (fair value)"]

    row_label(ws, r, "New deferred tax liability from fair-value step-ups (no tax basis step-up assumed)", bold=True)
    r += 1
    row_label(ws, r, "  = tax rate x (intangibles FV - old tax basis + PP&E step-up + inventory step-up)")
    dtl_formula = (
        f"={ref['tax_rate']}*(B{intang_row}-{ref['chart_intangibles']}"
        f"+(B{ppe_stepup_row}-{ref['chart_ppe']})+(B{inv_stepup_row}-{ref['chart_inventory']}))"
    )
    ws.cell(row=r, column=2, value=dtl_formula).number_format = USD_FMT
    new_dtl_row = r
    r += 2

    row_label(ws, r, "Fair value of liabilities assumed", bold=True)
    r += 1
    liab_items = [
        ("Accounts payable", f"={ref['chart_ap']}"),
        ("Customer advances and billings in excess", f"={ref['chart_advances']}"),
        ("Accrued interest", f"={ref['chart_accrued_interest']}"),
        ("Other current liabilities", f"={ref['chart_other_current_liab']}"),
        ("New deferred tax liability", f"=B{new_dtl_row}"),
        ("Other long-term liabilities", f"={ref['chart_other_lt_liab']}"),
    ]
    liab_rows = {}
    for label, formula in liab_items:
        row_label(ws, r, label)
        ws.cell(row=r, column=2, value=formula).number_format = USD_FMT
        liab_rows[label] = r
        r += 1
    total_liab_row = r
    row_label(ws, r, "Total fair value of liabilities", bold=True)
    ws.cell(row=r, column=2, value=f"=SUM(B{min(liab_rows.values())}:B{max(liab_rows.values())})").number_format = USD_FMT
    r += 2

    row_label(ws, r, "Net identifiable assets acquired", bold=True)
    net_row = r
    ws.cell(row=r, column=2, value=f"=B{total_assets_row}-B{total_liab_row}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Less: noncontrolling interests (measured at carrying value, illustrative simplification)")
    nci_row = r
    ws.cell(row=r, column=2, value=f"={ref['chart_nci']}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Equity purchase price (consideration transferred)")
    consid_row = r
    ws.cell(row=r, column=2, value=f"={su_ref['equity_purchase_price']}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Goodwill", bold=True)
    goodwill_row = r
    ws.cell(row=r, column=2, value=f"=B{consid_row}-(B{net_row}-B{nci_row})").font = BOLD
    ws.cell(row=r, column=2).number_format = USD_FMT

    refs = {
        "goodwill": f"Purchase_Price_Allocation!$B${goodwill_row}",
        "intangibles_fv": f"Purchase_Price_Allocation!$B${intang_row}",
        "ppe_fv": f"Purchase_Price_Allocation!$B${ppe_stepup_row}",
        "inventory_fv": f"Purchase_Price_Allocation!$B${inv_stepup_row}",
    }
    return refs


def build_proforma_is_sheet(wb, ref, su_ref, ppa_ref):
    ws = new_sheet(wb, "ProForma_IS")
    r = header(ws, "Combined Pro Forma Income Statement (accretion/dilution bridge, GAAP basis)")
    ws.cell(row=r - 1, column=2, value="Formula-driven bridge: standalone BKR NI + Chart's unlevered after-tax NI")

    for i, year in enumerate(FORECAST_YEARS):
        ws.cell(row=r, column=2 + i, value=year).font = SUBHEADER_FONT
    r += 1

    def sales_formula(i):
        if ",".join(map(str, [1])):
            pass
        return None

    # Chart standalone sales/EBITDA (base case growth x margin, scenario switch handled via IF)
    row_label(ws, r, "Chart standalone sales")
    sales_row = r
    prior_ref = ref["chart_sales_fy2025"]
    for i in range(5):
        growth_ref = ref["base_growth"][i]
        mgmt_ref = ref["mgmt_sales"][i]
        col = get_column_letter(2 + i)
        prior_cell = f"{get_column_letter(1 + i)}{sales_row}" if i > 0 else None
        base_formula = f"{prior_cell}*(1+{growth_ref})" if prior_cell else f"{prior_ref}*(1+{growth_ref})"
        formula = f"=IF({ref['scenario']}=2,{mgmt_ref},{base_formula})"
        ws.cell(row=r, column=2 + i, value=formula).number_format = USD_FMT
    r += 1

    row_label(ws, r, "Chart standalone EBITDA")
    ebitda_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        margin_ref = ref["base_margin"][i]
        mgmt_ref = ref["mgmt_ebitda"][i]
        sales_cell = f"{col}{sales_row}"
        formula = f"=IF({ref['scenario']}=2,{mgmt_ref},{sales_cell}*{margin_ref})"
        ws.cell(row=r, column=2 + i, value=formula).number_format = USD_FMT
    r += 1

    row_label(ws, r, "Less: Chart standalone D&A (held flat at FY2025 level)")
    da_row = r
    for i in range(5):
        ws.cell(row=r, column=2 + i, value=f"={ref['chart_da_fy2025']}").number_format = USD_FMT
    r += 1

    row_label(ws, r, "Chart standalone EBIT")
    ebit_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        ws.cell(row=r, column=2 + i, value=f"={col}{ebitda_row}-{col}{da_row}").number_format = USD_FMT
    r += 1

    row_label(ws, r, "Chart unlevered after-tax NI ((1 - tax rate) x EBIT)")
    unlev_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        ws.cell(row=r, column=2 + i, value=f"=(1-{ref['tax_rate']})*{col}{ebit_row}").number_format = USD_FMT
    r += 2

    row_label(ws, r, "Realized run-rate synergies")
    syn_row = r
    for i in range(5):
        phase_ref = ref["synergy_phase"][i]
        ws.cell(row=r, column=2 + i, value=f"={ref['stated_synergies_runrate']}*{phase_ref}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Less: integration costs")
    integ_row = r
    for i in range(5):
        phase_ref = ref["integration_cost_phase"][i]
        ws.cell(row=r, column=2 + i,
                value=f"={ref['stated_synergies_runrate']}*{ref['integration_cost_multiple']}*{phase_ref}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Less: one-time inventory step-up charge (year 1 only)")
    invchg_row = r
    for i in range(5):
        if i == 0:
            ws.cell(row=r, column=2 + i, value=f"={ppa_ref['inventory_fv']}-{ref['chart_inventory']}").number_format = USD_FMT
        else:
            ws.cell(row=r, column=2 + i, value=0).number_format = USD_FMT
    r += 1
    row_label(ws, r, "Net synergy impact, after tax")
    syn_net_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        ws.cell(row=r, column=2 + i,
                value=f"=(1-{ref['tax_rate']})*({col}{syn_row}-{col}{integ_row}-{col}{invchg_row})").number_format = USD_FMT
    r += 2

    row_label(ws, r, "Incremental D&A from step-up, after tax")
    incda_row = r
    for i in range(5):
        ws.cell(row=r, column=2 + i,
                value=f"=(1-{ref['tax_rate']})*(({ppa_ref['intangibles_fv']}/{ref['intangibles_life']})"
                      f"+(({ppa_ref['ppe_fv']}-{ref['chart_ppe']})/{ref['ppe_stepup_life']}))").number_format = USD_FMT
    r += 1
    row_label(ws, r, "New financing interest expense, after tax")
    intexp_row = r
    weighted_coupon = build_weighted_coupon_formula(ref)
    new_debt_total_formula = f"({ref['term_loan_total']}+{ref['usd_notes_face_total']}+" + "+".join(ref["eur_tranche_carrying"]) + ")"
    for i in range(5):
        ws.cell(row=r, column=2 + i,
                value=f"=(1-{ref['tax_rate']})*{new_debt_total_formula}*({weighted_coupon})/{new_debt_total_formula}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Financing fee amortization, after tax")
    feeamort_row = r
    for i in range(5):
        ws.cell(row=r, column=2 + i,
                value=f"=(1-{ref['tax_rate']})*{su_ref['financing_fees']}/{ref['fin_fee_amort_years']}").number_format = USD_FMT
    r += 2

    row_label(ws, r, "Baker Hughes standalone net income")
    bkrni_row = r
    for i in range(5):
        col_prior = get_column_letter(1 + i) if i > 0 else None
        prior = f"{col_prior}{bkrni_row}" if col_prior else f"{ref['bkr_ni']}"
        ws.cell(row=r, column=2 + i, value=f"={prior}*(1+{ref['bkr_ni_growth']})").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Less: noncontrolling interest income")
    nci_row = r
    for i in range(5):
        ws.cell(row=r, column=2 + i, value=f"={ref['nci_income']}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Pro forma net income", bold=True)
    pfni_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        formula = (
            f"={col}{bkrni_row}+{col}{unlev_row}+{col}{syn_net_row}"
            f"-{col}{incda_row}-{col}{intexp_row}-{col}{feeamort_row}-{col}{nci_row}"
        )
        cell = ws.cell(row=r, column=2 + i, value=formula)
        cell.font = BOLD
        cell.number_format = USD_FMT

    refs = {
        "pro_forma_ni_row": pfni_row,
        "bkr_ni_row": bkrni_row,
    }
    return refs


def build_weighted_coupon_formula(ref):
    usd_terms = "+".join(f"{f}*{c}" for f, c in zip(ref["usd_tranche_face"], ref["usd_tranche_coupons"]))
    eur_terms = "+".join(f"{e}*{c}" for e, c in zip(ref["eur_tranche_carrying"], ref["eur_tranche_coupons"]))
    tl_term = f"{ref['term_loan_total']}*({ref['sofr_at_close']}+{ref['term_loan_margin']})"
    return f"{usd_terms}+{eur_terms}+{tl_term}"


def build_accretion_dilution_sheet(wb, ref, pf_ref):
    ws = new_sheet(wb, "Accretion_Dilution")
    r = header(ws, "GAAP EPS Accretion / (Dilution) by Year")
    for i, year in enumerate(FORECAST_YEARS):
        ws.cell(row=r, column=2 + i, value=year).font = SUBHEADER_FONT
    r += 1

    row_label(ws, r, "Standalone BKR diluted EPS")
    sa_eps_row = r
    for i in range(5):
        col_prior = get_column_letter(1 + i) if i > 0 else None
        prior = f"{col_prior}{sa_eps_row}" if col_prior else f"{ref['bkr_diluted_eps']}"
        ws.cell(row=r, column=2 + i, value=f"={prior}*(1+{ref['bkr_ni_growth']})").number_format = '$0.00'
    r += 1

    row_label(ws, r, "Pro forma diluted EPS (all-cash deal, no new BKR shares)")
    pf_eps_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        ws.cell(row=r, column=2 + i,
                value=f"=ProForma_IS!{col}{pf_ref['pro_forma_ni_row']}/{ref['bkr_diluted_shares']}").number_format = '$0.00'
    r += 1

    row_label(ws, r, "Accretion / (dilution) %", bold=True)
    acc_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        cell = ws.cell(row=r, column=2 + i, value=f"={col}{pf_eps_row}/{col}{sa_eps_row}-1")
        cell.font = BOLD
        cell.number_format = PCT_FMT
    r += 2

    row_label(ws, r, "Breakeven run-rate synergies (year 3, if dilutive)")
    col3 = get_column_letter(2 + 2)
    formula = (
        f"=IF({col3}{acc_row}<0,ProForma_IS!{col3}{pf_ref['bkr_ni_row']}*0"
        f"+((({col3}{sa_eps_row}-{col3}{pf_eps_row})*{ref['bkr_diluted_shares']})/(1-{ref['tax_rate']}))"
        f"+{ref['stated_synergies_runrate']}*{ref['synergy_phase'][2]},"
        f"{ref['stated_synergies_runrate']}*{ref['synergy_phase'][2]})"
    )
    ws.cell(row=r, column=2, value=formula).number_format = USD_FMT
    return {}


def build_lbo_sheet(wb, ref):
    ws = new_sheet(wb, "LBO")
    r = header(ws, "Standalone LBO of Chart Industries (hypothetical sponsor bid at the disclosed $210/share price)")
    row_label(ws, r, "NOTE: hypothetical structure. Chart was acquired strategically by Baker Hughes; this sponsor", bold=True)
    r += 1
    row_label(ws, r, "capital structure and return profile are illustrative only, for comparison purposes.")
    r += 2

    row_label(ws, r, "Entry EBITDA (Chart FY2025 adjusted EBITDA)")
    entry_ebitda_row = r
    ws.cell(row=r, column=2, value=f"={ref['chart_adj_ebitda_fy2025']}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Entry enterprise value (stated deal EV)")
    entry_ev_row = r
    ws.cell(row=r, column=2, value=f"={ref['stated_ev']}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Equity purchase price (stated deal equity value)")
    equity_row = r
    ws.cell(row=r, column=2, value=f"={ref['stated_equity_value']}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Refinance existing Chart net debt")
    refi_row = r
    ws.cell(row=r, column=2, value=f"=B{entry_ev_row}-B{equity_row}").number_format = USD_FMT
    r += 2

    row_label(ws, r, "Term Loan B (x entry EBITDA)")
    tlb_row = r
    ws.cell(row=r, column=2, value=f"={ref['lbo_tlb_x']}*B{entry_ebitda_row}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Senior Notes (x entry EBITDA)")
    notes_row = r
    ws.cell(row=r, column=2, value=f"={ref['lbo_notes_x']}*B{entry_ebitda_row}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Total new debt", bold=True)
    newdebt_row = r
    ws.cell(row=r, column=2, value=f"=B{tlb_row}+B{notes_row}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Entry leverage (x EBITDA)")
    lev_row = r
    ws.cell(row=r, column=2, value=f"=B{newdebt_row}/B{entry_ebitda_row}").number_format = X_FMT
    r += 2

    row_label(ws, r, "Advisory fees (% of EV)")
    adv_row = r
    ws.cell(row=r, column=2, value=f"={ref['lbo_advisory_pct']}*B{entry_ev_row}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Financing fees (% of new debt)")
    fee_row = r
    ws.cell(row=r, column=2, value=f"={ref['lbo_fin_fee_pct']}*B{newdebt_row}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Total Uses", bold=True)
    totuses_row = r
    ws.cell(row=r, column=2, value=f"=B{equity_row}+B{refi_row}+B{adv_row}+B{fee_row}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Sponsor equity (plug)", bold=True)
    sponsor_row = r
    ws.cell(row=r, column=2, value=f"=B{totuses_row}-B{newdebt_row}").number_format = USD_FMT
    r += 2

    # Debt schedule
    row_label(ws, r, "Debt Schedule (circularity: interest on average balance; iterative calc is enabled on this workbook)", bold=True)
    r += 1
    for i, year in enumerate(FORECAST_YEARS):
        ws.cell(row=r, column=2 + i, value=year).font = SUBHEADER_FONT
    r += 1

    row_label(ws, r, "Chart standalone EBITDA")
    ebitda_row = r
    prior_sales_ref = ref["chart_sales_fy2025"]
    sales_cells = []
    for i in range(5):
        growth_ref = ref["base_growth"][i]
        mgmt_sales_ref = ref["mgmt_sales"][i]
        mgmt_ebitda_ref = ref["mgmt_ebitda"][i]
        margin_ref = ref["base_margin"][i]
        prior_sales = sales_cells[-1] if sales_cells else prior_sales_ref
        sales_formula_local = f"{prior_sales}*(1+{growth_ref})"
        sales_cells.append(f"({sales_formula_local})" if not sales_cells else f"({sales_formula_local})")
        col = get_column_letter(2 + i)
        formula = f"=IF({ref['scenario']}=2,{mgmt_ebitda_ref},({sales_formula_local})*{margin_ref})"
        ws.cell(row=r, column=2 + i, value=formula).number_format = USD_FMT
        sales_cells[-1] = sales_formula_local
    r += 1

    row_label(ws, r, "Capex (% of sales)")
    capex_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        formula = f"=IF({ref['scenario']}=2,{ref['mgmt_sales'][i]},({sales_cells[i]}))*{ref['capex_pct']}"
        ws.cell(row=r, column=2 + i, value=formula).number_format = USD_FMT
    r += 1

    row_label(ws, r, "Mandatory TLB amortization")
    amort_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        prior_bal_ref = f"{get_column_letter(1 + i)}{None}"  # placeholder, filled below once beg_tlb_row known
        ws.cell(row=r, column=2 + i, value=f"=MIN({ref['lbo_tlb_amort']}*$B${tlb_row}, ${get_column_letter(2+i)}$0)")
    r += 1

    row_label(ws, r, "Beginning TLB balance")
    beg_tlb_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        if i == 0:
            ws.cell(row=r, column=2 + i, value=f"=B{tlb_row}")
        else:
            prev_col = get_column_letter(1 + i)
            ws.cell(row=r, column=2 + i, value=f"={prev_col}{beg_tlb_row + 8}")  # points to ending TLB row, set below
        ws.cell(row=r, column=2 + i).number_format = USD_FMT
    r += 1

    row_label(ws, r, "Beginning Notes balance")
    beg_notes_row = r
    for i in range(5):
        col = get_column_letter(2 + i)
        if i == 0:
            ws.cell(row=r, column=2 + i, value=f"=B{notes_row}")
        else:
            prev_col = get_column_letter(1 + i)
            ws.cell(row=r, column=2 + i, value=f"={prev_col}{beg_notes_row + 8}")
        ws.cell(row=r, column=2 + i).number_format = USD_FMT
    r += 1

    # fix mandatory amort row formula now that beg_tlb_row is known
    for i in range(5):
        col = get_column_letter(2 + i)
        ws.cell(row=amort_row, column=2 + i,
                value=f"=MIN({ref['lbo_tlb_amort']}*$B${tlb_row},{col}{beg_tlb_row})").number_format = USD_FMT

    row_label(ws, r, "Cash sweep (% of post-amortization free cash flow)")
    sweep_row = r
    r += 1
    row_label(ws, r, "Interest expense (avg TLB balance x rate + Notes balance x rate)")
    int_row = r
    r += 1
    row_label(ws, r, "Net income")
    ni_row = r
    r += 1
    row_label(ws, r, "Free cash flow pre-sweep")
    fcf_row = r
    r += 1
    row_label(ws, r, "Ending TLB balance")
    end_tlb_row = r
    r += 1
    row_label(ws, r, "Ending Notes balance")
    end_notes_row = r
    r += 1
    row_label(ws, r, "Ending net debt (TLB + Notes - min cash)")
    end_netdebt_row = r
    r += 2

    for i in range(5):
        col = get_column_letter(2 + i)
        tlb_rate = f"({ref['sofr_at_close']}+{ref['lbo_tlb_margin']})"
        avg_tlb = f"(({col}{beg_tlb_row}+{col}{end_tlb_row})/2)"
        ws.cell(row=int_row, column=2 + i,
                value=f"={avg_tlb}*{tlb_rate}+{col}{beg_notes_row}*{ref['lbo_notes_rate']}").number_format = USD_FMT
        fee_amort_local = f"({ref['lbo_fin_fee_pct']}*B{newdebt_row}/{ref['lbo_fee_amort_years']})"
        ws.cell(row=ni_row, column=2 + i,
                value=f"=({col}{ebitda_row}-{ref['chart_da_fy2025']}-{col}{int_row}-{fee_amort_local})*(1-{ref['tax_rate']})"
                ).number_format = USD_FMT
        ws.cell(row=fcf_row, column=2 + i,
                value=f"={col}{ni_row}+{ref['chart_da_fy2025']}-{col}{capex_row}-{col}{amort_row}").number_format = USD_FMT
        ws.cell(row=sweep_row, column=2 + i,
                value=f"=MAX(MIN({ref['lbo_cash_sweep_pct']}*{col}{fcf_row},{col}{beg_tlb_row}-{col}{amort_row}),0)"
                ).number_format = USD_FMT
        ws.cell(row=end_tlb_row, column=2 + i,
                value=f"=MAX({col}{beg_tlb_row}-{col}{amort_row}-{col}{sweep_row},0)").number_format = USD_FMT
        ws.cell(row=end_notes_row, column=2 + i, value=f"={col}{beg_notes_row}").number_format = USD_FMT
        ws.cell(row=end_netdebt_row, column=2 + i,
                value=f"={col}{end_tlb_row}+{col}{end_notes_row}-{ref['lbo_min_cash']}").number_format = USD_FMT

    # patch beginning-balance formulas to reference the now-known ending-balance rows
    for i in range(1, 5):
        prev_col = get_column_letter(1 + i)
        ws.cell(row=beg_tlb_row, column=2 + i, value=f"={prev_col}{end_tlb_row}")
        ws.cell(row=beg_notes_row, column=2 + i, value=f"={prev_col}{end_notes_row}")

    r += 1
    row_label(ws, r, "Exit EBITDA (final year)")
    exit_ebitda_row = r
    last_col = get_column_letter(2 + 4)
    ws.cell(row=r, column=2, value=f"={last_col}{ebitda_row}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Exit multiple")
    exit_mult_row = r
    ws.cell(row=r, column=2, value=f"={ref['lbo_exit_multiple']}").number_format = X_FMT
    r += 1
    row_label(ws, r, "Exit enterprise value")
    exit_ev_row = r
    ws.cell(row=r, column=2, value=f"=B{exit_ebitda_row}*B{exit_mult_row}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Exit net debt")
    exit_netdebt_row = r
    ws.cell(row=r, column=2, value=f"={last_col}{end_netdebt_row}").number_format = USD_FMT
    r += 1
    row_label(ws, r, "Exit equity value")
    exit_eq_row = r
    ws.cell(row=r, column=2, value=f"=B{exit_ev_row}-B{exit_netdebt_row}").number_format = USD_FMT
    r += 2

    row_label(ws, r, "MOIC", bold=True)
    moic_row = r
    ws.cell(row=r, column=2, value=f"=B{exit_eq_row}/B{sponsor_row}").font = BOLD
    ws.cell(row=r, column=2).number_format = X_FMT
    r += 1
    row_label(ws, r, "IRR (5-year hold)", bold=True)
    irr_row = r
    ws.cell(row=r, column=2, value=f"=B{moic_row}^(1/{ref['lbo_hold_years']})-1").font = BOLD
    ws.cell(row=r, column=2).number_format = PCT_FMT

    refs = {
        "new_debt_row": newdebt_row, "sponsor_equity_row": sponsor_row, "total_uses_row": totuses_row,
        "end_tlb_row": end_tlb_row, "end_notes_row": end_notes_row, "beg_tlb_row": beg_tlb_row,
        "moic_row": moic_row, "irr_row": irr_row,
    }
    return refs


def build_sensitivities_sheet(wb, ref, lbo_ref):
    ws = new_sheet(wb, "Sensitivities")
    r = header(ws, "Sensitivity Grids (explicit formulas, not Excel data tables)")

    row_label(ws, r, "LBO IRR: Entry leverage (x EBITDA) vs. Exit multiple (x)", bold=True)
    r += 1
    lev_steps = [4.0, 5.0, 6.0, 7.0, 8.0]
    exit_steps = [10.0, 11.0, 12.0, 13.0, 14.0]
    ws.cell(row=r, column=1, value="Leverage \\ Exit x")
    for j, xm in enumerate(exit_steps):
        ws.cell(row=r, column=2 + j, value=xm).font = SUBHEADER_FONT
    r += 1
    for lev in lev_steps:
        ws.cell(row=r, column=1, value=lev)
        for j, xm in enumerate(exit_steps):
            new_debt = f"({lev}*{ref['chart_adj_ebitda_fy2025']})"
            uses = f"({ref['stated_equity_value']}+({ref['stated_ev']}-{ref['stated_equity_value']})+{ref['lbo_advisory_pct']}*{ref['stated_ev']}+{ref['lbo_fin_fee_pct']}*{new_debt})"
            sponsor_eq = f"({uses}-{new_debt})"
            exit_ev = f"({ref['chart_adj_ebitda_fy2025']}*{xm})"
            # approximate exit net debt as entry debt held flat (illustrative grid, not the full schedule)
            exit_eq = f"({exit_ev}-{new_debt})"
            moic = f"({exit_eq}/{sponsor_eq})"
            formula = f"=({moic})^(1/{ref['lbo_hold_years']})-1"
            ws.cell(row=r, column=2 + j, value=formula).number_format = PCT_FMT
        r += 1
    r += 2

    row_label(ws, r, "Accretion/(Dilution) Year 3: Premium paid vs. Run-rate synergies achieved", bold=True)
    r += 1
    premiums = [0.20, 0.25, 0.30, 0.35, 0.40]
    syn_levels = [0.50, 0.75, 1.00, 1.25, 1.50]
    ws.cell(row=r, column=1, value="Premium \\ Synergy x stated")
    for j, s in enumerate(syn_levels):
        ws.cell(row=r, column=2 + j, value=s).font = SUBHEADER_FONT
    r += 1
    for prem in premiums:
        ws.cell(row=r, column=1, value=prem)
        for j, s in enumerate(syn_levels):
            price = f"({ref['unaffected_price']}*(1+{prem}))"
            shares = f"{ref['bkr_diluted_shares']}"
            syn_after_tax = f"({ref['stated_synergies_runrate']}*{s}*(1-{ref['tax_rate']}))"
            # illustrative: incremental EPS impact of synergies alone, holding financing/PPA fixed at base case
            formula = f"={syn_after_tax}/{shares}/{ref['bkr_diluted_eps']}"
            ws.cell(row=r, column=2 + j, value=formula).number_format = PCT_FMT
        r += 1
    r += 2

    row_label(ws, r, "Note: leverage/exit grid recomputes MOIC/IRR from entry assumptions directly (net debt held", bold=False)
    r += 1
    row_label(ws, r, "flat at entry for this grid); premium/synergy grid isolates the synergy lever on EPS. See", bold=False)
    r += 1
    row_label(ws, r, "the LBO and Accretion_Dilution sheets for the full, path-dependent base case.", bold=False)


def build_checks_sheet(wb, su_ref, ppa_ref, lbo_ref):
    ws = new_sheet(wb, "Checks")
    r = header(ws, "Formula-Driven Checks")

    row_label(ws, r, "Merger sources = uses")
    c1 = r
    ws.cell(row=r, column=2, value=f"=ABS(Sources_Uses!$B${su_ref['total_sources'].split('$B$')[1]}-Sources_Uses!$B${su_ref['total_uses'].split('$B$')[1]})<0.5")
    r += 1
    row_label(ws, r, "Implied EV within 3% of stated EV")
    c2 = r
    ws.cell(row=r, column=2, value=f"=ABS({su_ref['implied_ev']}/{su_ref['stated_ev']}-1)<0.03")
    r += 1
    row_label(ws, r, "Goodwill is positive")
    c3 = r
    ws.cell(row=r, column=2, value=f"={ppa_ref['goodwill']}>0")
    r += 1
    row_label(ws, r, "LBO sources = uses")
    c4 = r
    ws.cell(row=r, column=2,
            value=f"=ABS((LBO!$B${lbo_ref['new_debt_row']}+LBO!$B${lbo_ref['sponsor_equity_row']})-LBO!$B${lbo_ref['total_uses_row']})<0.5")
    r += 1
    row_label(ws, r, "LBO ending debt balances non-negative (all years)")
    c5 = r
    cols = [get_column_letter(2 + i) for i in range(5)]
    tlb_terms = "+".join(f"({c}{lbo_ref['end_tlb_row']}>=-0.5)" for c in cols)
    notes_terms = "+".join(f"({c}{lbo_ref['end_notes_row']}>=-0.5)" for c in cols)
    ws.cell(row=r, column=2, value=f"=(({tlb_terms})+({notes_terms}))=10")
    r += 2

    row_label(ws, r, "MASTER CHECK (all TRUE)", bold=True)
    master_row = r
    ws.cell(row=r, column=2,
            value=f"=AND(B{c1},B{c2},B{c3},B{c4},B{c5})").font = BOLD
    return master_row


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    wb.properties.creator = "Yazad Madan"
    wb.properties.title = "Baker Hughes / Chart Industries Merger and LBO Model"
    wb.calculation = CalcProperties(iterate=True, iterateCount=100, iterateDelta=0.001, calcMode="auto")

    ref = build_inputs_sheet(wb)
    su_ref = build_sources_uses_sheet(wb, ref)
    ppa_ref = build_ppa_sheet(wb, ref, su_ref)
    pf_ref = build_proforma_is_sheet(wb, ref, su_ref, ppa_ref)
    build_accretion_dilution_sheet(wb, ref, pf_ref)
    lbo_ref = build_lbo_sheet(wb, ref)
    build_sensitivities_sheet(wb, ref, lbo_ref)
    build_checks_sheet(wb, su_ref, ppa_ref, lbo_ref)

    order = ["Inputs", "Sources_Uses", "Purchase_Price_Allocation", "ProForma_IS",
             "Accretion_Dilution", "LBO", "Sensitivities", "Checks"]
    wb._sheets = [wb[name] for name in order]

    EXCEL_OUTPUT.parent.mkdir(exist_ok=True)
    wb.save(EXCEL_OUTPUT)
    print("wrote", EXCEL_OUTPUT)


if __name__ == "__main__":
    main()
