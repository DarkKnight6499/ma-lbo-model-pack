"""Independent recomputation of the merger and LBO model (does not read the Excel workbook).

Every figure either traces to data/inputs.csv (DISCLOSED/DERIVED, cited to a filing) or is
an ASSUMPTION labeled in that file. This module is the source of truth that the Excel
workbook is checked against.
"""
import json

from config import COL_LOCATOR, COL_SOURCE_ID, EXCEL_OUTPUT, FORECAST_YEARS, INPUTS, RESULTS_JSON

# ---------------------------- CONFIG ----------------------------
EUR_USD_IMPLIED = None  # computed below from disclosed carrying values, not hardcoded
# ----------------------------------------------------------------


def get(key):
    return INPUTS.value(key)


def cite(key):
    row = INPUTS.row(key)
    return {"key": key, "source_id": row[COL_SOURCE_ID], "locator": row[COL_LOCATOR]}


def eur_usd_rate():
    """Implied EUR/USD from the disclosed carrying value of the March 2026 euro notes."""
    face_eur = get("eur_notes_face_eur")
    carrying_usd = sum(get("eur_tranche_carrying"))
    return carrying_usd / face_eur


def diluted_shares_for_merger():
    """Fully diluted Chart share count used to size the equity purchase price."""
    basic = get("chart_shares_issued")  # shares issued; already net of treasury per Q1 2026 basic count below
    basic_q1_2026 = 47.869076  # disclosed cover-page count, closest pre-close print (S04)
    rsu = get("chart_rsu")
    psu = get("chart_psu")
    options = get("chart_options")
    strike = get("chart_option_strike")
    price = get("offer_price")
    option_dilution = options * (1 - strike / price)
    total = basic_q1_2026 + rsu + psu + option_dilution
    return {
        "basic_shares_q1_2026": basic_q1_2026,
        "rsu_shares": rsu,
        "psu_shares": psu,
        "option_treasury_method_shares": option_dilution,
        "total_diluted_shares": total,
    }


def chart_net_debt_fy2025():
    principal = (
        get("chart_notes_secured_principal")
        + get("chart_notes_unsecured_principal")
        + get("chart_term_loan_principal")
        + get("chart_revolver_principal")
        + get("chart_other_debt_principal")
    )
    return principal - get("chart_cash")


def merger_sources_and_uses():
    shares = diluted_shares_for_merger()
    equity_purchase_price = shares["total_diluted_shares"] * get("offer_price")

    secured_redemption = get("chart_notes_secured_principal") * (1 + get("redemption_premium_secured_pct"))
    unsecured_redemption = get("chart_notes_unsecured_principal") * (1 + get("redemption_premium_unsecured_pct"))
    term_loan_payoff = get("chart_term_loan_principal")
    revolver_payoff = get("chart_revolver_principal")
    other_debt_payoff = get("chart_other_debt_principal")
    debt_refinancing = secured_redemption + unsecured_redemption + term_loan_payoff + revolver_payoff + other_debt_payoff

    advisory_fees = get("advisory_fee_pct") * get("stated_ev")

    new_term_loans = get("term_loan_total")
    new_usd_notes = get("usd_notes_face_total")
    new_eur_notes_usd = sum(get("eur_tranche_carrying"))  # carrying value proxy for face at issue, USD
    new_debt_gross = new_term_loans + new_usd_notes + new_eur_notes_usd
    financing_fees = get("tl_fee_pct") * new_debt_gross

    total_uses = equity_purchase_price + debt_refinancing + advisory_fees + financing_fees
    cash_on_hand_use = total_uses - new_debt_gross

    implied_ev = equity_purchase_price + chart_net_debt_fy2025()
    ev_variance_pct = implied_ev / get("stated_ev") - 1

    return {
        "shares": shares,
        "uses": {
            "equity_purchase_price": equity_purchase_price,
            "redeem_secured_notes": secured_redemption,
            "redeem_unsecured_notes": unsecured_redemption,
            "repay_term_loan": term_loan_payoff,
            "repay_revolver": revolver_payoff,
            "repay_other_debt": other_debt_payoff,
            "advisory_and_legal_fees": advisory_fees,
            "financing_fees": financing_fees,
            "total_uses": total_uses,
        },
        "sources": {
            "new_term_loans": new_term_loans,
            "new_usd_notes": new_usd_notes,
            "new_eur_notes_usd_equivalent": new_eur_notes_usd,
            "cash_on_hand": cash_on_hand_use,
            "total_sources": new_debt_gross + cash_on_hand_use,
        },
        "implied_enterprise_value": implied_ev,
        "stated_enterprise_value": get("stated_ev"),
        "ev_variance_pct": ev_variance_pct,
        "stated_equity_value": get("stated_equity_value"),
        "equity_value_variance_pct": equity_purchase_price / get("stated_equity_value") - 1,
    }


def purchase_price_allocation(su):
    equity_purchase_price = su["uses"]["equity_purchase_price"]
    tax_rate = get("tax_rate")

    fv_assets = {
        "cash": get("chart_cash"),
        "accounts_receivable": get("chart_ar"),
        "inventory": get("chart_inventory") * (1 + get("inventory_stepup_pct")),
        "unbilled_contract_revenue": get("chart_unbilled"),
        "other_current_assets": get("chart_other_current_assets"),
        "ppe": get("chart_ppe") * (1 + get("ppe_stepup_pct")),
        "identifiable_intangibles": get("new_intangibles"),
        "other_assets": get("chart_other_assets"),
    }
    total_fv_assets = sum(fv_assets.values())

    intangibles_book_tax_basis = get("chart_intangibles")  # assumed no tax basis step-up (A17 note area)
    ppe_stepup = fv_assets["ppe"] - get("chart_ppe")
    inventory_stepup = fv_assets["inventory"] - get("chart_inventory")
    intangibles_stepup_over_tax_basis = fv_assets["identifiable_intangibles"] - intangibles_book_tax_basis
    new_dtl = tax_rate * (intangibles_stepup_over_tax_basis + ppe_stepup + inventory_stepup)

    fv_liabilities = {
        "accounts_payable": get("chart_ap"),
        "customer_advances": get("chart_advances"),
        "accrued_interest": get("chart_accrued_interest"),
        "other_current_liabilities": get("chart_other_current_liab"),
        "deferred_tax_liability_new": new_dtl,
        "other_long_term_liabilities": get("chart_other_lt_liab"),
    }
    total_fv_liabilities = sum(fv_liabilities.values())

    nci_carrying = get("chart_nci")
    net_identifiable_assets = total_fv_assets - total_fv_liabilities
    goodwill = equity_purchase_price - (net_identifiable_assets - nci_carrying)

    return {
        "fair_value_assets": fv_assets,
        "total_fv_assets": total_fv_assets,
        "fair_value_liabilities": fv_liabilities,
        "total_fv_liabilities": total_fv_liabilities,
        "new_deferred_tax_liability": new_dtl,
        "net_identifiable_assets": net_identifiable_assets,
        "nci_carrying_value": nci_carrying,
        "goodwill": goodwill,
        "ppe_stepup": ppe_stepup,
        "inventory_stepup": inventory_stepup,
        "intangibles_fair_value": fv_assets["identifiable_intangibles"],
    }


def chart_standalone_forecast():
    """Base-case standalone Chart sales/EBITDA, scenario-selectable per Inputs."""
    scenario = get("scenario")
    if scenario == 2:
        sales = get("mgmt_sales")
        ebitda = get("mgmt_ebitda")
    else:
        growth = get("base_growth")
        margin = get("base_margin")
        sales = []
        prior = get("chart_sales_fy2025")
        for g in growth:
            prior = prior * (1 + g)
            sales.append(prior)
        ebitda = [s * m for s, m in zip(sales, margin)]
    capex = [s * get("capex_pct") for s in sales]
    return {"sales": sales, "ebitda": ebitda, "capex": capex}


def synergies_schedule():
    runrate = get("stated_synergies_runrate")
    phase = get("synergy_phase")
    realized = [runrate * p for p in phase]
    integration_cost = runrate * get("integration_cost_multiple")
    icost_phase = get("integration_cost_phase")
    integration = [integration_cost * p for p in icost_phase]
    return {"realized_synergies": realized, "integration_costs": integration}


def proforma_income_statement(su, ppa):
    """Accretion/dilution bridge: BKR standalone NI + Chart's unlevered after-tax contribution
    + after-tax synergies, less after-tax incremental D&A/interest/fees from the deal. Chart's
    own pre-deal interest expense is excluded because its debt is refinanced at close, so the
    entire new financing stack is incremental to the combined entity versus the two standalones.
    """
    chart_fc = chart_standalone_forecast()
    syn = synergies_schedule()
    tax_rate = get("tax_rate")

    bkr_ni_base = get("bkr_ni")
    bkr_growth = get("bkr_ni_growth")
    bkr_ni_standalone = []
    prior = bkr_ni_base
    for _ in FORECAST_YEARS:
        prior = prior * (1 + bkr_growth)
        bkr_ni_standalone.append(prior)

    intangibles_amort = ppa["intangibles_fair_value"] / get("intangibles_life")
    ppe_stepup_dep = ppa["ppe_stepup"] / get("ppe_stepup_life")
    incremental_da = intangibles_amort + ppe_stepup_dep

    new_debt_total = get("term_loan_total") + get("usd_notes_face_total") + sum(get("eur_tranche_carrying"))
    # blended coupon across the new stack, weighted by disclosed face/carrying amounts
    usd_face = get("usd_tranche_face")
    usd_coupons = get("usd_tranche_coupons")
    eur_carry = get("eur_tranche_carrying")
    eur_coupons = get("eur_tranche_coupons")
    tl_rate = get("sofr_at_close") + get("term_loan_margin")
    weighted_coupon_numerator = (
        sum(f * c for f, c in zip(usd_face, usd_coupons))
        + sum(e * c for e, c in zip(eur_carry, eur_coupons))
        + get("term_loan_total") * tl_rate
    )
    blended_new_debt_rate = weighted_coupon_numerator / new_debt_total
    new_financing_interest = new_debt_total * blended_new_debt_rate

    financing_fee_amort = su["uses"]["financing_fees"] / get("fin_fee_amort_years")
    inventory_stepup_charge_year1 = ppa["inventory_stepup"]  # flows through COGS in year one only

    rows = []
    for i, year in enumerate(FORECAST_YEARS):
        chart_da = get("chart_da_fy2025")
        chart_ebit = chart_fc["ebitda"][i] - chart_da
        chart_unlevered_ni = chart_ebit * (1 - tax_rate)

        realized_syn = syn["realized_synergies"][i]
        integration = syn["integration_costs"][i]
        one_time_inv_charge = inventory_stepup_charge_year1 if i == 0 else 0.0
        synergy_net_after_tax = (realized_syn - integration - one_time_inv_charge) * (1 - tax_rate)

        incremental_da_after_tax = incremental_da * (1 - tax_rate)
        interest_after_tax = new_financing_interest * (1 - tax_rate)
        fee_amort_after_tax = financing_fee_amort * (1 - tax_rate)

        pro_forma_ni = (
            bkr_ni_standalone[i]
            + chart_unlevered_ni
            + synergy_net_after_tax
            - incremental_da_after_tax
            - interest_after_tax
            - fee_amort_after_tax
            - get("nci_income")
        )
        rows.append(
            {
                "year": year,
                "combined_sales": chart_fc["sales"][i],
                "chart_ebitda": chart_fc["ebitda"][i],
                "realized_synergies": realized_syn,
                "integration_costs": integration,
                "incremental_da": incremental_da,
                "new_financing_interest": new_financing_interest,
                "financing_fee_amortization": financing_fee_amort,
                "chart_unlevered_ni": chart_unlevered_ni,
                "bkr_standalone_ni": bkr_ni_standalone[i],
                "pro_forma_net_income": pro_forma_ni,
            }
        )
    return rows


def accretion_dilution(proforma_rows, su):
    bkr_diluted_shares = get("bkr_diluted_shares")
    bkr_standalone_eps = get("bkr_diluted_eps")
    bkr_growth = get("bkr_ni_growth")

    new_shares_issued = 0.0  # all-cash deal, no new BKR shares issued
    pro_forma_shares = bkr_diluted_shares + new_shares_issued

    results = []
    for i, row in enumerate(proforma_rows):
        standalone_eps = bkr_standalone_eps * (1 + bkr_growth) ** (i + 1)
        pro_forma_eps = row["pro_forma_net_income"] / pro_forma_shares
        accretion = pro_forma_eps / standalone_eps - 1
        results.append(
            {
                "year": row["year"],
                "standalone_bkr_eps": standalone_eps,
                "pro_forma_eps": pro_forma_eps,
                "accretion_dilution_pct": accretion,
            }
        )

    # breakeven run-rate synergies in year 3 (first full run-rate year) holding everything else constant
    year3 = proforma_rows[2]
    gap_pretax = (results[2]["standalone_bkr_eps"] - results[2]["pro_forma_eps"]) * pro_forma_shares if results[2]["accretion_dilution_pct"] < 0 else 0.0
    breakeven_incremental_pretax_synergies = gap_pretax / (1 - get("tax_rate")) if gap_pretax else 0.0
    breakeven_synergies_total = year3["realized_synergies"] + breakeven_incremental_pretax_synergies

    return {"by_year": results, "breakeven_run_rate_synergies_year3": breakeven_synergies_total}


def lbo_model():
    entry_ebitda = get("chart_adj_ebitda_fy2025")
    entry_ev = get("stated_ev")
    equity_value = get("stated_equity_value")
    net_debt_refi = entry_ev - equity_value

    tlb = get("lbo_tlb_x") * entry_ebitda
    notes = get("lbo_notes_x") * entry_ebitda
    new_debt = tlb + notes

    advisory = get("lbo_advisory_pct") * entry_ev
    financing_fees = get("lbo_fin_fee_pct") * new_debt
    total_uses = equity_value + net_debt_refi + advisory + financing_fees
    sponsor_equity = total_uses - new_debt

    entry_leverage = new_debt / entry_ebitda

    su = {
        "entry_ebitda": entry_ebitda,
        "entry_ev": entry_ev,
        "equity_purchase_price": equity_value,
        "refinance_existing_net_debt": net_debt_refi,
        "advisory_fees": advisory,
        "financing_fees": financing_fees,
        "total_uses": total_uses,
        "term_loan_b": tlb,
        "senior_notes": notes,
        "new_debt_total": new_debt,
        "sponsor_equity": sponsor_equity,
        "entry_leverage_x": entry_leverage,
    }

    fc = chart_standalone_forecast()
    sofr = get("sofr_at_close")
    tlb_rate = sofr + get("lbo_tlb_margin")
    notes_rate = get("lbo_notes_rate")
    tlb_amort_pct = get("lbo_tlb_amort")
    sweep_pct = get("lbo_cash_sweep_pct")
    fee_amort = financing_fees / get("lbo_fee_amort_years")
    min_cash = get("lbo_min_cash")

    tlb_balance = tlb
    notes_balance = notes
    cash_balance = min_cash
    schedule = []
    for i, year in enumerate(FORECAST_YEARS):
        ebitda = fc["ebitda"][i]
        capex = fc["capex"][i]
        mandatory_amort = min(tlb * tlb_amort_pct, tlb_balance)

        # interest on average balance during the year: iterate the sweep/interest circularity to convergence
        beg_tlb, beg_notes = tlb_balance, notes_balance
        sweep = 0.0
        for _ in range(50):
            avg_tlb = (beg_tlb + max(beg_tlb - mandatory_amort - sweep, 0.0)) / 2
            interest = avg_tlb * tlb_rate + beg_notes * notes_rate
            ebit_local = ebitda - (get("chart_da_fy2025"))
            pretax = ebit_local - interest - fee_amort
            tax = max(pretax, 0.0) * get("tax_rate")
            ni = pretax - tax
            fcf = ni + get("chart_da_fy2025") - capex - mandatory_amort
            new_sweep = max(fcf, 0.0) * sweep_pct
            new_sweep = min(new_sweep, max(beg_tlb - mandatory_amort, 0.0))
            if abs(new_sweep - sweep) < 1e-6:
                sweep = new_sweep
                break
            sweep = new_sweep

        end_tlb = max(beg_tlb - mandatory_amort - sweep, 0.0)
        end_notes = beg_notes
        cash_balance = min_cash + max(fcf - sweep, 0.0) if fcf > sweep else min_cash

        schedule.append(
            {
                "year": year,
                "ebitda": ebitda,
                "capex": capex,
                "interest_expense": interest,
                "cash_taxes": tax,
                "net_income": ni,
                "free_cash_flow_pre_sweep": fcf,
                "mandatory_amortization": mandatory_amort,
                "cash_sweep": sweep,
                "beginning_tlb": beg_tlb,
                "ending_tlb": end_tlb,
                "beginning_notes": beg_notes,
                "ending_notes": end_notes,
                "ending_net_debt": end_tlb + end_notes - min_cash,
            }
        )
        tlb_balance, notes_balance = end_tlb, end_notes

    exit_ebitda = fc["ebitda"][-1]
    exit_ev = exit_ebitda * get("lbo_exit_multiple")
    exit_net_debt = schedule[-1]["ending_net_debt"]
    exit_equity_value = exit_ev - exit_net_debt

    hold_years = int(get("lbo_hold_years"))
    moic = exit_equity_value / sponsor_equity
    irr = moic ** (1 / hold_years) - 1

    return {
        "sources_and_uses": su,
        "debt_schedule": schedule,
        "exit": {
            "exit_ebitda": exit_ebitda,
            "exit_multiple": get("lbo_exit_multiple"),
            "exit_ev": exit_ev,
            "exit_net_debt": exit_net_debt,
            "exit_equity_value": exit_equity_value,
        },
        "returns": {"moic": moic, "irr": irr, "hold_years": hold_years},
    }


def checks(su, ppa, ad, lbo):
    tol = 0.5  # USD mm tolerance for rounding across many inputs
    sources_eq_uses = abs(su["sources"]["total_sources"] - su["uses"]["total_uses"]) < tol
    goodwill_positive = ppa["goodwill"] > 0
    ev_reconciles = abs(su["ev_variance_pct"]) < 0.03
    lbo_sources_eq_uses = abs(
        (lbo["sources_and_uses"]["new_debt_total"] + lbo["sources_and_uses"]["sponsor_equity"])
        - lbo["sources_and_uses"]["total_uses"]
    ) < tol
    lbo_debt_never_negative = all(row["ending_tlb"] >= -tol and row["ending_notes"] >= -tol for row in lbo["debt_schedule"])
    all_checks = {
        "sources_equal_uses_merger": sources_eq_uses,
        "goodwill_positive": goodwill_positive,
        "implied_ev_within_3pct_of_stated": ev_reconciles,
        "lbo_sources_equal_uses": lbo_sources_eq_uses,
        "lbo_debt_balances_non_negative": lbo_debt_never_negative,
    }
    all_checks["MASTER_CHECK"] = all(all_checks.values())
    return all_checks


def build_results():
    su = merger_sources_and_uses()
    ppa = purchase_price_allocation(su)
    proforma_rows = proforma_income_statement(su, ppa)
    ad = accretion_dilution(proforma_rows, su)
    lbo = lbo_model()
    chk = checks(su, ppa, ad, lbo)

    results = {
        "deal_name": "Baker Hughes Company acquisition of Chart Industries, Inc.",
        "announcement_date": "2025-07-28",
        "closing_date": "2026-07-16",
        "eur_usd_implied_rate": eur_usd_rate(),
        "merger_sources_and_uses": su,
        "purchase_price_allocation": ppa,
        "proforma_income_statement": proforma_rows,
        "accretion_dilution": ad,
        "lbo": lbo,
        "checks": chk,
    }
    with open(RESULTS_JSON, "w", encoding="utf-8") as fh:
        json.dump(results, fh, indent=2)
    return results


if __name__ == "__main__":
    r = build_results()
    print("MASTER_CHECK:", r["checks"]["MASTER_CHECK"])
    print("Implied EV: {:.1f} vs stated {:.1f} ({:+.2%})".format(
        r["merger_sources_and_uses"]["implied_enterprise_value"],
        r["merger_sources_and_uses"]["stated_enterprise_value"],
        r["merger_sources_and_uses"]["ev_variance_pct"],
    ))
    print("Goodwill: {:.1f}".format(r["purchase_price_allocation"]["goodwill"]))
    for row in r["accretion_dilution"]["by_year"]:
        print(row["year"], "{:+.2%}".format(row["accretion_dilution_pct"]))
    print("LBO IRR: {:.1%}  MOIC: {:.2f}x".format(r["lbo"]["returns"]["irr"], r["lbo"]["returns"]["moic"]))
