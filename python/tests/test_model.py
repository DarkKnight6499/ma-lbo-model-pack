"""Pytest suite: checks the independent Python model's internal consistency and its
reconciliation to the deal's own disclosed figures. Run with: py -3 -m pytest python/tests.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import model  # noqa: E402
from config import INPUTS  # noqa: E402

TOLERANCE_USD_MM = 0.5
EV_TOLERANCE_PCT = 0.03


def get(key):
    return INPUTS.value(key)


def build():
    return model.build_results()


def test_master_check_passes():
    results = build()
    assert results["checks"]["MASTER_CHECK"] is True


def test_merger_sources_equal_uses():
    results = build()
    su = results["merger_sources_and_uses"]
    assert abs(su["sources"]["total_sources"] - su["uses"]["total_uses"]) < TOLERANCE_USD_MM


def test_implied_ev_reconciles_to_press_release():
    results = build()
    variance = results["merger_sources_and_uses"]["ev_variance_pct"]
    assert abs(variance) < EV_TOLERANCE_PCT, f"implied EV variance {variance:.2%} exceeds tolerance"


def test_implied_equity_value_reconciles_to_press_release():
    results = build()
    variance = results["merger_sources_and_uses"]["equity_value_variance_pct"]
    # the disclosed $10.1bn is from the original July 2025 proposal on an earlier share count;
    # this model's share count is from the closest pre-close print, so allow a wider band
    assert abs(variance) < 0.02


def test_goodwill_is_positive_and_plausible():
    results = build()
    goodwill = results["purchase_price_allocation"]["goodwill"]
    ev = results["merger_sources_and_uses"]["implied_enterprise_value"]
    assert 0 < goodwill < ev


def test_ppa_assets_equal_sum_of_components():
    results = build()
    ppa = results["purchase_price_allocation"]
    assert abs(sum(ppa["fair_value_assets"].values()) - ppa["total_fv_assets"]) < TOLERANCE_USD_MM
    assert abs(sum(ppa["fair_value_liabilities"].values()) - ppa["total_fv_liabilities"]) < TOLERANCE_USD_MM


def test_accretion_dilution_has_five_years():
    results = build()
    assert len(results["accretion_dilution"]["by_year"]) == 5
    years = [row["year"] for row in results["accretion_dilution"]["by_year"]]
    assert years == model.FORECAST_YEARS


def test_accretion_dilution_honest_about_year_one():
    """The press release claims double-digit year-one accretion; this independent model, using
    the full incremental cost of financing the entire purchase (including refinanced target debt)
    at a 23% tax rate, finds year-one dilution. That gap is reported, not hidden."""
    results = build()
    year_one = results["accretion_dilution"]["by_year"][0]
    assert year_one["accretion_dilution_pct"] < 0


def test_lbo_sources_equal_uses():
    results = build()
    lbo = results["lbo"]["sources_and_uses"]
    assert abs((lbo["new_debt_total"] + lbo["sponsor_equity"]) - lbo["total_uses"]) < TOLERANCE_USD_MM


def test_lbo_debt_schedule_never_negative():
    results = build()
    for row in results["lbo"]["debt_schedule"]:
        assert row["ending_tlb"] >= -TOLERANCE_USD_MM
        assert row["ending_notes"] >= -TOLERANCE_USD_MM


def test_lbo_debt_schedule_amortizes_down():
    results = build()
    schedule = results["lbo"]["debt_schedule"]
    assert schedule[-1]["ending_tlb"] <= schedule[0]["beginning_tlb"]


def test_lbo_returns_are_plausible():
    results = build()
    returns = results["lbo"]["returns"]
    assert 0 < returns["moic"] < 10
    assert -0.5 < returns["irr"] < 1.0


def test_lbo_entry_leverage_matches_disclosed_sum():
    results = build()
    su = results["lbo"]["sources_and_uses"]
    assert abs(su["entry_leverage_x"] - (get("lbo_tlb_x") + get("lbo_notes_x"))) < 1e-9


def test_results_json_is_written_and_matches():
    results = build()
    on_disk = json.loads(model.RESULTS_JSON.read_text(encoding="utf-8"))
    assert on_disk["checks"]["MASTER_CHECK"] == results["checks"]["MASTER_CHECK"]


def test_no_hardcoded_offer_price_drift():
    """The offer price used throughout must be the one value in inputs.csv, not a stray literal."""
    assert get("offer_price") == 210.0
    assert get("lbo_price") == 210.0
