"""Builds the two charts used in docs/pitch_summary.docx, from results.json only."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from config import CHARTS_DIR, RESULTS_JSON

# ---------------------------- CONFIG ----------------------------
ACCENT = "#2E5C8A"
NEGATIVE = "#B3541E"
FIG_DPI = 150
# ----------------------------------------------------------------


def load_results():
    return json.loads(RESULTS_JSON.read_text(encoding="utf-8"))


def chart_accretion_dilution(results, out_path):
    rows = results["accretion_dilution"]["by_year"]
    years = [str(r["year"]) for r in rows]
    values = [r["accretion_dilution_pct"] * 100 for r in rows]
    colors = [NEGATIVE if v < 0 else ACCENT for v in values]

    fig, ax = plt.subplots(figsize=(6.5, 3.6))
    ax.bar(years, values, color=colors)
    ax.axhline(0, color="#444444", linewidth=0.8)
    ax.set_ylabel("GAAP EPS accretion / (dilution), %")
    ax.set_title("Pro forma EPS accretion / (dilution) by year")
    for i, v in enumerate(values):
        ax.annotate(f"{v:+.1f}%", (i, v), textcoords="offset points",
                    xytext=(0, 4 if v >= 0 else -14), ha="center", fontsize=9)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(out_path, dpi=FIG_DPI)
    plt.close(fig)


def chart_lbo_debt_paydown(results, out_path):
    schedule = results["lbo"]["debt_schedule"]
    years = [str(r["year"]) for r in schedule]
    tlb = [r["ending_tlb"] for r in schedule]
    notes = [r["ending_notes"] for r in schedule]

    fig, ax = plt.subplots(figsize=(6.5, 3.6))
    ax.bar(years, tlb, label="Term Loan B", color=ACCENT)
    ax.bar(years, notes, bottom=tlb, label="Senior Notes", color="#8AA9C4")
    ax.set_ylabel("USD millions")
    ax.set_title("LBO debt balance, end of year")
    ax.legend(frameon=False, fontsize=9)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(out_path, dpi=FIG_DPI)
    plt.close(fig)


def main():
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    results = load_results()
    chart_accretion_dilution(results, CHARTS_DIR / "accretion_dilution.png")
    chart_lbo_debt_paydown(results, CHARTS_DIR / "lbo_debt_paydown.png")
    print("wrote charts to", CHARTS_DIR)


if __name__ == "__main__":
    main()
