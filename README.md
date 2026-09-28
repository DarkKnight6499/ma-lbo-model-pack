# M&A and LBO Model Pack: Baker Hughes / Chart Industries

A public-filings-only merger model and standalone LBO for one real, completed US public-company
acquisition: Baker Hughes Company's (NASDAQ: BKR) acquisition of Chart Industries, Inc. (NYSE: GTLS).
Every disclosed figure traces to a specific SEC filing, table and paragraph; every figure the filings
do not disclose is labeled as an illustrative assumption with its rationale. The LBO section is a
clearly marked hypothetical, since this deal was a strategic acquisition, not a sponsor buyout.

## Why this deal

The brief called for a mid-size ($2 billion to $15 billion enterprise value) US public-company deal
announced within roughly the last 18 months, with a definitive merger agreement, clean financials for
both parties, and a disclosed price, consideration mix and financing plan. Baker Hughes / Chart fits
all of that:

- Announced July 28, 2025 (8-K, Exhibit 99.1 press release) and closed July 16, 2026 (8-K).
- All-cash, single-price consideration: $210.00 per Chart share, no collar or stock mix to model
  around.
- A stated enterprise value (~$13.6 billion) and equity value (~$10.1 billion) squarely in the
  target size range.
- Fully disclosed financing: a $14.9 billion bridge facility at signing, later replaced by a $2.6
  billion delayed-draw term loan, then a March 2026 offering of $6.5 billion of USD notes and EUR 3.0
  billion of euro notes, and finally two $1.0 billion two-year term loans signed the day before
  closing. Chart's own notes, term loan and revolver were redeemed or repaid at close.
- Both companies file 10-Ks and 10-Qs with SEC EDGAR, and Chart's DEFM14A (definitive merger proxy)
  discloses management's own financial projections and the fairness-opinion valuation ranges used
  by its financial advisor, Wells Fargo.

Because the actual deal was a strategic, all-cash acquisition, the merger model (Sources & Uses,
purchase price allocation, pro forma income statement, accretion/dilution) is built on the deal's
actual, disclosed terms. The LBO section is a labeled hypothetical: what a financial sponsor could
have paid for Chart at the same $210.00 per share price, using an illustrative capital structure,
purely so the two forms of value creation can be compared side by side. It is not a claim that a
sponsor bid for Chart, disclosed or otherwise, ever existed.

## What is in the repository

```
data/               Raw filing excerpts, a provenance table (sources.csv), and the single
                     inputs.csv that holds every disclosed figure and every labeled assumption
excel/               ma_lbo_model.xlsx: a formula-driven Excel workbook
python/               An independent Python recomputation, the workbook builder, chart and
                     docx builders, and the pytest suite
docs/                pitch_summary.docx (generated from results.json) and its charts
```

## Method

1. **Data collection** (`python/fetch_data.py`): downloads the relevant 8-K, 10-K, 10-Q and
   DEFM14A filings directly from SEC EDGAR, extracts the tables and paragraphs the model needs,
   and snapshots them as text excerpts under `data/excerpts/` together with a FRED SOFR series
   used for the term loan floating rate. Every value used downstream is recorded in
   `data/inputs.csv` with its source filing and locator (or, for an assumption, its rationale),
   and `data/sources.csv` lists every source with its retrieval date and status.
2. **Independent Python model** (`python/model.py`): computes the fully diluted share count,
   Sources & Uses, purchase price allocation (goodwill, intangibles, a rebuilt deferred tax
   liability), the pro forma accretion/dilution bridge, and the LBO debt schedule and returns,
   reading only from `data/inputs.csv`. This module does not read the Excel workbook.
3. **Excel workbook** (`python/build_workbook.py`): builds the same model as live Excel
   formulas (via openpyxl), referencing a single Inputs sheet, so every output cell can be
   traced back to a disclosed figure or a labeled assumption cell. The LBO debt schedule has a
   genuine circularity (interest depends on the average balance, which depends on the cash
   sweep, which depends on interest); the workbook enables Excel's iterative calculation
   (100 iterations, 0.001 tolerance) to resolve it, and this is documented on the LBO sheet.
4. **Cross-check**: the workbook is recalculated headlessly with LibreOffice and every output
   cell compared against the independent Python model in `python/tests/test_workbook_recalc.py`.
   The LBO returns are allowed a 2% tolerance because Excel's whole-sheet iterative solver and
   the Python per-year fixed-point iteration converge the circularity slightly differently;
   every other check uses a tight (under 1%) or exact tolerance.
5. **Pitch summary** (`python/build_pitch_summary.py`): builds `docs/pitch_summary.docx`
   entirely from `python/results.json`; no figure in the document is typed by hand.

## Headline results

All figures are this model's own computation (see Limitations for what is illustrative).

**Merger (actual deal terms)**
- Fully diluted Chart share count used for consideration: approximately 48.2 million shares
- Equity purchase price: approximately $10,125 million
- Implied enterprise value: approximately $13,415 million, about 1.4% below the disclosed
  $13,600 million (explained in the model by differing share-count and net-debt measurement
  dates; see Limitations)
- Total Sources = Total Uses: approximately $13,992 million
- Goodwill: approximately $3,962 million (on illustrative purchase-price-allocation step-ups)
- GAAP EPS accretion/(dilution) by year: 2026 about -9.2%, 2027 about -0.8%, 2028 about +5.8%,
  2029 about +9.3%, 2030 about +10.8%

**LBO (hypothetical sponsor bid, same $210.00/share price)**
- Entry leverage: 6.0x EBITDA ($4,564 million Term Loan B + $1,521 million senior notes)
- Sponsor equity: approximately $7,802 million
- 5-year hold, exit at 12.0x EBITDA
- MOIC: approximately 1.61x; IRR: approximately 10.0%

The accretion/dilution and LBO return figures are reported honestly even where they read as
weak: this model finds the deal modestly dilutive in year one against the company's own
double-digit-accretion claim (see Limitations for why), and finds the hypothetical LBO's
returns well below a typical sponsor's hurdle rate at this price and leverage.

## Limitations

- **Purchase price allocation is illustrative.** Baker Hughes has not publicly disclosed a
  final fair-value allocation of the Chart purchase price as of this writing. The step-ups
  applied to inventory, PP&E and identifiable intangibles, and the resulting deferred tax
  liability, are the author's stated assumptions, not disclosed figures.
- **The accretion/dilution result differs from the company's own claim.** Baker Hughes' press
  release states the deal is "immediately accretive to growth, margins and cash flow, with
  double-digit EPS accretion in the first full year." This model, which charges the combined
  entity the full incremental interest cost of financing the entire purchase price (including
  refinancing Chart's existing debt) at an illustrative 23% tax rate and layers in amortization
  of newly created intangibles, instead finds year-one dilution, turning accretive from year
  three. The gap most likely reflects a non-GAAP or cash-EPS basis in the company's own claim, a
  lower assumed financing cost, and/or synergies realized faster than this model's illustrative
  ramp. That gap is reported rather than adjusted away.
- **Financing terms not fully disclosed.** The March 2026 notes offering and July 2026 term
  loans are disclosed by coupon, face amount and maturity, but the ratings-based pricing grids
  and exact all-in fees are not; a term loan margin and financing/advisory fee percentages are
  stated assumptions.
- **Tax rate is a simplifying, deliberately conservative assumption.** A blended 23% rate is
  used for deal adjustments; Baker Hughes' own FY2025 effective tax rate was materially lower
  (8.8%), so this likely understates pro forma net income relative to the combined company's
  actual cash tax position.
- **Synergy phasing is illustrative.** The $325 million year-three run-rate figure is Baker
  Hughes' own disclosed number; the year-by-year ramp and the one-time cost to achieve it are
  not disclosed and are assumed.
- **The LBO is entirely hypothetical.** Chart was not taken private. The leverage split, pricing,
  cash sweep percentage and exit multiple are illustrative choices for a cyclical industrial
  credit, not a real financing that was arranged or contemplated.
- **Measurement-date mismatch.** The merger model's net debt and PP&E/inventory bases use
  Chart's FY2025 (December 31, 2025) 10-K, the last full balance sheet with a complete debt
  footnote available before close; actual balances at the July 16, 2026 closing date were not
  separately broken out in the filings reviewed.
- **Noncontrolling interests** in Chart's consolidated subsidiaries are measured at carrying
  value rather than an independently appraised fair value, a simplifying assumption noted in
  `data/inputs.csv`.

## Reproducing the model

```
py -3 -m pip install -r python/requirements.txt
py -3 python/fetch_data.py          # re-downloads and re-snapshots the SEC/FRED data
py -3 python/model.py               # runs the independent Python model, writes results.json
py -3 python/build_workbook.py      # builds excel/ma_lbo_model.xlsx
py -3 python/make_charts.py         # builds docs/charts/*.png from results.json
py -3 python/build_pitch_summary.py # builds docs/pitch_summary.docx from results.json
py -3 -m pytest python/tests -q     # runs the full test suite, including the Excel cross-check
```

## Disclaimer

This project is an educational, illustrative analysis built from public SEC filings and clearly
labeled assumptions. It is not investment advice, a fairness opinion, or a solicitation to buy or
sell any security, and it was not prepared by, for, or in coordination with either company named
above.
