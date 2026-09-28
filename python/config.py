"""Single config block: load every user-facing input from data/inputs.csv by name.

Nothing downstream should hardcode a filing figure or an assumption; every
number in the model is looked up here by its INPUTS key, never retyped.
"""
import csv
from pathlib import Path

# ---------------------------- CONFIG ----------------------------
REPO_ROOT = Path(__file__).resolve().parent.parent
INPUTS_CSV = REPO_ROOT / "data" / "inputs.csv"
SOURCES_CSV = REPO_ROOT / "data" / "sources.csv"
RESULTS_JSON = REPO_ROOT / "python" / "results.json"
EXCEL_OUTPUT = REPO_ROOT / "excel" / "ma_lbo_model.xlsx"
DOCX_OUTPUT = REPO_ROOT / "docs" / "pitch_summary.docx"
CHARTS_DIR = REPO_ROOT / "docs" / "charts"
DEAL_NAME = "Baker Hughes / Chart Industries"
AUTHOR = "Yazad Madan"
FORECAST_YEARS = [2026, 2027, 2028, 2029, 2030]
# ----------------------------------------------------------------

# Named column constants for data/inputs.csv (never referenced by position)
COL_KEY = "key"
COL_LABEL = "label"
COL_VALUE = "value"
COL_UNIT = "unit"
COL_KIND = "kind"
COL_SOURCE_ID = "source_id"
COL_LOCATOR = "locator"


def _parse_value(raw):
    """A '|'-joined value becomes a list of floats (one per forecast year); else a float."""
    if "|" in raw:
        return [float(v) for v in raw.split("|")]
    try:
        return float(raw)
    except ValueError:
        return raw


class Inputs:
    """Dict-like accessor over data/inputs.csv, keyed by the `key` column."""

    def __init__(self, csv_path=INPUTS_CSV):
        self._rows = {}
        with open(csv_path, newline="", encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                self._rows[row[COL_KEY]] = row

    def value(self, key):
        return _parse_value(self._rows[key][COL_VALUE])

    def row(self, key):
        return self._rows[key]

    def keys(self):
        return self._rows.keys()

    def source_id(self, key):
        return self._rows[key][COL_SOURCE_ID]


INPUTS = Inputs()
