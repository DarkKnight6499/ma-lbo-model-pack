"""Recalculates excel/ma_lbo_model.xlsx with LibreOffice headless and diffs key outputs
against the independent Python model (model.py), which does not read the workbook.
Skips gracefully if soffice is not on this machine.
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import model  # noqa: E402
from config import EXCEL_OUTPUT  # noqa: E402

# ---------------------------- CONFIG ----------------------------
SOFFICE_CANDIDATES = [
    "soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
]
RELATIVE_TOLERANCE = 0.02  # the LBO circularity is solved differently (whole-sheet iterative
# calc in Excel vs. a per-year fixed point in Python), so allow a 2% band on derived returns
# ----------------------------------------------------------------


def find_soffice():
    for candidate in SOFFICE_CANDIDATES:
        path = shutil.which(candidate) or (candidate if Path(candidate).exists() else None)
        if path:
            return path
    return None


@pytest.fixture(scope="module")
def recalculated_workbook():
    soffice = find_soffice()
    if not soffice:
        pytest.skip("LibreOffice (soffice) not found on this machine; skipping recalculation check")
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(
            [soffice, "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, str(EXCEL_OUTPUT)],
            check=True, capture_output=True, timeout=120,
        )
        out_path = Path(tmp) / EXCEL_OUTPUT.name
        wb = openpyxl.load_workbook(out_path, data_only=True)
        yield wb


def test_workbook_has_no_formula_errors(recalculated_workbook):
    errors = []
    for sheet in recalculated_workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("#"):
                    errors.append((sheet.title, cell.coordinate, cell.value))
    assert errors == [], f"formula errors found: {errors}"


def test_workbook_master_check_is_true(recalculated_workbook):
    ws = recalculated_workbook["Checks"]
    values = [c.value for row in ws.iter_rows() for c in row if c.value in (True, False)]
    # the last boolean in the sheet is the MASTER CHECK row
    assert values[-1] is True


def test_workbook_goodwill_matches_python(recalculated_workbook):
    ws = recalculated_workbook["Purchase_Price_Allocation"]
    excel_goodwill = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == "Goodwill":
                excel_goodwill = row[1].value
    py_goodwill = model.build_results()["purchase_price_allocation"]["goodwill"]
    assert excel_goodwill is not None
    assert abs(excel_goodwill - py_goodwill) / py_goodwill < RELATIVE_TOLERANCE


def test_workbook_accretion_dilution_matches_python(recalculated_workbook):
    ws = recalculated_workbook["Accretion_Dilution"]
    excel_row = None
    for row in ws.iter_rows():
        if row[0].value == "Accretion / (dilution) %":
            excel_row = [c.value for c in row[1:6]]
    py_row = [r["accretion_dilution_pct"] for r in model.build_results()["accretion_dilution"]["by_year"]]
    assert excel_row is not None
    for excel_val, py_val in zip(excel_row, py_row):
        assert abs(excel_val - py_val) < 0.01, f"{excel_val} vs {py_val}"


def test_workbook_lbo_returns_match_python(recalculated_workbook):
    ws = recalculated_workbook["LBO"]
    moic = irr = None
    for row in ws.iter_rows():
        if row[0].value == "MOIC":
            moic = row[1].value
        if row[0].value and "IRR" in str(row[0].value):
            irr = row[1].value
    py_returns = model.build_results()["lbo"]["returns"]
    assert moic is not None and irr is not None
    assert abs(moic - py_returns["moic"]) / py_returns["moic"] < RELATIVE_TOLERANCE
    assert abs(irr - py_returns["irr"]) / py_returns["irr"] < RELATIVE_TOLERANCE
