#!/usr/bin/env python3
"""
Extract everything the standalone Egypt Distribution workbook needs from the
original CPAT workbook's cached values (pyxlsb) and from the already-built
data_tabular tables (Mapping, IO_GTAP, HHSurvey, HH_Elast, ASPIRE, WHOCooking,
GDPRatios).

Writes one pickle (egypt_data.pkl) containing:
  - "scalars": dict of {row: (label, description, unit, source, code, value)}
    for every populated row in Section B (rows 120-488) of the "Distribution"
    sheet, i.e. the module's own key-assumptions block, already Egypt-specific
    (Egypt is the currently-selected country in the source workbook).
  - "price_change_direct": {fuel_code: pct} for the 8 direct fuels (Step 1
    output, taken as given per user instruction -- these come from the
    Mitigation module's carbon-price pass-through, out of scope to re-derive
    here).
  - "price_change_indirect": {category_code: pct} for the 14 CPAT indirect
    consumption categories (Step 1.6 output, likewise taken as given).
  - "passthrough_table": the small, country-independent B.IV cost
    pass-through coefficient reference table (16 CPAT sectors).
  - "hhsurvey": (header, rows) filtered to iso3 == EGY
  - "hh_elast": (header, rows) filtered to iso3 == EGY
  - "aspire": (header, rows) filtered to ISO-3 Country Code == EGY
  - "whocooking": (header, rows) filtered to iso3 == EGY
  - "gdpratios": (header, rows) filtered to Country Code == EGY
  - "io_gtap": (header, rows) (already Egypt-only in the source)
  - "mapping_*": the five Mapping_* tables, unfiltered (country-independent
    crosswalks)
"""
import os
import pickle

from openpyxl import load_workbook
from pyxlsb import open_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.dirname(HERE)
CPAT_EXCEL_DIR = os.path.dirname(DIST_DIR)
SOURCE_XLSB = os.path.join(CPAT_EXCEL_DIR, "original", "CPAT 1.0pre_456_NoPropData.xlsb")
TABULAR_XLSX = os.path.join(DIST_DIR, "data_tabular", "CPAT_DistributionalData.xlsx")
OUT_PKL = os.path.join(HERE, "egypt_data.pkl")


def load_distribution_grid():
    with open_workbook(SOURCE_XLSB) as wb:
        with wb.get_sheet("Distribution") as sheet:
            rows = list(sheet.rows())
    ncols = max((len(r) for r in rows), default=0)
    grid = []
    for r in rows:
        rowvals = [None] * ncols
        for c in r:
            rowvals[c.c] = c.v
        grid.append(rowvals)
    return grid


def extract_section_b_scalars(grid):
    """Rows 120-488: label(col2), description(col3), unit(col5), source(col6),
    code(col7), value(col8). Only rows that carry a label in col2 and are not
    part of the large embedded pass-through year-trajectory sub-block
    (rows ~223-463, excluded -- see module docstring / build notes) are kept.
    Section/subsection header rows (col0 populated) are kept as markers with
    label=None so the generator can still place headings."""
    out = []
    for r in range(120, 489):
        row = grid[r]
        section = row[0]
        title = row[1]
        label = row[2]
        if section and title:
            out.append({"row": r, "kind": "header", "section": section, "title": title})
            continue
        if label is None:
            continue
        out.append({
            "row": r,
            "kind": "item",
            "label": label,
            "description": row[3],
            "unit": row[5],
            "source": row[6],
            "code": row[7],
            "value": row[8],
        })
    return out


FUEL_CODES = ["coa", "ely", "nga", "oil", "gso", "die", "ker", "lpg"]
CPAT_CATEGORIES = [
    "app", "che", "clo", "com", "edu", "food", "hea", "hou",
    "oth", "pap", "pha", "ret", "teq", "tpu",
]


def extract_price_changes(grid):
    # rows 770-777: C.II "Price Changes (%)" for the 8 fuels -- the price
    # response actually consumed by the direct-effect formula (Step 4), NOT
    # the raw retail "egy.dir.<fuel>" figure from C.I (which differs by the
    # C.I revenue-reconciliation/sector-matching adjustment). Stored as a
    # fraction (0.2545 = 25.45%), so x100 here.
    direct = {}
    code_alias = {"oop": "oil"}  # C.II spells non-road oil "oop", C.III/C.V use "oil"
    for r in range(770, 778):
        row = grid[r]
        code = code_alias.get(row[7], row[7])
        val = row[16]
        if code in FUEL_CODES and val is not None:
            direct[code] = val * 100
    assert set(direct) == set(FUEL_CODES), f"missing direct fuels: {set(FUEL_CODES) - set(direct)}"

    # rows 666-735ish: HH-demand-weighted average price increase by CPAT
    # category (C.I), already stored as a percent (no x100 needed) -- cross-
    # checked against the C.II row (786 etc.), which stores the same number
    # as a fraction, to 16 significant figures.
    indirect = {}
    for r in range(666, 736):
        row = grid[r]
        code = row[2]
        val = row[6]
        if code in CPAT_CATEGORIES and val is not None:
            indirect[code] = val
    assert set(indirect) == set(CPAT_CATEGORIES), f"missing categories: {set(CPAT_CATEGORIES) - set(indirect)}"
    return direct, indirect


def extract_elasticity_adjustment_factors(grid):
    """C.II rows 800-807 (fuels) and 811-824 (categories): a per-decile
    multiplier (col16=decile1 .. col25=decile10, and col16 doubles as the
    basket/national factor) applied on top of the price change before it
    hits the budget share in the direct/indirect effect formulas. Close to 1
    for categories, well below 1 for fuels -- NOT safe to assume away."""
    out = {}
    code_alias = {"oop": "oil"}
    for r in list(range(800, 808)) + list(range(811, 825)):
        row = grid[r]
        code = code_alias.get(row[7], row[7])
        if code not in FUEL_CODES and code not in CPAT_CATEGORIES:
            continue
        vals = [row[16 + i] for i in range(10)]
        if any(v is None for v in vals):
            continue
        out[code] = vals
    missing = (set(FUEL_CODES) | set(CPAT_CATEGORIES)) - set(out)
    assert not missing, f"missing elasticity adjustment factors: {missing}"
    return out


def extract_passthrough_table(grid):
    # rows 469-486: sector code(col1), sector name(col2), selected coefficient(col8)
    out = []
    for r in range(469, 486):
        row = grid[r]
        code, name, coef = row[1], row[2], row[8]
        if code and coef is not None:
            out.append((code, name, coef))
    return out


def read_tabular(name):
    wb = load_workbook(TABULAR_XLSX, read_only=True, data_only=True)
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    header = list(rows[0])
    data = [list(r) for r in rows[1:]]
    wb.close()
    return header, data


def filt(header, data, col, val):
    idx = header.index(col)
    return header, [r for r in data if r[idx] == val]


def main():
    print("Loading Distribution sheet grid from xlsb...")
    grid = load_distribution_grid()

    scalars = extract_section_b_scalars(grid)
    print(f"  {len(scalars)} Section B rows (headers + items)")

    direct, indirect = extract_price_changes(grid)
    print(f"  direct price changes: {direct}")
    print(f"  indirect price changes: {indirect}")

    elast_adj = extract_elasticity_adjustment_factors(grid)
    print(f"  elasticity adjustment factors: {len(elast_adj)} items")

    passthrough = extract_passthrough_table(grid)
    print(f"  passthrough table rows: {len(passthrough)}")

    out = {
        "scalars": scalars,
        "price_change_direct": direct,
        "price_change_indirect": indirect,
        "elasticity_adjustment": elast_adj,
        "passthrough_table": passthrough,
    }

    print("Loading + filtering data_tabular tables...")
    out["hhsurvey"] = filt(*read_tabular("HHSurvey"), "iso3", "EGY")
    out["hh_elast"] = filt(*read_tabular("HH_Elast"), "iso3", "EGY")
    out["aspire"] = filt(*read_tabular("ASPIRE"), "ISO-3 Country Code", "EGY")
    out["whocooking"] = filt(*read_tabular("WHOCooking"), "iso3", "EGY")
    out["gdpratios"] = filt(*read_tabular("GDPRatios"), "Country Code", "EGY")
    out["io_gtap"] = read_tabular("IO_GTAP")
    for m in [
        "Mapping_SectorCrosswalk", "Mapping_CPATSectorsToISIC",
        "Mapping_IEAFlowsToISIC", "Mapping_ISICToCPAT", "Mapping_CountriesToGTAP10",
    ]:
        out[m.lower()] = read_tabular(m)

    for k in ["hhsurvey", "hh_elast", "aspire", "whocooking", "gdpratios", "io_gtap"]:
        h, d = out[k]
        print(f"  {k}: {len(d)} rows x {len(h)} cols")

    with open(OUT_PKL, "wb") as f:
        pickle.dump(out, f)
    print(f"Wrote {OUT_PKL}")


if __name__ == "__main__":
    main()
