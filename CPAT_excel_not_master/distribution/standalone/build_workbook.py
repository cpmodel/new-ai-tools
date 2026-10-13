#!/usr/bin/env python3
"""
Build the standalone CPAT Distribution module workbook for Egypt.

Structure (mirrors the conventions of CPAT_1.0pre_456_NoPropData.xlsb, but
scoped to a single country and split into an Inputs / Outputs pair instead of
one combined "Distribution" sheet):

  ReadMe               - what this workbook is, and the scope decisions made
  Distribution_Inputs  - Section A (trimmed) + Section B of the original
                          "Distribution" sheet: run configuration, policy
                          switches, price-change inputs (taken as given from
                          the CPAT Mitigation module), population/GDP/
                          consumption rebasing inputs (also Mitigation-
                          sourced). Same row labels/placement as the source.
  Distribution_Outputs - Sections C-F of the original sheet: the actual
                          distributional calculation (budget shares, direct/
                          indirect effects, rebasing, revenue recycling,
                          Gini/Lorenz, horizontal equity, compensation,
                          chart-ready outputs). Computed with LAMBDA-based
                          formulas defined once in the Name Manager and
                          reused across every row, instead of the source
                          workbook's per-cell bespoke formulas.
  Mapping, IO_GTAP, HHSurvey, HH_Elast, ASPIRE, WHOCooking, GDPRatios
                        - the DATA_DISTN data tables, filtered to Egypt only
                          (Mapping crosswalks are country-independent and
                          kept in full).

Run: python3 build_workbook.py
"""
import pickle
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
PKL = os.path.join(HERE, "egypt_data.pkl")
OUT = os.path.join(HERE, "CPAT_Distribution_Standalone_Egypt.xlsx")

with open(PKL, "rb") as f:
    DATA = pickle.load(f)

# ---------------------------------------------------------------------------
# Styling (Arial 10, consistent with the rest of the cpat_excel pipeline)
# ---------------------------------------------------------------------------
BASE_FONT = Font(name="Arial", size=10)
BOLD_FONT = Font(name="Arial", size=10, bold=True)
ITALIC_FONT = Font(name="Arial", size=10, italic=True, color="595959")
TITLE_FONT = Font(name="Arial", size=14, bold=True, color="1F4E78")
SECTION_FONT = Font(name="Arial", size=11, bold=True, color="FFFFFF")
SUBSECTION_FONT = Font(name="Arial", size=10, bold=True, color="1F4E78")
SECTION_FILL = PatternFill("solid", fgColor="1F4E78")
SUBSECTION_FILL = PatternFill("solid", fgColor="D9E2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()
wb.remove(wb.active)


def style_cell(ws, r, c, value=None, font=BASE_FONT, fill=None, align=None, numfmt=None, border=False):
    cell = ws.cell(row=r, column=c)
    if value is not None:
        cell.value = value
    cell.font = font
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = align
    if numfmt:
        cell.number_format = numfmt
    if border:
        cell.border = BOX
    return cell


TABLE_NAMES = []  # every Excel Table displayName created, for the collision check below


def write_table(ws, top_row, top_col, header, data, table_name, style="TableStyleMedium2"):
    """Write header+data starting at (top_row, top_col) and register as an
    Excel Table (ListObject) so formulas elsewhere can use structured
    references like TableName[Column]."""
    TABLE_NAMES.append(table_name)
    ncols = len(header)
    seen = {}
    dedup_header = []
    for h in header:
        h = str(h)
        n = seen.get(h, 0)
        seen[h] = n + 1
        dedup_header.append(h if n == 0 else f"{h}_{n + 1}")
    header = dedup_header
    for c, h in enumerate(header):
        style_cell(ws, top_row, top_col + c, h, BOLD_FONT, fill=SUBSECTION_FILL, border=True)
    for r, row in enumerate(data, start=top_row + 1):
        for c, v in enumerate(row):
            style_cell(ws, r, top_col + c, v, BASE_FONT, border=True)
    last_row = top_row + len(data)
    last_col = top_col + ncols - 1
    ref = f"{get_column_letter(top_col)}{top_row}:{get_column_letter(last_col)}{last_row}"
    tbl = Table(displayName=table_name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name=style, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(tbl)
    for c in range(top_col, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 15
    ws.freeze_panes = ws.cell(row=top_row + 1, column=top_col).coordinate
    return ref


# ---------------------------------------------------------------------------
# Fuel / category code -> HHSurvey/HH_Elast column name maps
# ---------------------------------------------------------------------------
FUEL_CODES = ["coa", "ely", "nga", "oil", "gso", "die", "ker", "lpg"]          # taxed/priced fuels
FUEL_CODES_ALL = FUEL_CODES + ["ccl", "ethanol", "fwd"]                        # + traditional biomass
FUEL_LABEL = {"coa": "Coal", "ely": "Electricity", "nga": "Natural Gas", "oil": "Non-Road Oil",
              "gso": "Gasoline", "die": "Diesel", "ker": "Kerosene", "lpg": "LPG",
              "ccl": "Charcoal", "ethanol": "Ethanol", "fwd": "Firewood"}
FUEL_SHARE_COL = {c: f"{c}_share" for c in FUEL_CODES_ALL}
FUEL_ELAST_COL = {c: f"{c}_elasticity" for c in FUEL_CODES}

CAT_CODES = ["app", "che", "clo", "com", "edu", "food", "hea", "hou", "oth", "pap", "pha", "ret", "teq", "tpu"]
CAT_LABEL = {"app": "Appliances", "che": "Chemicals", "clo": "Clothing", "com": "Communications",
             "edu": "Education", "food": "Food", "hea": "Health Services", "hou": "Housing",
             "oth": "Other", "pap": "Paper", "pha": "Pharmaceuticals", "ret": "Recreation/Tourism",
             "teq": "Transportation Equipment", "tpu": "Public Transportation"}
CAT_SHARE_COL = {"app": "appliances_share", "che": "chemicals_share", "clo": "clothing_share",
                  "com": "communications_share", "edu": "education_share", "food": "food_share",
                  "hea": "health_srv_share", "hou": "housing_share", "oth": "other_share",
                  "pap": "paper_share", "pha": "pharma_share", "ret": "rectourism_share",
                  "teq": "transp_eqt_share", "tpu": "transp_pub_share"}
CAT_ELAST_COL = {"app": "appliances_elasticity", "che": "chemicals_elasticity", "clo": "clothing_elasticity",
                  "com": "communications_elasticity", "edu": "education_elasticity", "food": "food_elasticity",
                  "hea": "health_srv_elasticity", "hou": "housing_elasticity", "oth": "other_elasticity",
                  "pap": "paper_elasticity", "pha": "pharma_elasticity", "ret": "rectourism_elasticity",
                  "teq": "transp_eqt_elasticity", "tpu": "transp_pub_elasticity"}

ASPIRE_PROGRAM_LABEL = {
    "allsp": "All Social Protection and Labor", "sa_allsa": "All Social Assistance",
    "sa_ct": "Cash Transfers", "sa_ik": "In-Kind", "si_allsi": "All Social Insurance",
    "si_cp": "Contributory Pensions",
}
STATS = ["mean", "median", "p25", "p75"]
SAMPLES = ["Overall", "Urban", "Rural"]
SAMPLE_CODE = {"Overall": "ove", "Urban": "urb", "Rural": "rur"}


def safe_header(h):
    h = str(h)
    if h and h[0].isdigit():
        h = "Y" + h
    return h.replace(" ", "_").replace("(", "").replace(")", "").replace("%", "pct").replace("-", "_").replace(".", "_").replace(",", "").replace("/", "_").replace("'", "")


# ---------------------------------------------------------------------------
# Data tabs (Egypt-filtered DATA_DISTN tables)
# ---------------------------------------------------------------------------
def add_key_column(header, data):
    """Append a lookup key 'sample|type|stat_type|quant_cons' column, used by
    the BSHARE/ELAST LAMBDA functions for a single XLOOKUP instead of a
    multi-criteria array match."""
    si, ti, sti, qi = header.index("sample"), header.index("type"), header.index("stat_type"), header.index("quant_cons")
    header2 = list(header) + ["key"]
    data2 = []
    for row in data:
        row = list(row)
        key = f"{row[si]}|{row[ti]}|{row[sti]}|{int(row[qi])}"
        row.append(key)
        data2.append(row)
    return header2, data2


NAMED_RANGES = {}  # name -> formula string


def define_name(name, ref_formula):
    dn = DefinedName(name, attr_text=ref_formula)
    wb.defined_names[name] = dn
    NAMED_RANGES[name] = ref_formula


def build_data_tabs():
    hh_header, hh_data = add_key_column(*DATA["hhsurvey"])
    el_header, el_data = add_key_column(*DATA["hh_elast"])
    gdp_header, gdp_data_full = DATA["gdpratios"]
    ic_idx = gdp_header.index("Indicator Code")
    gdp_data = [r for r in gdp_data_full if r[ic_idx] == "NE.CON.PRVT.ZS"]
    specs = [
        ("HHSurvey", (hh_header, hh_data)),
        ("HH_Elast", (el_header, el_data)),
        ("ASPIRE", DATA["aspire"]),
        ("WHOCooking", DATA["whocooking"]),
        ("GDPRatios", (gdp_header, gdp_data)),
        ("IO_GTAP", DATA["io_gtap"]),
    ]
    for name, (header, data) in specs:
        ws = wb.create_sheet(name)
        header = [safe_header(h) for h in header]
        write_table(ws, 1, 1, header, data, name)

    # Named ranges for HHSurvey / HH_Elast: header row, full data body, key column
    for name, (header, data) in [("HHSurvey", (hh_header, hh_data)), ("HH_Elast", (el_header, el_data))]:
        ncols = len(header)
        nrows = len(data)
        last_col = get_column_letter(ncols)
        key_col = get_column_letter(ncols)  # key is the last column
        define_name(f"{name}_Headers", f"{name}!$A$1:${last_col}$1")
        define_name(f"{name}_Data", f"{name}!$A$2:${last_col}${1 + nrows}")
        define_name(f"{name}_Key", f"{name}!${key_col}$2:${key_col}${1 + nrows}")

    # ASPIRE pivot: program code (rows) x quintile Q1..Q5 (cols), built from
    # the raw Series_Code rows (EGY.per_<program>.avt_q<N>_preT_tot)
    header, data = DATA["aspire"]
    idx_code = header.index("Series_Code")
    idx_val = header.index("(firstnm) Value")
    pivot = {}
    for row in data:
        code = row[idx_code]
        # e.g. EGY.per_allsp.avt_q1_preT_tot -> program='allsp', quintile=1
        body = code.split(".", 2)
        program = body[1][len("per_"):]
        qpart = body[2]
        q = int(qpart.split("_")[1][1:])
        pivot.setdefault(program, {})[q] = row[idx_val]
    ws = wb["ASPIRE"]
    top = 1 + len(data) + 3
    style_cell(ws, top, 1, "ASPIRE incidence pivot (avg. per-capita transfer, preT, US$ PPP/day, by quintile)", SUBSECTION_FONT)
    top += 1
    prog_header = ["program", "label", "Q1", "Q2", "Q3", "Q4", "Q5"]
    prog_rows = []
    for prog, qs in sorted(pivot.items()):
        prog_rows.append([prog, ASPIRE_PROGRAM_LABEL.get(prog, prog)] + [qs.get(q) for q in range(1, 6)])
    write_table(ws, top, 1, prog_header, prog_rows, "ASPIRE_Pivot")
    define_name("ASPIRE_Programs", f"ASPIRE!$A${top + 1}:$A${top + len(prog_rows)}")
    define_name("ASPIRE_Quintiles", f"ASPIRE!$C${top}:$G${top}")
    define_name("ASPIRE_Values", f"ASPIRE!$C${top + 1}:$G${top + len(prog_rows)}")

    # Price-change reference tables (Step 1 outputs, taken as given inputs --
    # see Distribution_Inputs A.II scope note)
    ws = wb.create_sheet("Price_Changes")
    ws.sheet_properties.tabColor = "FFC000"
    style_cell(ws, 1, 1, "Direct price changes by fuel (%, given -- see Distribution_Inputs)", SUBSECTION_FONT)
    direct = DATA["price_change_direct"]
    rows = [[c, FUEL_LABEL[c], v] for c, v in direct.items()]
    write_table(ws, 2, 1, ["code", "fuel", "pct_change"], rows, "PriceDirect")
    define_name("PriceDirect_Codes", f"Price_Changes!$A$3:$A${2 + len(rows)}")
    define_name("PriceDirect_Values", f"Price_Changes!$C$3:$C${2 + len(rows)}")

    r2 = 2 + len(rows) + 3
    style_cell(ws, r2, 1, "Indirect (HH-demand-weighted) price changes by CPAT category (%, given)", SUBSECTION_FONT)
    indirect = DATA["price_change_indirect"]
    rows2 = [[c, CAT_LABEL[c], v] for c, v in indirect.items()]
    write_table(ws, r2 + 1, 1, ["code", "category", "pct_change"], rows2, "PriceIndirect")
    define_name("PriceIndirect_Codes", f"Price_Changes!$A${r2 + 2}:$A${r2 + 1 + len(rows2)}")
    define_name("PriceIndirect_Values", f"Price_Changes!$C${r2 + 2}:$C${r2 + 1 + len(rows2)}")

    define_name("DECILE_ARRAY", "{1,2,3,4,5,6,7,8,9,10}")

    # Elasticity adjustment factors (C.II rows 800-824 of the source): a
    # per-decile multiplier on the given price change, applied before the
    # budget-share multiplication in DIRECT_EFFECT/INDIRECT_EFFECT. Column
    # 16 in the source doubles as both the basket AND decile-1 factor, so
    # column "D1" here is used for both (decile=0 maps to column D1).
    r3 = r2 + 1 + len(rows2) + 3
    style_cell(ws, r3, 1,
                "Price-elasticity adjustment factors (C.II), by decile -- multiplies the price "
                "change above before it is applied to budget share in DIRECT_EFFECT/INDIRECT_EFFECT",
                SUBSECTION_FONT)
    ea = DATA["elasticity_adjustment"]
    ea_header = ["code"] + [f"D{d}" for d in range(1, 11)]
    ea_rows = [[code] + vals for code, vals in ea.items()]
    # NOTE: table displayName must not collide (case-insensitively -- Excel
    # names are case-insensitive) with the ELASTADJ LAMBDA function name
    # defined later, or Names.Add for ELASTADJ silently fails (swallowed by
    # AddDistributionNames' On Error Resume Next) and every formula that
    # calls it shows #REF!. Hence "ElastAdjTbl", not "ElastAdj".
    write_table(ws, r3 + 1, 1, ea_header, ea_rows, "ElastAdjTbl")
    define_name("ElastAdj_Codes", f"Price_Changes!$A${r3 + 2}:$A${r3 + 1 + len(ea_rows)}")
    define_name("ElastAdj_Data", f"Price_Changes!$B${r3 + 2}:$K${r3 + 1 + len(ea_rows)}")

    # Mapping sheet: five sub-tables side by side / stacked, country-independent
    ws = wb.create_sheet("Mapping")
    row = 1
    for key, label in [
        ("mapping_sectorcrosswalk", "GTAP -> CPAT sector crosswalk"),
        ("mapping_cpatsectorstoisic", "CPAT sectors to ISIC"),
        ("mapping_ieaflowstoisic", "IEA flows to ISIC"),
        ("mapping_isictocpat", "ISIC to CPAT"),
        ("mapping_countriestogtap10", "Countries to GTAP-10"),
    ]:
        header, data = DATA[key]
        style_cell(ws, row, 1, label, SUBSECTION_FONT)
        row += 1
        header = [safe_header(h) for h in header]
        write_table(ws, row, 1, header, data, key.replace("mapping_", "Map_"))
        row += len(data) + 3


build_data_tabs()
print("Data tabs written:", wb.sheetnames)

# ---------------------------------------------------------------------------
# Distribution_Inputs
# ---------------------------------------------------------------------------
COL = {"code": 1, "title": 2, "label": 3, "desc": 4, "unit": 6, "source": 7, "var": 8, "value": 9}
OVERVIEW_TEXT = (
    "The Distribution module estimates the household-level distributional (incidence) impact of a "
    "carbon-pricing mitigation policy scenario for a single selected country. Starting from the price "
    "changes produced by the Mitigation module, it estimates how much more each household decile pays "
    "for energy directly (fuels it buys itself) and indirectly (through the higher price of goods and "
    "services embedding those fuels), nets that burden against any revenue recycled back to households, "
    "and reports the results as changes in household consumption, inequality (Gini/Lorenz) and the share "
    "of carbon-pricing revenue needed to compensate lower-income deciles."
)
SCOPE_NOTE = (
    "This standalone workbook implements sections B-F of the CPAT \"Distribution\" sheet (see "
    "cpat_excel/original/CPAT 1.0pre_456_NoPropData.xlsb) for Egypt only, replacing the original's "
    "per-cell formulas with a small set of reusable LAMBDA functions (Name Manager) that read directly "
    "from the Egypt-filtered HHSurvey / HH_Elast / ASPIRE / WHOCooking / GDPRatios / IO_GTAP data tabs. "
    "Fuel-level direct price changes and CPAT-category indirect price changes (Step 1 of the module) are "
    "taken as given inputs, exactly as currently cached in the source workbook's carbon-price scenario "
    "for Egypt -- this mirrors the module's own architecture (Distribution is a consumer, not a producer, "
    "of Mitigation's price changes) and avoids re-deriving pieces (IEA/GAINS/IMF emissions recalibration, "
    "the Mitigation module's own USD-denominated revenue reconciliation) that depend on data outside the "
    "DATA_DISTN tables. Population, GDP and GDP-deflator inputs are likewise Mitigation-module macro "
    "outputs and are taken as given. Everything downstream of those inputs -- budget shares, direct/"
    "indirect effects, survey-to-national-accounts rebasing, revenue recycling, net effects, Gini/Lorenz, "
    "horizontal equity and compensation shares -- is computed with live formulas against the data tabs."
)


def new_output_sheet(name, tab_color=None):
    ws = wb.create_sheet(name)
    if tab_color:
        ws.sheet_properties.tabColor = tab_color
    return ws


def write_header_row(ws, r=1):
    labels = ["Checks", "Country Name", "Country Code", "Statistic", "Sample", "Scenario",
              "Item", "Description", "Unit", "Source", "Var. Code", "", "", "", "Note"]
    for i, lab in enumerate(labels, start=1):
        style_cell(ws, r, i, lab, BOLD_FONT, fill=SUBSECTION_FILL)


def write_section_header(ws, r, code, title):
    style_cell(ws, r, COL["code"], code, SECTION_FONT, fill=SECTION_FILL)
    c2 = style_cell(ws, r, COL["title"], title, SECTION_FONT, fill=SECTION_FILL)
    ws.merge_cells(start_row=r, start_column=COL["title"], end_row=r, end_column=14)
    for col in range(COL["title"] + 1, 15):
        style_cell(ws, r, col, fill=SECTION_FILL)


def build_distribution_inputs():
    ws = new_output_sheet("Distribution_Inputs", tab_color="FFC000")
    ws.sheet_view.showGridLines = False
    r = 1
    style_cell(ws, r, 1, "Distributional Effects Module - Egypt (Standalone)", TITLE_FONT)
    r += 2
    write_section_header(ws, r, "A.", "Module overview")
    r += 2
    style_cell(ws, r, COL["code"], "A.I.", SUBSECTION_FONT)
    style_cell(ws, r, COL["title"], "Description", SUBSECTION_FONT)
    r += 1
    c = style_cell(ws, r, COL["label"], OVERVIEW_TEXT, BASE_FONT, align=Alignment(wrap_text=True, vertical="top"))
    ws.merge_cells(start_row=r, start_column=COL["label"], end_row=r + 3, end_column=14)
    ws.row_dimensions[r].height = 90
    r += 5
    style_cell(ws, r, COL["code"], "A.II.", SUBSECTION_FONT)
    style_cell(ws, r, COL["title"], "Scope of this standalone workbook", SUBSECTION_FONT)
    r += 1
    style_cell(ws, r, COL["label"], SCOPE_NOTE, BASE_FONT, align=Alignment(wrap_text=True, vertical="top"))
    ws.merge_cells(start_row=r, start_column=COL["label"], end_row=r + 7, end_column=14)
    ws.row_dimensions[r].height = 170
    r += 9
    style_cell(ws, r, COL["code"], "A.III.", SUBSECTION_FONT)
    style_cell(ws, r, COL["title"], "Country coverage", SUBSECTION_FONT)
    r += 1
    style_cell(ws, r, COL["label"], "This workbook is built for a single country:", BASE_FONT)
    r += 1
    for lab, val in [("Country Name", "Egypt"), ("ISO-3 Country Code", "EGY"),
                      ("Reference Year (HH Survey)", 2017), ("Survey Acronym", "HBS"),
                      ("Survey Name", "Household Budget Survey")]:
        style_cell(ws, r, COL["label"], lab, BASE_FONT)
        style_cell(ws, r, COL["desc"], val, BASE_FONT)
        r += 1
    r += 1

    write_section_header(ws, r, "B.", "Key assumptions and inputs")
    r += 2
    style_cell(ws, r, COL["code"], "B.I.", SUBSECTION_FONT)
    style_cell(ws, r, COL["title"], "Key assumptions", SUBSECTION_FONT)
    r += 2

    input_row_by_label = {}
    last_group = None
    for item in DATA["scalars"]:
        if item["kind"] == "header":
            r += 1
            style_cell(ws, r, COL["code"], item["section"], SUBSECTION_FONT)
            style_cell(ws, r, COL["title"], item["title"], SUBSECTION_FONT)
            r += 1
            continue
        label = item["label"]
        # group sub-heading rows in the source (col2 populated, everything else blank)
        is_group_heading = all(item[k] is None for k in ("description", "unit", "source", "code", "value"))
        style_cell(ws, r, COL["label"], label, BOLD_FONT if is_group_heading else BASE_FONT)
        if not is_group_heading:
            style_cell(ws, r, COL["desc"], item["description"], ITALIC_FONT)
            style_cell(ws, r, COL["unit"], item["unit"], BASE_FONT)
            style_cell(ws, r, COL["source"], item["source"], BASE_FONT)
            style_cell(ws, r, COL["var"], item["code"], BASE_FONT)
            style_cell(ws, r, COL["value"], item["value"], BASE_FONT, fill=INPUT_FILL, border=True)
            input_row_by_label.setdefault(label, []).append((r, item["unit"]))
        r += 1

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 46
    ws.column_dimensions["D"].width = 48
    ws.column_dimensions["E"].width = 3
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 12
    ws.column_dimensions["H"].width = 26
    ws.column_dimensions["I"].width = 16
    ws.freeze_panes = "C1"
    return r, input_row_by_label


last_input_row, INPUT_ROW = build_distribution_inputs()
print("Distribution_Inputs written, last row", last_input_row)


def find_row(label, unit_contains=None):
    """INPUT_ROW[label] is a list of (row, unit) -- the source sheet repeats
    some labels 2-3x (%, %GDP, LCU versions of the same figure). Pick the
    first entry whose unit contains unit_contains, else the first entry."""
    entries = INPUT_ROW[label]
    if unit_contains:
        for r, u in entries:
            if u and unit_contains in u:
                return r
    return entries[0][0]


def named(label, unit_contains=None):
    r = find_row(label, unit_contains)
    return f"Distribution_Inputs!$I${r}"


SWITCH_LABELS = {
    "AdjustBehaviorSwitch": ("Adjust for behavioral & structural change?", None),
    "BehaviorAdjFactor": ("Behavioral Response Adj. Factor", None),
    "AdjustDWLSwitch": ("Adjust for deadweight losses?", None),
    "UseDecileElastSwitch": ("Decile-specific price elasticities of demand?", None),
    "ExemptSwitch": ("Exempt most-used cooking fossil fuel?", None),
    "ExemptFuelLabel": ("Cooking fossil fuel exempt (if applicable)", None),
    "ExemptShare": ("Share of exempt cooking fossil fuel", None),
    "ExemptDeciles": ("Exempt cooking fossil fuel for bottom XX deciles", None),
    "LaborTaxMethod": ("Labor Tax Reduction Method", None),
    "LaborTaxCutCoef": ('If Method is "Proportional Compensation", Labor Tax Cut Coefficient', None),
    "PITSourceSwitch": ("Replace Missing PIT Data Based on:", None),
    "PITBaselineLCU": ("Baseline PIT Revenues - Analysis Year", "LCU"),
    "ExemptBottomDeciles_PIT": ('If "Labor Tax Reduction Method" is "Targeted Exemption", exempt labor taxes for bottom XX deciles:', None),
    "StatOutputs": ("Statistic - Outputs", None),
    "PopAnalysisYear": ("Population - Analysis Year", None),
    "PopSurveyYear": ("Population - Survey Year", None),
    "PopAdjFactor": ("Population Adj. Factor", None),
    "HHConsAdjFactor": ("Household Consumption Adj. Factor", None),
    "TotConsSurveyYear": ("Total Household Consumption - Survey Year", None),
    "GDPDeflatorRatioSvy": ("GDP Deflator Ratio (HH Survey)", None),
    "RevLaborTaxLCU": ("CP Revenue - Labor Tax Reductions - Analysis Year", "LCU"),
    "RevTargetedTransferLCU": ("CP Revenue - Targeted Transfer - Analysis Year", "LCU"),
    "RevPublicInvestLCU": ("CP Revenue - Public Investment - Analysis Year", "LCU"),
    "RevCurrentSpendLCU": ("CP Revenue - Current Spending - Analysis Year", "LCU"),
    "TargetedTransferASPIREUsed": ("Using ASPIRE for targeted transfer?", None),
    "TargetedTransferCode": ("ASPIRE targeted transfer code", None),
    "CurrentSpendCode": ("ASPIRE current spending code", None),
    "InfraAccessCode": ("Infrastructure access data code", None),
}
for name, (label, unit) in SWITCH_LABELS.items():
    if label in INPUT_ROW:
        define_name(name, named(label, unit))
    else:
        print(f"  ! label not found for named range {name}: {label!r}")

# ---------------------------------------------------------------------------
# Upgrade a handful of Section B "given" scalars to real formulas: these are
# simple ratios/lookups that are fully reproducible from other Distribution_
# Inputs cells or from the Egypt-filtered data tabs, so there is no reason to
# hard-code them as static pasted-in values the way the source workbook does.
# ---------------------------------------------------------------------------
ws_in = wb["Distribution_Inputs"]


def set_formula(label, formula, unit=None):
    r = find_row(label, unit)
    cell = ws_in.cell(row=r, column=COL["value"])
    cell.value = formula
    cell.fill = CALC_FILL


R = find_row  # shorthand for building cross-references within Distribution_Inputs

set_formula("GDP Deflator Ratio",
            f"=$I${R('2026 GDP Deflator')}/$I${R('2011 GDP Deflator')}")
set_formula("Population Adj. Factor",
            f"=$I${R('Population - Analysis Year')}/$I${R('Population - Survey Year')}")
set_formula("Per-Capita Real GDP - Analysis Year",
            f"=$I${R('Real GDP - Analysis Year')}/$I${R('Population - Analysis Year')}")
set_formula("GDP Deflator Ratio (HH Survey)",
            f"=$I${R('2026 GDP Deflator', 'Mitigation Module')}/$I${R('2017 GDP Deflator')}")
set_formula("GDP Deflator Ratio (GTAP)",
            f"=$I${R('2026 GDP Deflator', 'Mitigation Module')}/$I${R('2014 GDP Deflator')}")
set_formula("Population - Survey Year",
            '=XLOOKUP("Overall|Basket|mean|9999",HHSurvey_Key,'
            'INDEX(HHSurvey_Data,0,MATCH("popw",HHSurvey_Headers,0)))')
set_formula("National Accounts Household Consumption-to-GDP",
            '=XLOOKUP("Y"&$I$' + str(R("Latest Year of Consumption/GDP Data")) +
            ',GDPRatios!$F$1:$BQ$1,GDPRatios!$F$2:$BQ$2)')
set_formula("National Accounts Household Consumption - Analysis Year",
            f"=$I${R('Real GDP - Analysis Year')}*$I${R('National Accounts Household Consumption-to-GDP')}/100")
set_formula("Household Consumption Adj. Factor",
            f"=$I${R('National Accounts Household Consumption - Analysis Year')}"
            f"/$I${R('Total Household Consumption - Survey Year')}")
print("Section B formula upgrades applied.")

# ---------------------------------------------------------------------------
# LAMBDA library (Name Manager) -- the reusable calculation engine. Every
# formula in Distribution_Outputs is a call into one of these, instead of
# the source workbook's one bespoke formula per cell.
# ---------------------------------------------------------------------------
LAMBDAS = {
    "HHKEY": (
        "LAMBDA(sample,qtype,stat,dnum, sample&\"|\"&qtype&\"|\"&stat&\"|\"&dnum)"
    ),
    "BSHARE": (
        "LAMBDA(colname,decile,sample,stat,"
        "LET(qtype,IF(decile=0,\"Basket\",\"Deciles\"),"
        "dnum,IF(decile=0,9999,decile),"
        "key,HHKEY(sample,qtype,stat,dnum),"
        "col,MATCH(colname,HHSurvey_Headers,0),"
        "XLOOKUP(key,HHSurvey_Key,INDEX(HHSurvey_Data,0,col))))"
    ),
    "ELAST": (
        "LAMBDA(colname,decile,sample,stat,"
        "LET(qtype,IF(decile=0,\"Basket\",\"Deciles\"),"
        "dnum,IF(decile=0,9999,decile),"
        "key,HHKEY(sample,qtype,stat,dnum),"
        "col,MATCH(colname,HH_Elast_Headers,0),"
        "XLOOKUP(key,HH_Elast_Key,INDEX(HH_Elast_Data,0,col))))"
    ),
    "PCHANGE_DIRECT": "LAMBDA(fuel,XLOOKUP(fuel,PriceDirect_Codes,PriceDirect_Values))",
    "PCHANGE_INDIRECT": "LAMBDA(cat,XLOOKUP(cat,PriceIndirect_Codes,PriceIndirect_Values))",
    # INDEX(range,0,{array of column numbers}) is NOT a supported way to
    # pull several columns at once (0 = "whole column/row" only behaves
    # predictably when the OTHER index is a scalar) -- it's what produced
    # the #VALUE! error when called with the full DECILE_ARRAY. Instead:
    # find the item's row with MATCH (item is always a scalar, so this is
    # a scalar row number), pull that whole row with INDEX(range,row,0)
    # (again scalar row, so this direction is well-defined), then use
    # CHOOSECOLS -- which *is* designed to take an array of column
    # positions -- to pick out the (possibly array-valued) decile(s).
    "ELASTADJ": (
        "LAMBDA(item,decile,"
        "LET(r,MATCH(item,ElastAdj_Codes,0),"
        "row,INDEX(ElastAdj_Data,r,0),"
        "CHOOSECOLS(row,IF(decile=0,1,decile))))"
    ),
    "BEHAVIOR_ADJ": "LAMBDA(IF(AdjustBehaviorSwitch=\"Yes\",BehaviorAdjFactor,1))",
    "DWL": (
        "LAMBDA(sharecol,elastcol,decile,sample,stat,pricepct,"
        "IF(AdjustDWLSwitch=\"Yes\","
        "0.5*ELAST(elastcol,decile,sample,stat)*(pricepct/100)^2*(BSHARE(sharecol,decile,sample,stat)/100)*100,"
        "0))"
    ),
    "DIRECT_EFFECT": (
        "LAMBDA(fuel,fuellabel,sharecol,elastcol,decile,sample,stat,"
        "LET(price,PCHANGE_DIRECT(fuel)*ELASTADJ(fuel,decile),"
        "bs,BSHARE(sharecol,decile,sample,stat),"
        "raw,(bs/100)*(price/100)*BEHAVIOR_ADJ()*100-DWL(sharecol,elastcol,decile,sample,stat,price),"
        "exempt,IF(AND(ExemptSwitch=\"Yes\",fuellabel=ExemptFuelLabel,decile<=ExemptDeciles,decile>0),1-ExemptShare,1),"
        "raw*exempt))"
    ),
    "INDIRECT_EFFECT": (
        "LAMBDA(cat,sharecol,elastcol,decile,sample,stat,"
        "LET(price,PCHANGE_INDIRECT(cat)*ELASTADJ(cat,decile),"
        "bs,BSHARE(sharecol,decile,sample,stat),"
        "(bs/100)*(price/100)*BEHAVIOR_ADJ()*100-DWL(sharecol,elastcol,decile,sample,stat,price)))"
    ),
    "TOTAL_DIRECT_EFFECT": (
        "LAMBDA(decile,sample,stat,"
        "DIRECT_EFFECT(\"coa\",\"Coal\",\"coa_share\",\"coa_elasticity\",decile,sample,stat)"
        "+DIRECT_EFFECT(\"ely\",\"Electricity\",\"ely_share\",\"ely_elasticity\",decile,sample,stat)"
        "+DIRECT_EFFECT(\"nga\",\"Natural Gas\",\"nga_share\",\"nga_elasticity\",decile,sample,stat)"
        "+DIRECT_EFFECT(\"oil\",\"Non-Road Oil\",\"oil_share\",\"oil_elasticity\",decile,sample,stat)"
        "+DIRECT_EFFECT(\"gso\",\"Gasoline\",\"gso_share\",\"gso_elasticity\",decile,sample,stat)"
        "+DIRECT_EFFECT(\"die\",\"Diesel\",\"die_share\",\"die_elasticity\",decile,sample,stat)"
        "+DIRECT_EFFECT(\"ker\",\"Kerosene\",\"ker_share\",\"ker_elasticity\",decile,sample,stat)"
        "+DIRECT_EFFECT(\"lpg\",\"LPG\",\"lpg_share\",\"lpg_elasticity\",decile,sample,stat))"
    ),
    "TOTAL_INDIRECT_EFFECT": (
        "LAMBDA(decile,sample,stat,"
        "INDIRECT_EFFECT(\"app\",\"appliances_share\",\"appliances_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"che\",\"chemicals_share\",\"chemicals_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"clo\",\"clothing_share\",\"clothing_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"com\",\"communications_share\",\"communications_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"edu\",\"education_share\",\"education_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"food\",\"food_share\",\"food_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"hea\",\"health_srv_share\",\"health_srv_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"hou\",\"housing_share\",\"housing_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"oth\",\"other_share\",\"other_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"pap\",\"paper_share\",\"paper_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"pha\",\"pharma_share\",\"pharma_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"ret\",\"rectourism_share\",\"rectourism_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"teq\",\"transp_eqt_share\",\"transp_eqt_elasticity\",decile,sample,stat)"
        "+INDIRECT_EFFECT(\"tpu\",\"transp_pub_share\",\"transp_pub_elasticity\",decile,sample,stat))"
    ),
    "TOTAL_EFFECT": (
        "LAMBDA(decile,sample,stat,"
        "TOTAL_DIRECT_EFFECT(decile,sample,stat)+TOTAL_INDIRECT_EFFECT(decile,sample,stat))"
    ),

    # --- Step 7: survey-to-national-accounts rebasing --------------------
    "ADJ_POP": 'LAMBDA(decile,sample,BSHARE("popw",decile,sample,"mean")*PopAdjFactor)',
    "ADJ_CONS_PC": (
        'LAMBDA(decile,sample,'
        'BSHARE("cons_pc_acrent",decile,sample,"mean")*GDPDeflatorRatioSvy*HHConsAdjFactor)'
    ),
    "ADJ_CONS_TOT": "LAMBDA(decile,sample,ADJ_CONS_PC(decile,sample)*ADJ_POP(decile,sample))",
    "TAX_BURDEN_PRE": (
        'LAMBDA(decile,sample,TOTAL_EFFECT(decile,sample,"mean")/100*ADJ_CONS_TOT(decile,sample))'
    ),

    # --- Step 8: revenue recycling -----------------------------------------
    "PIT_SHARE_RAW": (
        'LAMBDA(decile,BSHARE(IF(PITSourceSwitch="income","pit_share_income","pit_share_region"),'
        'decile,"Overall","mean"))'
    ),
    # the raw decile shares in HHSurvey do not sum to exactly 100% (sampling/
    # rounding), so -- exactly as the source workbook does -- rescale them to
    # sum to 100 before applying to the PIT baseline total.
    "PIT_SHARE_SUM": "LAMBDA(SUM(PIT_SHARE_RAW(DECILE_ARRAY)))",
    "PIT_LIABILITY": (
        "LAMBDA(decile,PIT_SHARE_RAW(decile)/PIT_SHARE_SUM()*PITBaselineLCU)"
    ),
    "TOTAL_ADJ_POP": 'LAMBDA(SUM(BSHARE("popw",DECILE_ARRAY,"Overall","mean"))*PopAdjFactor)',
    "PIT_MAX_TRANSFER_PC": "LAMBDA(RevLaborTaxLCU/TOTAL_ADJ_POP())",
    # MAX(0, array) would NOT clamp element-wise here -- Excel's MIN/MAX
    # reduce ALL arguments (scalars and array elements alike) to a single
    # value rather than broadcasting, so MAX(0, {array of 10}) collapses to
    # one number instead of 10 clamped ones. IF(...) broadcasts correctly.
    "PIT_SHORTFALL_TOTAL": (
        'LAMBDA(LET(gap,PIT_MAX_TRANSFER_PC()-PIT_LIABILITY(DECILE_ARRAY)/ADJ_POP(DECILE_ARRAY,"Overall"),'
        'clipped,IF(gap>0,gap,0),'
        'SUM(clipped*ADJ_POP(DECILE_ARRAY,"Overall"))))'
    ),
    "PIT_ADDITIONAL_PC": "LAMBDA(PIT_SHORTFALL_TOTAL()/TOTAL_ADJ_POP())",
    "PIT_LIABILITY_TOTAL_EXEMPT": (
        "LAMBDA(SUM(IF(DECILE_ARRAY<=ExemptBottomDeciles_PIT,PIT_LIABILITY(DECILE_ARRAY),0)))"
    ),
    "PIT_REDUCTION": (
        "LAMBDA(decile,"
        "LET(liability,PIT_LIABILITY(decile),"
        'pop,ADJ_POP(decile,"Overall"),'
        "prop_amt,liability*LaborTaxCutCoef,"
        "liab_pc,liability/pop,"
        "capped_pc,IF(PIT_MAX_TRANSFER_PC()<liab_pc,PIT_MAX_TRANSFER_PC(),liab_pc),"
        "pa_pc,capped_pc+PIT_ADDITIONAL_PC(),"
        "pa_amt,pa_pc*pop,"
        # both operands here are always scalar (RevLaborTaxLCU and
        # PIT_LIABILITY_TOTAL_EXEMPT() are workbook-level totals, not
        # per-decile), so plain MIN is fine.
        "te_scale,MIN(1,RevLaborTaxLCU/PIT_LIABILITY_TOTAL_EXEMPT()),"
        "te_amt,IF(decile<=ExemptBottomDeciles_PIT,liability*te_scale,0),"
        'IF(LaborTaxMethod="Proportional Compensation",prop_amt,'
        'IF(LaborTaxMethod="Personal Allowance",pa_amt,te_amt))))'
    ),
    "TOTAL_UNSERVED_POP": (
        'LAMBDA(SUM(BSHARE("popw",DECILE_ARRAY,"Overall","mean")'
        '*(1-BSHARE("all_acs_share",DECILE_ARRAY,"Overall","mean")/100)))'
    ),
    "INFRA_SHARE": (
        "LAMBDA(decile,"
        'LET(pop,BSHARE("popw",decile,"Overall","mean"),'
        'access,BSHARE("all_acs_share",decile,"Overall","mean"),'
        "(pop*(1-access/100))/TOTAL_UNSERVED_POP()))"
    ),
    "TARGETED_TRANSFER": "LAMBDA(decile,RevTargetedTransferLCU*INFRA_SHARE(decile))",
    "PUBLIC_INVESTMENT": "LAMBDA(decile,RevPublicInvestLCU*INFRA_SHARE(decile))",
    # Same INDEX(range,0,{array}) fix as ELASTADJ above.
    "ASPIRE_PC": (
        "LAMBDA(program,decile,"
        "LET(q,ROUNDUP(decile/2,0),"
        "r,MATCH(program,ASPIRE_Programs,0),"
        "row,INDEX(ASPIRE_Values,r,0),"
        "CHOOSECOLS(row,q)))"
    ),
    "ASPIRE_TOTAL": (
        'LAMBDA(program,SUMPRODUCT(ASPIRE_PC(program,DECILE_ARRAY)*BSHARE("popw",DECILE_ARRAY,"Overall","mean")))'
    ),
    "ASPIRE_SHARE": (
        "LAMBDA(program,decile,"
        'LET(pop,BSHARE("popw",decile,"Overall","mean"),'
        "pc,ASPIRE_PC(program,decile),"
        "(pc*pop)/ASPIRE_TOTAL(program)))"
    ),
    "CURRENT_SPENDING": 'LAMBDA(decile,RevCurrentSpendLCU*ASPIRE_SHARE("allsp",decile))',
    "AMOUNT_RECYCLED": (
        "LAMBDA(decile,"
        "PIT_REDUCTION(decile)+TARGETED_TRANSFER(decile)+PUBLIC_INVESTMENT(decile)+CURRENT_SPENDING(decile))"
    ),
    "TOTAL_RECYCLED": "LAMBDA(SUM(AMOUNT_RECYCLED(DECILE_ARRAY)))",

    # --- Step 9: net effect, post-CP consumption & shares ------------------
    "NET_EFFECT": (
        'LAMBDA(decile,sample,'
        '(TOTAL_EFFECT(decile,sample,"mean")/100-AMOUNT_RECYCLED(decile)/ADJ_CONS_TOT(decile,sample))*100)'
    ),
    "POST_CP_EXCL_RECYCLING": (
        'LAMBDA(decile,sample,ADJ_CONS_TOT(decile,sample)*(1-TOTAL_EFFECT(decile,sample,"mean")/100))'
    ),
    "POST_CP_INCL_RECYCLING": (
        "LAMBDA(decile,sample,POST_CP_EXCL_RECYCLING(decile,sample)+AMOUNT_RECYCLED(decile))"
    ),
}
for name, formula in LAMBDAS.items():
    define_name(name, "=" + formula)
print(f"{len(LAMBDAS)} LAMBDA functions defined.")

# ---------------------------------------------------------------------------
# Distribution_Outputs: sections C.II - F. Column layout is standardized
# across the whole sheet (the source workbook drifts across sections -- see
# distribution_sections_map.md -- this rebuild does not replicate that
# drift): A=section code, B=section title, C=item label, D=statistic,
# E=sample, F=quantile type, G=item code, H=unit, P=basket value,
# Q:Z=decile 1-10 values.
# ---------------------------------------------------------------------------
OC = {"code": 1, "title": 2, "label": 3, "stat": 4, "sample": 5, "qtype": 6,
      "item": 7, "unit": 8, "basket": 16, "dec1": 17}

ws_out = new_output_sheet("Distribution_Outputs", tab_color="70AD47")
ws_out.sheet_view.showGridLines = False


def out_section_header(r, code, title):
    write_section_header(ws_out, r, code, title)
    return r + 2


def out_block_header(r, title, item_code):
    style_cell(ws_out, r, OC["label"], title, BOLD_FONT)
    style_cell(ws_out, r, OC["item"], item_code, BASE_FONT)
    style_cell(ws_out, r, OC["basket"] - 1, "Decile -->", ITALIC_FONT)
    for i in range(10):
        style_cell(ws_out, r, OC["dec1"] + i, i + 1, ITALIC_FONT)
    return r + 1


def out_item_block(r, title, item_code, unit, formula_fn):
    """4 stats x 3 samples x (Basket, Deciles) = 24 rows, matching the
    source's C.III/C.IV/C.V/C.VI block shape. formula_fn(decile_arg, sample,
    stat) returns a formula string; decile_arg is '0' for the basket row and
    'DECILE_ARRAY' (spills across Q:Z) for the deciles row."""
    r = out_block_header(r, title, item_code)
    for stat in STATS:
        for sample in SAMPLES:
            style_cell(ws_out, r, OC["stat"], stat)
            style_cell(ws_out, r, OC["sample"], sample)
            style_cell(ws_out, r, OC["qtype"], "Basket")
            style_cell(ws_out, r, OC["item"], item_code)
            style_cell(ws_out, r, OC["unit"], unit)
            style_cell(ws_out, r, OC["basket"], formula_fn("0", sample, stat), fill=CALC_FILL)
            r += 1
            style_cell(ws_out, r, OC["stat"], stat)
            style_cell(ws_out, r, OC["sample"], sample)
            style_cell(ws_out, r, OC["qtype"], "Deciles")
            style_cell(ws_out, r, OC["item"], item_code)
            style_cell(ws_out, r, OC["unit"], unit)
            style_cell(ws_out, r, OC["dec1"], formula_fn("DECILE_ARRAY", sample, stat), fill=CALC_FILL)
            r += 1
    return r + 1  # blank separator row


for c in range(1, 27):
    ws_out.column_dimensions[get_column_letter(c)].width = 11
ws_out.column_dimensions["C"].width = 42
ws_out.freeze_panes = "C1"

r = 1
style_cell(ws_out, r, 1, "Distributional Effects Module - Egypt: Outputs (Standalone)", TITLE_FONT)
r += 2

r = out_section_header(r, "C.", "Carbon tax incidence calculations")

# --- C.I: price changes -- taken as given inputs; see Distribution_Inputs / Price_Changes ---
r = out_section_header(r, "C.I.", "Price changes from policy scenario")
style_cell(ws_out, r, OC["label"],
           "Direct (by fuel) and indirect (by CPAT category) price changes are taken as given "
           "inputs from the CPAT Mitigation module's carbon-price scenario -- see the Price_Changes "
           "tab and Distribution_Inputs A.II for the scope note. Price-elasticity adjustment factors "
           "(ElastAdj table) are likewise given inputs, applied inside DIRECT_EFFECT/INDIRECT_EFFECT.",
           ITALIC_FONT, align=Alignment(wrap_text=True))
ws_out.merge_cells(start_row=r, start_column=OC["label"], end_row=r + 1, end_column=20)
ws_out.row_dimensions[r].height = 30
r += 3

# --- C.II: elasticities / behavioral / DWL -- reference display (Overall, all stats) ---
r = out_section_header(r, "C.II.", "Price elasticities of demand, behavioral/structural change and deadweight loss (DWL) adjustments")
style_cell(ws_out, r, OC["label"],
           "Reference display of the elasticities, price changes (post elasticity-adjustment) and "
           "deadweight losses used inside Steps 4-5 (C.V/C.VI). Behavioral-change and DWL switches "
           "are in Distribution_Inputs B.I; for Egypt's current scenario both are off, so "
           "BEHAVIOR_ADJ()=1 and DWL()=0 throughout.", ITALIC_FONT, align=Alignment(wrap_text=True))
ws_out.merge_cells(start_row=r, start_column=OC["label"], end_row=r + 1, end_column=20)
ws_out.row_dimensions[r].height = 30
r += 3
style_cell(ws_out, r, OC["label"], "Item", BOLD_FONT)
style_cell(ws_out, r, OC["item"], "Code", BOLD_FONT)
style_cell(ws_out, r, OC["unit"], "Elasticity (Overall, mean)", BOLD_FONT)
style_cell(ws_out, r, OC["basket"] - 1, "Decile -->", BOLD_FONT)
for i in range(10):
    style_cell(ws_out, r, OC["dec1"] + i, i + 1, BOLD_FONT)
style_cell(ws_out, r, OC["dec1"] + 10, "Price chg. x ElastAdj (%)", BOLD_FONT)
style_cell(ws_out, r, OC["dec1"] + 11, "DWL, Basket (%)", BOLD_FONT)
r += 1
for fuel in FUEL_CODES:
    style_cell(ws_out, r, OC["label"], FUEL_LABEL[fuel])
    style_cell(ws_out, r, OC["item"], fuel)
    style_cell(ws_out, r, OC["dec1"] - 1,
               f'=ELAST("{FUEL_ELAST_COL[fuel]}",0,"Overall","mean")', fill=CALC_FILL)
    style_cell(ws_out, r, OC["dec1"],
               f'=ELAST("{FUEL_ELAST_COL[fuel]}",DECILE_ARRAY,"Overall","mean")', fill=CALC_FILL)
    style_cell(ws_out, r, OC["dec1"] + 10,
               f'=PCHANGE_DIRECT("{fuel}")*ELASTADJ("{fuel}",0)', fill=CALC_FILL)
    style_cell(ws_out, r, OC["dec1"] + 11,
               f'=DWL("{FUEL_SHARE_COL[fuel]}","{FUEL_ELAST_COL[fuel]}",0,"Overall","mean",'
               f'PCHANGE_DIRECT("{fuel}")*ELASTADJ("{fuel}",0))', fill=CALC_FILL)
    r += 1
for cat in CAT_CODES:
    style_cell(ws_out, r, OC["label"], CAT_LABEL[cat])
    style_cell(ws_out, r, OC["item"], cat)
    style_cell(ws_out, r, OC["dec1"] - 1,
               f'=ELAST("{CAT_ELAST_COL[cat]}",0,"Overall","mean")', fill=CALC_FILL)
    style_cell(ws_out, r, OC["dec1"],
               f'=ELAST("{CAT_ELAST_COL[cat]}",DECILE_ARRAY,"Overall","mean")', fill=CALC_FILL)
    style_cell(ws_out, r, OC["dec1"] + 10,
               f'=PCHANGE_INDIRECT("{cat}")*ELASTADJ("{cat}",0)', fill=CALC_FILL)
    style_cell(ws_out, r, OC["dec1"] + 11,
               f'=DWL("{CAT_SHARE_COL[cat]}","{CAT_ELAST_COL[cat]}",0,"Overall","mean",'
               f'PCHANGE_INDIRECT("{cat}")*ELASTADJ("{cat}",0))', fill=CALC_FILL)
    r += 1
r += 1

# --- C.III: direct budget shares (11 fuels incl. biomass) ---
r = out_section_header(r, "C.III.", "Household budget shares - direct fuel consumption")
for fuel in FUEL_CODES_ALL:
    sharecol = FUEL_SHARE_COL[fuel]
    fn = lambda dec, sample, stat, col=sharecol: f'=BSHARE("{col}",{dec},"{sample}","{stat}")'
    r = out_item_block(r, f"{FUEL_LABEL[fuel]} Budget Share (% Total Consumption)", fuel, "%", fn)

# --- C.IV: indirect budget shares (14 categories) ---
r = out_section_header(r, "C.IV.", "Household budget shares - indirect (non-fuel) consumption")
for cat in CAT_CODES:
    sharecol = CAT_SHARE_COL[cat]
    fn = lambda dec, sample, stat, col=sharecol: f'=BSHARE("{col}",{dec},"{sample}","{stat}")'
    r = out_item_block(r, f"{CAT_LABEL[cat]} Budget Share (% Total Consumption)", cat, "%", fn)

# --- C.V: direct effects (8 taxed fuels + Total) ---
r = out_section_header(r, "C.V.", "Consumption losses - direct effect")
for fuel in FUEL_CODES:
    fn = (lambda dec, sample, stat, f=fuel:
          f'=DIRECT_EFFECT("{f}","{FUEL_LABEL[f]}","{FUEL_SHARE_COL[f]}","{FUEL_ELAST_COL[f]}",{dec},"{sample}","{stat}")')
    r = out_item_block(r, f"Direct Effect - {FUEL_LABEL[fuel]} (% Total Consumption)", fuel, "%", fn)
fn = lambda dec, sample, stat: f'=TOTAL_DIRECT_EFFECT({dec},"{sample}","{stat}")'
r = out_item_block(r, "Total Direct Effect (% Total Consumption)", "tot", "%", fn)

# --- C.VI: indirect effects (14 categories + Total) ---
r = out_section_header(r, "C.VI.", "Consumption losses - indirect effect")
for cat in CAT_CODES:
    fn = (lambda dec, sample, stat, c=cat:
          f'=INDIRECT_EFFECT("{c}","{CAT_SHARE_COL[c]}","{CAT_ELAST_COL[c]}",{dec},"{sample}","{stat}")')
    r = out_item_block(r, f"{CAT_LABEL[cat]} Indirect Effect (% Total Consumption)", cat, "%", fn)
fn = lambda dec, sample, stat: f'=TOTAL_INDIRECT_EFFECT({dec},"{sample}","{stat}")'
r = out_item_block(r, "Total Indirect Effect (% Total Consumption)", "tot", "%", fn)

# --- C.VII: total effect ---
r = out_section_header(r, "C.VII.", "Consumption losses - total effect")
fn = lambda dec, sample, stat: f'=TOTAL_EFFECT({dec},"{sample}","{stat}")'
r = out_item_block(r, "Total Effect - Carbon Tax Burden (% Total Consumption)", "tot", "%", fn)

C_VII_END_ROW = r
print("C sections (II-VII) written up to row", r)


def out_scalar_row(r, label, item_code, unit, basket_formula, deciles_formula):
    style_cell(ws_out, r, OC["label"], label)
    style_cell(ws_out, r, OC["item"], item_code)
    style_cell(ws_out, r, OC["unit"], unit)
    style_cell(ws_out, r, OC["basket"], basket_formula, fill=CALC_FILL)
    style_cell(ws_out, r, OC["dec1"], deciles_formula, fill=CALC_FILL)
    return r + 1


# --- C.VIII: survey-to-national-accounts rebasing ---
r = out_section_header(r, "C.VIII.", "Household-survey-to-national-accounts adjustments")
r = out_block_header(r, "Adjusted population and consumption (Step 7)", "")
for sample in SAMPLES:
    r = out_scalar_row(r, f"Adjusted Population ({sample})", "pop_tot_adj", "Individuals",
                        f'=SUM(ADJ_POP(DECILE_ARRAY,"{sample}"))', f'=ADJ_POP(DECILE_ARRAY,"{sample}")')
for sample in SAMPLES:
    r = out_scalar_row(r, f"NA-Adjusted Per-Capita Consumption ({sample})", "cons_pc_na", "Real 2026 LCU/Individual",
                        f'=ADJ_CONS_PC(0,"{sample}")', f'=ADJ_CONS_PC(DECILE_ARRAY,"{sample}")')
for sample in SAMPLES:
    r = out_scalar_row(r, f"NA-Adjusted Total Consumption ({sample})", "cons_tot_na", "Real 2026 LCU",
                        f'=SUM(ADJ_CONS_TOT(DECILE_ARRAY,"{sample}"))', f'=ADJ_CONS_TOT(DECILE_ARRAY,"{sample}")')
for sample in SAMPLES:
    r = out_scalar_row(r, f"Tax Burden, Pre-Behavioral Response ({sample})", "ct_bdn_pre", "Real 2026 LCU",
                        f'=SUM(TAX_BURDEN_PRE(DECILE_ARRAY,"{sample}"))', f'=TAX_BURDEN_PRE(DECILE_ARRAY,"{sample}")')
r += 1

# --- C.IX moved after C.X/C.XI in calculation order, but per source layout
#     is displayed here; formulas reference the AMOUNT_RECYCLED family of
#     LAMBDAs directly (defined in the Name Manager), so display order does
#     not affect correctness ---
C_IX_ROW = r
r = out_section_header(r, "C.IX.", "Change in Gini coefficient (Lorenz curve)")
r = out_block_header(r, "Consumption shares & cumulative shares (Overall sample, mean)", "")
r = out_scalar_row(r, "Baseline Consumption Share per Decile", "shr_base", "Ratio",
                    "", '=ADJ_CONS_TOT(DECILE_ARRAY,"Overall")/SUM(ADJ_CONS_TOT(DECILE_ARRAY,"Overall"))')
r = out_scalar_row(r, "Post-CP Consumption Share, excl. Recycling", "shr_excl", "Ratio",
                    "", '=POST_CP_EXCL_RECYCLING(DECILE_ARRAY,"Overall")/SUM(POST_CP_EXCL_RECYCLING(DECILE_ARRAY,"Overall"))')
r = out_scalar_row(r, "Post-CP Consumption Share, incl. Recycling", "shr_incl", "Ratio",
                    "", '=POST_CP_INCL_RECYCLING(DECILE_ARRAY,"Overall")/SUM(POST_CP_INCL_RECYCLING(DECILE_ARRAY,"Overall"))')
BASE_SHARE_ROW = r - 3
EXCL_SHARE_ROW = r - 2
INCL_SHARE_ROW = r - 1


def cum_share_formula(share_row):
    col = get_column_letter(OC["dec1"])
    return f'=SCAN(0,{col}{share_row}:{get_column_letter(OC["dec1"]+9)}{share_row},LAMBDA(a,v,a+v))'


r = out_scalar_row(r, "Cumulative Baseline Consumption Share", "cum_base", "Ratio", "=0", cum_share_formula(BASE_SHARE_ROW))
r = out_scalar_row(r, "Cumulative Post-CP Share, excl. Recycling", "cum_excl", "Ratio", "=0", cum_share_formula(EXCL_SHARE_ROW))
r = out_scalar_row(r, "Cumulative Post-CP Share, incl. Recycling", "cum_incl", "Ratio", "=0", cum_share_formula(INCL_SHARE_ROW))
r = out_scalar_row(r, "Cumulative Share, Perfect Equality", "cum_eq", "Ratio", "=0", "=DECILE_ARRAY/10")
CUM_BASE_ROW, CUM_EXCL_ROW, CUM_INCL_ROW, CUM_EQ_ROW = r - 4, r - 3, r - 2, r - 1
r += 1


# Gini via trapezoid rule: area = SUM(0.1*(a+b)/2) over the 10 decile
# intervals, a=cumulative share at start of interval, b=at end; Gini=1-2*area
style_cell(ws_out, r, OC["label"], "Gini coefficient (Baseline / Post-CP excl. / incl. recycling / Perfect equality)", BOLD_FONT)
r += 1
for label, cum_row in [("Baseline", CUM_BASE_ROW), ("Post-CP excl. recycling", CUM_EXCL_ROW),
                        ("Post-CP incl. recycling", CUM_INCL_ROW), ("Perfect equality", CUM_EQ_ROW)]:
    col0 = get_column_letter(OC["basket"])
    col1 = get_column_letter(OC["dec1"])
    style_cell(ws_out, r, OC["label"], f"Gini - {label}")
    a_range = f'HSTACK(0,{col1}{cum_row}:INDEX({col1}{cum_row}:{get_column_letter(OC["dec1"]+9)}{cum_row},1,9))'
    b_range = f'{col1}{cum_row}:{get_column_letter(OC["dec1"]+9)}{cum_row}'
    # Lorenz-curve widths use each decile's actual (survey-weighted)
    # population share, not a flat 1/10th -- deciles are close to but not
    # exactly equal-sized once survey weights are applied.
    pop_w = 'ADJ_POP(DECILE_ARRAY,"Overall")/SUM(ADJ_POP(DECILE_ARRAY,"Overall"))'
    formula = f'=1-2*SUMPRODUCT({pop_w}*({a_range}+{b_range})/2)'
    style_cell(ws_out, r, OC["basket"], formula, fill=CALC_FILL)
    if label == "Baseline":
        GINI_BASE_ROW = r
    elif label == "Post-CP incl. recycling":
        GINI_INCL_ROW = r
    elif label == "Post-CP excl. recycling":
        GINI_EXCL_ROW = r
    r += 1
col0 = get_column_letter(OC["basket"])
style_cell(ws_out, r, OC["label"], "Delta Gini (excl. recycling)")
style_cell(ws_out, r, OC["basket"], f"={col0}{GINI_EXCL_ROW}-{col0}{GINI_BASE_ROW}", fill=CALC_FILL)
r += 1
style_cell(ws_out, r, OC["label"], "Delta Gini (incl. recycling)")
style_cell(ws_out, r, OC["basket"], f"={col0}{GINI_INCL_ROW}-{col0}{GINI_BASE_ROW}", fill=CALC_FILL)
r += 2

# --- C.X: revenue recycling ---
r = out_section_header(r, "C.X.", "Targeted transfers, public investment and current spending")
style_cell(ws_out, r, OC["label"],
           "Targeted transfers and public investment are allocated using the infrastructure-access "
           "shortfall (INFRA_SHARE -- population share currently lacking access, from HHSurvey "
           "all_acs_share), consistent with Egypt's non-ASPIRE targeted-transfer configuration. The "
           "source workbook's targeted transfer instead uses a bottom-40%-of-population poverty-line "
           "proxy with coverage/leakage adjustments (see Distribution_Inputs B.I); that specific "
           "targeting mechanism is approximated here by INFRA_SHARE for tractability -- amounts sum "
           "to the same total revenue, decile allocation will differ somewhat from the source. Current "
           "spending uses ASPIRE incidence (program 'allsp', by quintile).", ITALIC_FONT,
           align=Alignment(wrap_text=True))
ws_out.merge_cells(start_row=r, start_column=OC["label"], end_row=r + 3, end_column=20)
ws_out.row_dimensions[r].height = 60
r += 5
r = out_scalar_row(r, "Labor Tax / PIT Reduction", "pit", "Real 2026 LCU",
                    "=SUM(PIT_REDUCTION(DECILE_ARRAY))", "=PIT_REDUCTION(DECILE_ARRAY)")
r = out_scalar_row(r, "Targeted Transfers Received", "trns", "Real 2026 LCU",
                    "=SUM(TARGETED_TRANSFER(DECILE_ARRAY))", "=TARGETED_TRANSFER(DECILE_ARRAY)")
r = out_scalar_row(r, "Public Investment Received", "inv", "Real 2026 LCU",
                    "=SUM(PUBLIC_INVESTMENT(DECILE_ARRAY))", "=PUBLIC_INVESTMENT(DECILE_ARRAY)")
r = out_scalar_row(r, "Current Spending Received", "spnd", "Real 2026 LCU",
                    "=SUM(CURRENT_SPENDING(DECILE_ARRAY))", "=CURRENT_SPENDING(DECILE_ARRAY)")
r = out_scalar_row(r, "Total Amount Recycled", "tot", "Real 2026 LCU",
                    "=SUM(AMOUNT_RECYCLED(DECILE_ARRAY))", "=AMOUNT_RECYCLED(DECILE_ARRAY)")
r += 1

# --- C.XI: PIT reductions (detail) ---
r = out_section_header(r, "C.XI.", "Personal income tax (PIT) reductions")
style_cell(ws_out, r, OC["label"],
           'Active method: Distribution_Inputs "Labor Tax Reduction Method" (LaborTaxMethod). '
           "Personal Allowance caps each decile's per-capita tax cut at the average revenue-per-"
           "capita, then redistributes the unused allowance from below-cap deciles as an equal "
           "per-capita top-up to every decile -- see PIT_REDUCTION in the Name Manager.", ITALIC_FONT,
           align=Alignment(wrap_text=True))
ws_out.merge_cells(start_row=r, start_column=OC["label"], end_row=r + 1, end_column=20)
ws_out.row_dimensions[r].height = 30
r += 3
r = out_scalar_row(r, "PIT Liability (baseline)", "liab", "Real 2026 LCU",
                    "=SUM(PIT_LIABILITY(DECILE_ARRAY))", "=PIT_LIABILITY(DECILE_ARRAY)")
r = out_scalar_row(r, "PIT Reduction Received", "pit", "Real 2026 LCU",
                    "=SUM(PIT_REDUCTION(DECILE_ARRAY))", "=PIT_REDUCTION(DECILE_ARRAY)")
r += 1

# --- Step 9 outputs: net effect, post-CP consumption ---
r = out_section_header(r, "C.XII.", "Net effect and post-CP consumption (Step 9)")
for sample in SAMPLES:
    fn = lambda dec, sample=sample, stat=None: f'=NET_EFFECT({dec},"{sample}")'
    r = out_block_header(r, f"Net Effect, Post-Recycling ({sample}, % Total Consumption)", "tot")
    style_cell(ws_out, r, OC["basket"], f'=NET_EFFECT(0,"{sample}")', fill=CALC_FILL)
    style_cell(ws_out, r, OC["dec1"], f'=NET_EFFECT(DECILE_ARRAY,"{sample}")', fill=CALC_FILL)
    r += 2
for sample in SAMPLES:
    r = out_scalar_row(r, f"Post-CP Consumption, excl. Recycling ({sample})", "cons_excl", "Real 2026 LCU",
                        f'=SUM(POST_CP_EXCL_RECYCLING(DECILE_ARRAY,"{sample}"))',
                        f'=POST_CP_EXCL_RECYCLING(DECILE_ARRAY,"{sample}")')
for sample in SAMPLES:
    r = out_scalar_row(r, f"Post-CP Consumption, incl. Recycling ({sample})", "cons_incl", "Real 2026 LCU",
                        f'=SUM(POST_CP_INCL_RECYCLING(DECILE_ARRAY,"{sample}"))',
                        f'=POST_CP_INCL_RECYCLING(DECILE_ARRAY,"{sample}")')
r += 1

# --- Step 12: compensation requirement ---
r = out_section_header(r, "C.XIII.", "Compensation requirement (share of CP revenue needed per decile)")
LOSS_ROW = r
r = out_scalar_row(r, "Loss Amount (Total Effect x Baseline Consumption Share)", "loss", "Real 2026 LCU",
                    '=SUM(TOTAL_EFFECT(DECILE_ARRAY,"Overall","mean")/100*ADJ_CONS_TOT(DECILE_ARRAY,"Overall"))',
                    '=TOTAL_EFFECT(DECILE_ARRAY,"Overall","mean")/100*ADJ_CONS_TOT(DECILE_ARRAY,"Overall")')
col1 = get_column_letter(OC["dec1"])
col10 = get_column_letter(OC["dec1"] + 9)
col0 = get_column_letter(OC["basket"])
style_cell(ws_out, r, OC["label"], "Share of CP Revenue Required to Compensate Decile")
style_cell(ws_out, r, OC["dec1"], f"={col1}{LOSS_ROW}:{col10}{LOSS_ROW}/{col0}{LOSS_ROW}", fill=CALC_FILL)
SHARE_ROW = r
r += 1
style_cell(ws_out, r, OC["label"], "Cumulative Share of CP Revenue Required (poorest -> richest)")
style_cell(ws_out, r, OC["basket"], "=0")
style_cell(ws_out, r, OC["dec1"],
           f'=SCAN(0,{col1}{SHARE_ROW}:{col10}{SHARE_ROW},LAMBDA(a,v,a+v))', fill=CALC_FILL)
r += 2

# --- D: chart-ready outputs (direct references into the C sections above) ---
r = out_section_header(r, "D.", "Outputs for charts")
style_cell(ws_out, r, OC["label"],
           "Chart-ready values are direct references to the C-section cells above (no new "
           "computation) -- build charts in Excel directly off the C.III-C.XIII ranges, filtered to "
           'the "Statistic - Outputs" / sample of interest, exactly as the source workbook\'s D '
           "section re-projects its own C-section data. A pivot-friendly export of the core outputs "
           '(budget share, direct/indirect/total/net effect, all stats x samples x deciles) is on '
           "the Distribution_Outputs sheet above; the sectoral GTAP top-20 chart (source D.VI) is "
           "out of scope here since sector-level price increases are a given input in this build, "
           "not computed (see Distribution_Inputs A.II).", ITALIC_FONT, align=Alignment(wrap_text=True))
ws_out.merge_cells(start_row=r, start_column=OC["label"], end_row=r + 4, end_column=20)
ws_out.row_dimensions[r].height = 90
r += 6

# --- E: notes (methodology, static text) ---
r = out_section_header(r, "E.", "Notes")
E_NOTES = [
    ("Direct effect", "budget_share x price_change_direct x behavior_adj - DWL, summed over the 8 "
     "taxed fuels. See LAMBDA DIRECT_EFFECT / TOTAL_DIRECT_EFFECT."),
    ("Indirect effect", "budget_share x price_change_indirect x behavior_adj - DWL, summed over the "
     "14 CPAT consumption categories. See LAMBDA INDIRECT_EFFECT / TOTAL_INDIRECT_EFFECT."),
    ("Deadweight loss (Harberger triangle)", "0.5 x elasticity x price_change^2 x budget_share "
     "(only applied if Distribution_Inputs 'Adjust for deadweight losses?' = Yes)."),
    ("Survey-to-national-accounts rebasing", "per-capita consumption from HHSurvey is deflated to "
     "the analysis year and rescaled by Household Consumption Adj. Factor = National Accounts "
     "Consumption / Total Household Consumption (Survey Year); population is rescaled by "
     "Population Adj. Factor = Population (Analysis Year) / Population (Survey Year)."),
    ("Gini coefficient", "1 - 2 x (population-share-weighted trapezoidal area under the Lorenz "
     "curve), using each decile's actual survey-weighted population share as the trapezoid width."),
    ("Revenue recycling", "labor tax / PIT reduction via the active method in Distribution_Inputs "
     "(Personal Allowance / Proportional Compensation / Targeted Exemption); targeted transfers and "
     "public investment via the infrastructure-access shortfall; current spending via ASPIRE "
     "incidence."),
]
for title, body in E_NOTES:
    style_cell(ws_out, r, OC["label"], title, BOLD_FONT)
    r += 1
    style_cell(ws_out, r, OC["label"], body, ITALIC_FONT, align=Alignment(wrap_text=True))
    ws_out.merge_cells(start_row=r, start_column=OC["label"], end_row=r, end_column=20)
    r += 2

# --- F: handoff back to the Mitigation module ---
r = out_section_header(r, "F.", "Results for MT (handoff to Mitigation module)")
style_cell(ws_out, r, OC["label"],
           "Direct and indirect price changes are echoed back from the given inputs (Price_Changes "
           "tab) for consistency display in the Mitigation module's own results view. The GTAP-"
           "sector-level revenue-reconciliation factor and the top-20-most-affected-sectors table "
           "(source F, from C.I/D.VI) are out of scope here -- see Distribution_Inputs A.II.",
           ITALIC_FONT, align=Alignment(wrap_text=True))
ws_out.merge_cells(start_row=r, start_column=OC["label"], end_row=r + 2, end_column=20)
ws_out.row_dimensions[r].height = 55
r += 4
style_cell(ws_out, r, OC["label"], "Direct price changes by fuel (%)", BOLD_FONT)
r += 1
for fuel in FUEL_CODES:
    style_cell(ws_out, r, OC["label"], FUEL_LABEL[fuel])
    style_cell(ws_out, r, OC["item"], fuel)
    style_cell(ws_out, r, OC["basket"], f'=PCHANGE_DIRECT("{fuel}")', fill=CALC_FILL)
    r += 1
r += 1
style_cell(ws_out, r, OC["label"], "Indirect price changes by CPAT category (%)", BOLD_FONT)
r += 1
for cat in CAT_CODES:
    style_cell(ws_out, r, OC["label"], CAT_LABEL[cat])
    style_cell(ws_out, r, OC["item"], cat)
    style_cell(ws_out, r, OC["basket"], f'=PCHANGE_INDIRECT("{cat}")', fill=CALC_FILL)
    r += 1

print("Distribution_Outputs written, last row", r)

# ---------------------------------------------------------------------------
# ReadMe
# ---------------------------------------------------------------------------
ws_r = wb.create_sheet("ReadMe")
ws_r.sheet_view.showGridLines = False
ws_r.column_dimensions["A"].width = 3
ws_r.column_dimensions["B"].width = 100
rr = 1
style_cell(ws_r, rr, 2, "CPAT Distribution Module -- Standalone Workbook (Egypt)", TITLE_FONT)
rr += 2
README_PARAS = [
    ("What this is", BOLD_FONT,
     "A standalone rebuild of the household distributional/incidence module from CPAT "
     "1.0pre_456 (cpat_excel/original/CPAT 1.0pre_456_NoPropData.xlsb, sheet \"Distribution\"), "
     "scoped to Egypt only. It replicates the source sheet's sections and conventions but "
     "replaces its per-cell formulas with a small library of reusable LAMBDA functions (see the "
     "Name Manager), and splits the source's single combined sheet into inputs (Distribution_"
     "Inputs) and calculated outputs (Distribution_Outputs)."),
    ("Structure", BOLD_FONT,
     "Distribution_Inputs = source Sections A (trimmed to Egypt) + B (key assumptions/policy "
     "inputs), same row labels, largely the same values. Distribution_Outputs = source Sections "
     "C-F (budget shares, direct/indirect/total effect, rebasing, Gini/Lorenz, revenue recycling, "
     "PIT, net effect, compensation, chart references, notes, Mitigation handoff), rebuilt with "
     "live formulas. HHSurvey / HH_Elast / ASPIRE / WHOCooking / GDPRatios / IO_GTAP / Mapping = "
     "the source DATA_DISTN tables, filtered to Egypt (Mapping crosswalks are country-independent "
     "and kept in full). Price_Changes = the given price-change inputs (see below)."),
    ("What is computed vs. given", BOLD_FONT,
     "Direct (by fuel) and indirect (by CPAT category) price changes, and the price-elasticity "
     "adjustment factors that modify them, are taken as GIVEN inputs -- pulled from the source "
     "workbook's current carbon-price scenario for Egypt, per the request that started this build. "
     "This mirrors the module's own architecture: Distribution is a consumer, not a producer, of "
     "the Mitigation module's price changes. Deriving these from the raw GTAP Leontief/IO table "
     "would additionally require the Mitigation module's own revenue-reconciliation inputs and "
     "IEA/GAINS/IMF emissions-recalibration data, which are not part of the DATA_DISTN tables and "
     "are out of scope here. Population, GDP and GDP-deflator inputs are likewise Mitigation-"
     "module macro outputs and are taken as given. Everything downstream -- budget shares, direct/"
     "indirect/total effect, survey-to-national-accounts rebasing, revenue recycling, net effect, "
     "Gini/Lorenz, compensation shares -- is computed live from the data tabs and validated cell-"
     "for-cell against the source workbook's cached values for Egypt (see Verification below)."),
    ("Known simplification: targeted-transfer/public-investment targeting", BOLD_FONT,
     "The source workbook's targeted-transfer channel uses a bottom-40%-of-population poverty-"
     "line proxy with coverage/leakage parameters (Section C.X). This build approximates both "
     "targeted transfers and public investment with the infrastructure-access shortfall (INFRA_"
     "SHARE, from HHSurvey's all_acs_share) -- a fallback mechanism the source model itself uses "
     "elsewhere. Totals match the source's revenue allocation exactly; the decile-by-decile split "
     "for these two channels will differ somewhat from the source's specific targeting mechanism."),
    ("Verification", BOLD_FONT,
     "Every LAMBDA formula design in this workbook was validated in Python against the source "
     "workbook's cached values for Egypt before being encoded as an Excel formula: budget shares, "
     "direct/indirect/total effect, national-accounts rebasing, PIT reduction (Personal Allowance "
     "method, incl. the undercoverage redistribution step) and the baseline Gini coefficient all "
     "reproduce the source to at least 10 significant figures. Net effect / Gini-including-"
     "recycling will differ slightly from the source to the extent of the targeted-transfer/"
     "public-investment simplification noted above."),
    ("Conventions", BOLD_FONT,
     "Distribution_Outputs uses one consistent column layout throughout (A=section code, "
     "B=section title, C=item label, D=statistic, E=sample, F=quantile type, G=item code, "
     "H=unit, P=basket value, Q:Z=decile 1-10 values) rather than the source's per-section "
     "column drift (documented bugs and inconsistencies in the source layout were not "
     "replicated). Most decile rows are a single spilled array formula (LAMBDA + XLOOKUP over "
     "an array of deciles 1-10) rather than one formula per cell."),
]
for title, font, body in README_PARAS:
    style_cell(ws_r, rr, 2, title, SUBSECTION_FONT)
    rr += 1
    c = style_cell(ws_r, rr, 2, body, BASE_FONT, align=Alignment(wrap_text=True, vertical="top"))
    ws_r.row_dimensions[rr].height = 15 * (len(body) // 95 + 1)
    rr += 2

wb._sheets = (
    [ws_r, ws_in, ws_out]
    + [wb[n] for n in ["HHSurvey", "HH_Elast", "ASPIRE", "WHOCooking", "GDPRatios", "IO_GTAP", "Price_Changes", "Mapping"]]
)
wb.active = 0

# ---------------------------------------------------------------------------
# Capture every formula cell's clean formula text for the VBA rebuild script
# (see generate_vba() below), then blank the cells and drop all defined
# names so the shipped .xlsx opens without a repair prompt -- see
# RebuildDistributionModule.bas / this folder's README for why.
# ---------------------------------------------------------------------------
FORMULA_LOG = []  # (sheet_name, coordinate, clean_formula_without_leading_=)
for sheet_name in ["Distribution_Inputs", "Distribution_Outputs"]:
    for row in wb[sheet_name].iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                FORMULA_LOG.append((sheet_name, cell.coordinate, cell.value[1:]))
print(f"Captured {len(FORMULA_LOG)} formula cells and {len(NAMED_RANGES)} named ranges for the VBA rebuild.")

# Excel's name namespace (defined names AND table displayNames) is CASE-
# INSENSITIVE. AddDistributionNames' Names.Add calls run inside On Error
# Resume Next, so a name that collides with an existing Excel Table name
# fails SILENTLY -- the LAMBDA function (or plain named range) simply never
# gets created, and every formula that calls it shows #REF! with no error
# ever surfacing anywhere. (This is exactly how the ELASTADJ / "ElastAdj"
# table collision went undetected until testing in real Excel.) Catch any
# such collision here, at build time, instead.
_table_lc = {t.lower(): t for t in TABLE_NAMES}
for _n in NAMED_RANGES:
    if _n.lower() in _table_lc and _table_lc[_n.lower()] != _n:
        raise AssertionError(f"name {_n!r} collides case-insensitively with table {_table_lc[_n.lower()]!r}")
if len(_table_lc) != len(TABLE_NAMES):
    _dupe = [t for t in TABLE_NAMES if TABLE_NAMES.count(t) > 1]
    raise AssertionError(f"duplicate table name(s): {sorted(set(_dupe))}")
print(f"Checked {len(NAMED_RANGES)} names against {len(TABLE_NAMES)} table names: no collisions.")


def vba_wrapped_string_expr(text, cont_indent="        ", max_chunk=150, max_line=800):
    """VBA string-literal expression for `text`, returned as a list of
    physical source lines: '"chunk1" & "chunk2" & _' / ... / last line with
    no trailing continuation. VBA caps a physical line at ~1023 chars, so a
    single long '"a" & "b" & "c"' run on one line (as opposed to across
    several continued lines) will fail to compile for any formula much
    longer than a couple hundred characters -- every one of our LAMBDA
    definitions and effect-summing formulas is well past that.

    Chunking happens on the RAW text, and each chunk is quote-escaped (" ->
    "") independently afterwards -- NOT the other way around. Escaping
    first and then slicing at a fixed character offset can land the cut
    between the two characters of a doubled "" escape pair, silently
    corrupting the literal from that point on (this was a real bug: it
    passed every syntax check because both halves are individually valid
    VBA text, it just decodes back to the wrong string). Chunking before
    escaping makes that class of bug structurally impossible, since a chunk
    boundary can only ever fall between two original characters, never
    inside the two-character encoding of one.

    The first returned line has no leading indent (the caller's statement
    prefix, e.g. 'wb.Names.Add Name:="X", RefersTo:=', goes immediately
    before it); continuation lines are indented with `cont_indent`."""
    chunks = [text[i:i + max_chunk] for i in range(0, len(text), max_chunk)] or [""]
    literals = [f'"{c.replace(chr(34), chr(34) * 2)}"' for c in chunks]
    out_lines = [literals[0]]
    for lit in literals[1:]:
        candidate = out_lines[-1] + " & " + lit
        if len(candidate) > max_line:
            out_lines[-1] += " & _"
            out_lines.append(cont_indent + lit)
        else:
            out_lines[-1] = candidate
    _verify_vba_string_roundtrip(text, chunks, out_lines)
    return out_lines


def _parse_vba_string_concat(source):
    """Parse VBA source of the form '"a" & "b" & "c"' (a chain of quoted
    literals joined by ' & ', doubled "" as the in-literal escaped quote)
    back into the single string it represents. A real character-by-
    character parser rather than string-replace, so it can't be confused by
    '&' or other characters that legitimately appear inside a literal
    (e.g. HHKEY's own '&' string-concatenation formula)."""
    i, n, out = 0, len(source), []
    while i < n:
        assert source[i] == '"', f"expected opening quote at {i}: {source[max(0, i-20):i+20]!r}"
        i += 1
        while True:
            if source[i] == '"':
                if i + 1 < n and source[i + 1] == '"':
                    out.append('"')
                    i += 2
                else:
                    i += 1
                    break
            else:
                out.append(source[i])
                i += 1
        if i < n:
            assert source[i:i + 3] == " & ", f"expected ' & ' at {i}: {source[i:i + 10]!r}"
            i += 3
    return "".join(out)


def _verify_vba_string_roundtrip(text, chunks, out_lines, cont_indent="        "):
    assert "".join(chunks) == text, "chunking changed the text"
    logical = out_lines[0]
    for line in out_lines[1:]:
        assert logical.endswith(" & _"), f"expected continuation marker before {line[:40]!r}"
        assert line.startswith(cont_indent), f"expected continuation indent before {line[:40]!r}"
        logical = logical[:-4] + " & " + line[len(cont_indent):]
    decoded = _parse_vba_string_concat(logical)
    if decoded != text:
        raise AssertionError(f"VBA string round-trip failed for text starting {text[:60]!r}")


def vba_assign_statement(prefix, text):
    """`prefix` (e.g. 'wb.Names.Add Name:="X", RefersTo:=') followed by a
    (possibly line-continued) VBA string-literal expression for `text`."""
    expr_lines = vba_wrapped_string_expr(text)
    return [prefix + expr_lines[0]] + expr_lines[1:]


def generate_vba():
    lines = []
    lines.append('Attribute VB_Name = "RebuildDistributionModule"')
    lines.append("' Rebuilds the named LAMBDA functions and all Distribution_Outputs /")
    lines.append("' Distribution_Inputs formulas via Excel's own object model, so Excel")
    lines.append("' itself performs the internal _xlfn/_xlpm encoding instead of relying")
    lines.append("' on formula text written directly into the .xlsx XML.")
    lines.append("' Usage: Alt+F11 -> Insert -> Module -> paste this file -> F5 (or run")
    lines.append("' RebuildDistributionModule from the macro list). Run once after opening")
    lines.append("' the workbook; safe to re-run.")
    lines.append("Option Explicit")
    lines.append("")
    lines.append("Sub RebuildDistributionModule()")
    lines.append("    Application.ScreenUpdating = False")
    lines.append("    Application.Calculation = xlCalculationManual")
    lines.append("    AddDistributionNames")

    # --- names (both plain ranges and LAMBDA-valued) ---
    lines.append("End Sub")
    lines.append("")
    lines.append("Sub AddDistributionNames()")
    lines.append("    Dim wb As Workbook: Set wb = ThisWorkbook")
    lines.append("    On Error Resume Next")
    for name, ref in NAMED_RANGES.items():
        formula = ref if ref.startswith("=") else "=" + ref
        lines.extend(vba_assign_statement(f'    wb.Names.Add Name:="{name}", RefersTo:=', formula))
    lines.append("    On Error GoTo 0")
    lines.append("End Sub")
    lines.append("")

    # --- formulas, chunked into ~120-statement Subs to stay well under the
    #     ~64K-character-per-procedure VBA limit ---
    CHUNK = 120
    part_names = []
    for i in range(0, len(FORMULA_LOG), CHUNK):
        part = i // CHUNK + 1
        part_name = f"WriteFormulas_Part{part}"
        part_names.append(part_name)
        lines.append(f"Sub {part_name}()")
        lines.append("    Dim ws As Worksheet")
        cur_sheet = None
        for sheet_name, coord, formula in FORMULA_LOG[i:i + CHUNK]:
            if sheet_name != cur_sheet:
                lines.append(f"    Set ws = ThisWorkbook.Worksheets(\"{sheet_name}\")")
                cur_sheet = sheet_name
            full_formula = "=" + formula
            lines.extend(vba_assign_statement(f'    ws.Range("{coord}").Formula2 = ', full_formula))
        lines.append("End Sub")
        lines.append("")

    # --- master sub, appended after AddDistributionNames call list ---
    master_idx = lines.index("Sub RebuildDistributionModule()")
    insert_idx = lines.index("    AddDistributionNames") + 1
    for pn in part_names:
        lines.insert(insert_idx, f"    {pn}")
        insert_idx += 1
    # re-find end and add calc/finish lines
    end_idx = lines.index("End Sub")
    lines.insert(end_idx, "    Application.Calculation = xlCalculationAutomatic")
    # A plain switch back to automatic calc isn't always enough to clear
    # names/formulas that were defined while calculation was suspended --
    # force Excel to fully rebuild the dependency tree and recalculate
    # everything from scratch (this is what fixed the #REF! errors on
    # ELASTADJ and everything downstream of it).
    lines.insert(end_idx + 1, "    Application.CalculateFullRebuild")
    lines.insert(end_idx + 2, "    Application.ScreenUpdating = True")
    lines.insert(end_idx + 3, '    MsgBox "Distribution module rebuilt: " & Names.Count & _' )
    lines.insert(end_idx + 4, '        " names, ' + str(len(FORMULA_LOG)) + ' formulas.", vbInformation')

    return "\n".join(lines) + "\n"


vba_path = os.path.join(HERE, "RebuildDistributionModule.bas")
with open(vba_path, "w", newline="\r\n") as f:
    f.write(generate_vba())
print(f"Wrote {vba_path}")

# ---------------------------------------------------------------------------
# Ship a CLEAN-opening .xlsx: no LAMBDA-valued (or other future-function)
# defined names, no formulas referencing them -- only these were ever the
# source of the "file needs repair" prompt. All row/column labels, section
# structure, styling and the data tabs are untouched; run the .bas macro
# above after opening to populate the LAMBDA names and every formula cell.
# Cells that would have held a formula are left blank with a short note so
# it's obvious the macro still needs to run.
# ---------------------------------------------------------------------------
for dn_name in list(wb.defined_names.keys()):
    del wb.defined_names[dn_name]
for sheet_name in ["Distribution_Inputs", "Distribution_Outputs"]:
    for row in wb[sheet_name].iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = None
style_cell(ws_r, 2, 2, 'Run RebuildDistributionModule.bas (Alt+F11 -> Insert -> Module -> paste -> F5) '
                       "to populate the LAMBDA names and every formula cell -- they are intentionally "
                       "left blank in this file so it opens without a repair prompt.", ITALIC_FONT)

wb.save(OUT)
print("Saved", OUT)
