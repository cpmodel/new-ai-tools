#!/usr/bin/env python3
# =============================================================================
# MITIGATION (NET ZERO) INVESTMENT NEEDS - workbook builder
# =============================================================================
# Builds InvestmentNeeds/Mitigation_Investment_Needs_v1.0.xlsx: a formula-driven
# Excel model of total and incremental mitigation investment needs, and their
# incremental cost, by country and sector, under three cost-of-capital lenses:
#   (a) market cost of capital (country WACC),
#   (b) World Bank rate (6% social discount rate, Fay et al. 2016),
#   (c) World Bank rate with local co-benefits.
#
# Method: "Global Investment Needs for Climate Action: A Method Based on
# SISEPUEDE" (Oct 2026), equations 1-7. This file implements its Tier 3
# (allocation of IHLEG/IEA group totals by country drivers), with the country
# cost of capital, savings and co-benefits differentiated from country data.
#
# Inputs (all read from the repo; no network):
#   TaxMacroFiscalDashboard/build/output/global_data_wide.csv
#       dashboard panel: cost-of-capital block (IRENA, Steffen, IEA, IDS) and
#       PIMA block (BOOST, PEFA, UNDP), plus WB region / income group
#   TaxMacroFiscalDashboard/build/output/imffossilfuelsubsidiesdata_sheets.xlsx
#       IMF CPAT panel (sheet "data"): emissions, energy use by fuel and sector,
#       supply costs, air-pollution and road externalities, 2024
#   TaxMacroFiscalDashboard/build/raw_data/dataset_..._IMF_RES_WEO_9_0_0.csv
#       IMF WEO: GDP (USD, PPP), real growth 2025-26
#   TaxMacroFiscalDashboard/build/raw_data/SteffenEtAl2025_WACC_database.csv
#       Steffen et al. (2025): median leverage for the Damodaran WACC build
#   InvestmentNeeds/raw_data/ctryprem.xlsx            (OPTIONAL)
#       Damodaran "Country Default Spreads and Risk Premiums". If present it is
#       loaded into the Damodaran sheet; if absent the sheet is left for pasting.
#
# Run:  python InvestmentNeeds/build/build_investment_model.py
# Then recalculate (Excel does this on open; or LibreOffice headless).
# Deps: pandas, numpy, openpyxl
# =============================================================================

import os, re, unicodedata
import numpy as np, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
DASH = os.path.join(ROOT, "TaxMacroFiscalDashboard", "build")
OUT_XLSX = os.path.join(ROOT, "InvestmentNeeds", "Mitigation_Investment_Needs_v1.1.xlsx")
DAM_F = os.path.join(ROOT, "InvestmentNeeds", "raw_data", "ctryprem.xlsx")
WEO_F = os.path.join(DASH, "raw_data",
                     "dataset_2026-08-26T13_41_09_874550254Z_DEFAULT_INTEGRATION_IMF_RES_WEO_9_0_0.csv")
VERSION = "v1.1"

# -----------------------------------------------------------------------------
# 1. DATA
# -----------------------------------------------------------------------------
wide = pd.read_csv(os.path.join(DASH, "output", "global_data_wide.csv"))
wide = wide.set_index("iso3")

# --- IMF WEO ---------------------------------------------------------------
weo = pd.read_csv(WEO_F, low_memory=False)
weo["iso"] = weo.SERIES_CODE.str.split(".").str[0].replace({"KOS": "XKX", "WBG": "PSE"})
weo = weo[weo.iso.str.len() == 3]
def weo_get(ind, year):
    s = weo[weo.INDICATOR == ind].groupby("iso")[year].first()
    return pd.to_numeric(s, errors="coerce")
GDPU = "Gross domestic product (GDP), Current prices, US dollar"
GDPP = ("Gross domestic product (GDP), Current prices, Purchasing power parity (PPP) "
        "international dollar, ICP benchmarks 2017-2021")
GRW = "Gross domestic product (GDP), Constant prices, Percent change"
gdp24, ppp24 = weo_get(GDPU, "2024"), weo_get(GDPP, "2024")
g25, g26 = weo_get(GRW, "2025"), weo_get(GRW, "2026")
weo_name = weo[weo.INDICATOR == GDPU].groupby("iso").COUNTRY.first()

# --- IMF CPAT panel (2024) ---------------------------------------------------
cp = pd.read_excel(os.path.join(DASH, "output", "imffossilfuelsubsidiesdata_sheets.xlsx"),
                   sheet_name="data")
cp.columns = ["country", "u", "label", "code", "iso", "inc", "reg", "mt", "sub",
              "y21", "y22", "y23", "y24"]
cp = cp.drop_duplicates(["iso", "mt"])          # scenario rows repeat the baseline
CP = cp.pivot(index="iso", columns="mt", values="y24")
cpat_name = cp.groupby("iso").country.first()
CPAT_COLS = [  # (column header, CPAT code, unit)
    ("Population (mn)", "mit.pop.mn", "mn"),
    ("CO2, energy (Mt)", "mit.co2.tot.1", "Mt"),
    ("GHG excl. LULUCF (MtCO2e)", "mit.ghg.tot.exc.1", "Mt"),
    ("GHG incl. LULUCF (MtCO2e)", "mit.ghg.tot.inc.1", "Mt"),
    ("Electricity, industry (TWh)", "mit.con.ecy.ind.1", "TWh"),
    ("Electricity, residential (TWh)", "mit.con.ecy.res.1", "TWh"),
    ("Electricity, other (TWh)", "mit.con.ecy.other.1", "TWh"),
    ("Gas, industry (EJ)", "mit.con.nga.ind.1", "EJ"),
    ("Gas, residential (EJ)", "mit.con.nga.res.1", "EJ"),
    ("Gas, power (EJ)", "mit.con.nga.pow.1", "EJ"),
    ("Gas, other (EJ)", "mit.con.nga.other.1", "EJ"),
    ("Coal, industry (EJ)", "mit.con.coa.ind.1", "EJ"),
    ("Coal, residential (EJ)", "mit.con.coa.res.1", "EJ"),
    ("Coal, power (EJ)", "mit.con.coa.pow.1", "EJ"),
    ("Coal, other (EJ)", "mit.con.coa.other.1", "EJ"),
    ("Gasoline (bn litres)", "mit.con.gso.all.1", "bn l"),
    ("Diesel (bn litres)", "mit.con.die.all.1", "bn l"),
    ("LPG (bn litres)", "mit.con.lpg.all.1", "bn l"),
    ("Kerosene (bn litres)", "mit.con.ker.all.1", "bn l"),
    ("Other oil products (bn litres)", "mit.con.oop.all.1", "bn l"),
    ("Supply cost gasoline ($/l)", "mit.sup.cost.gso.all.1", "$/l"),
    ("Supply cost diesel ($/l)", "mit.sup.cost.die.all.1", "$/l"),
    ("Supply cost LPG ($/l)", "mit.sup.cost.lpg.all.1", "$/l"),
    ("Supply cost kerosene ($/l)", "mit.sup.cost.ker.all.1", "$/l"),
    ("Supply cost other oil ($/l)", "mit.sup.cost.oop.all.1", "$/l"),
    ("Supply cost coal, industry ($/GJ)", "mit.sup.cost.coa.ind.1", "$/GJ"),
    ("Supply cost coal, residential ($/GJ)", "mit.sup.cost.coa.res.1", "$/GJ"),
    ("Supply cost coal, power ($/GJ)", "mit.sup.cost.coa.pow.1", "$/GJ"),
    ("Supply cost coal, other ($/GJ)", "mit.sup.cost.coa.other.1", "$/GJ"),
    ("Supply cost gas, industry ($/GJ)", "mit.sup.cost.nga.ind.1", "$/GJ"),
    ("Supply cost gas, residential ($/GJ)", "mit.sup.cost.nga.res.1", "$/GJ"),
    ("Supply cost gas, power ($/GJ)", "mit.sup.cost.nga.pow.1", "$/GJ"),
    ("Supply cost gas, other ($/GJ)", "mit.sup.cost.nga.other.1", "$/GJ"),
    ("Local air pollution cost (US$ bn)", "mit.airpol.con.all.all.1", "$bn"),
    ("Congestion cost (US$ bn)", "mit.con.con.all.all.1", "$bn"),
    ("Road accident cost (US$ bn)", "mit.roda.con.all.all.1", "$bn"),
    ("Road damage cost (US$ bn)", "mit.rodd.con.all.all.1", "$bn"),
    ("Renewable share of electricity (CPAT)", "mit.ecy.renshare.1", "share"),
    ("Supply cost electricity, residential ($/kWh)", "mit.sup.cost.ecy.res.1", "$/kWh"),
    ("Air pollution cost coal, power ($/GJ)", "mit.airpol.cost.coa.pow.1", "$/GJ"),
    ("Air pollution cost gas, power ($/GJ)", "mit.airpol.cost.nga.pow.1", "$/GJ"),
    ("Air pollution cost other oil ($/l)", "mit.airpol.cost.oop.1", "$/l"),
    ("Air pollution cost gasoline ($/l)", "mit.airpol.cost.gso.all.1", "$/l"),
]

# --- Ember: generation by fossil source, latest year -------------------------
em = pd.read_csv(os.path.join(DASH, "raw_data", "release_generation_yearly_global.csv"), low_memory=False)
em = em[(em["Area type"] == "Country or economy") & em["ISO 3 code"].notna()]
em_last = em.groupby("ISO 3 code").Year.max()
em = em[em.Year == em["ISO 3 code"].map(em_last)]
EMB = em.pivot_table(index="ISO 3 code", columns="Electricity source", values="Generation (TWh)", aggfunc="sum")
EMB["year"] = em_last
EMBER_COLS = [("Ember year", "year"), ("Generation, total (TWh), Ember", "Total generation"),
              ("Coal (TWh), Ember", "Coal"), ("Gas (TWh), Ember", "Gas"), ("Other fossil (TWh), Ember", "Other fossil")]
for _, c_ in EMBER_COLS:
    assert c_ in EMB.columns, c_

# --- Dashboard indicators carried into the model ------------------------------
DASH_COLS = [  # (header, panel code, block)
    ("Gross debt, gen. govt (% GDP)", "gross_debt_pct_gdp", "Fiscal"),
    ("Interest-to-revenue (%)", "interest_to_revenue_pct", "Fiscal"),
    ("Capex execution rate (%)", "capex_execution_rate_pct", "PIMA"),
    ("PEFA PI-11 PIM score (1-4)", "pefa_pi11_pim_score", "PIMA"),
    ("PEFA PI-12.2 asset monitoring (1-4)", "pefa_pi122_asset_monitoring", "PIMA"),
    ("Climate-PFM diagnostic on record (0/1)", "climate_pfm_diagnostic_on_record", "PIMA"),
    ("WACC solar, real after-tax (%), IRENA 2021", "wacc_solar_real_aftertax_pct", "CoC"),
    ("WACC nominal after-tax (%), Steffen et al.", "wacc_steffen_nominal_aftertax_pct", "CoC"),
    ("Steffen observation year", "wacc_steffen_nominal_aftertax_pct__year", "CoC"),
    ("WACC evidence strength (0-3)", "wacc_n_sources", "CoC"),
    ("IEA CoC Observatory (0/1)", "wacc_iea_observatory_available", "CoC"),
    ("Official-private rate spread (pp), IDS", "ids_official_private_rate_spread_pp", "CoC"),
    ("Private creditor rate (%), IDS", "ids_private_creditor_rate_pct", "CoC"),
    ("Avg maturity new commitments (yrs), IDS", "ids_avg_maturity_years", "CoC"),
    ("Mitigation finance received (% GDP), OECD", "mitigation_finance_received_pct_gdp", "Finance"),
    ("PPI commitments (% GDP)", "ppi_commitments_pct_gdp", "Finance"),
]

# --- IMF advanced economies (WEO classification, 41 economies) -------------
IMF_AE = set("""AND AUS AUT BEL CAN HRV CYP CZE DNK EST FIN FRA DEU GRC HKG ISL IRL ISR ITA
JPN KOR LVA LTU LUX MAC MLT NLD NZL NOR PRT PRI SMR SGP SVK SVN ESP SWE CHE TWN GBR USA""".split())

# --- Country universe: every economy with WEO GDP 2024 -----------------------
iso_all = sorted(i for i in gdp24.dropna().index)
def name_of(i):
    if i in wide.index: return wide.loc[i, "country"]
    if i in cpat_name.index: return cpat_name[i]
    return weo_name.get(i, i)
rows = []
for i in iso_all:
    region = wide.loc[i, "region"] if i in wide.index else ("East Asia & Pacific" if i == "TWN" else "")
    income = wide.loc[i, "income"] if i in wide.index else ("High income" if i == "TWN" else "")
    grp = "AE" if i in IMF_AE else ("China" if i == "CHN" else "EMDC")
    rows.append(dict(iso=i, name=name_of(i), region=region, income=income, group=grp))
C = pd.DataFrame(rows)
C["gdp24"] = C.iso.map(gdp24)
C = C.sort_values("gdp24", ascending=False).reset_index(drop=True)
N = len(C)

# --- Damodaran (optional) ----------------------------------------------------
def _norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z]", "", s)
ALIASES = {}            # normalised name -> iso3
for i in C.iso:
    for n in (name_of(i), weo_name.get(i), cpat_name.get(i)):
        if isinstance(n, str): ALIASES[_norm(n)] = i
# Common spellings in market-data sources (Damodaran and others) that differ from WDI/WEO.
EXTRA = {"Korea": "KOR", "South Korea": "KOR", "Korea, Republic of": "KOR",
 "Congo (Democratic Republic of)": "COD", "Democratic Republic of Congo": "COD", "Congo, DR": "COD",
 "Congo (Republic of)": "COG", "Republic of Congo": "COG", "Congo": "COG",
 "Ivory Coast": "CIV", "Cote d'Ivoire": "CIV", "Côte d'Ivoire": "CIV",
 "Egypt": "EGY", "Iran": "IRN", "Russia": "RUS", "Venezuela": "VEN", "Vietnam": "VNM",
 "Viet Nam": "VNM", "Laos": "LAO", "Kyrgyzstan": "KGZ", "Slovakia": "SVK", "Turkey": "TUR",
 "Türkiye": "TUR", "Turkiye": "TUR", "Syria": "SYR", "Yemen": "YEM", "Gambia": "GMB",
 "The Gambia": "GMB", "Bahamas": "BHS", "The Bahamas": "BHS", "Hong Kong": "HKG",
 "Macao": "MAC", "Macau": "MAC", "Taiwan": "TWN", "Micronesia": "FSM", "St. Lucia": "LCA",
 "Saint Lucia": "LCA", "St. Vincent & the Grenadines": "VCT", "Saint Vincent and the Grenadines": "VCT",
 "St. Kitts & Nevis": "KNA", "Saint Kitts and Nevis": "KNA", "Cape Verde": "CPV", "Cabo Verde": "CPV",
 "Swaziland": "SWZ", "Eswatini": "SWZ", "Macedonia": "MKD", "North Macedonia": "MKD",
 "Czech Republic": "CZE", "Czechia": "CZE", "Bosnia and Herzegovina": "BIH", "Bosnia & Herzegovina": "BIH",
 "Trinidad and Tobago": "TTO", "Trinidad & Tobago": "TTO", "Antigua & Barbuda": "ATG",
 "Brunei": "BRN", "Palestine": "PSE", "West Bank and Gaza": "PSE", "Kosovo": "XKX",
 "United States": "USA", "United States of America": "USA", "United Kingdom": "GBR",
 "UK": "GBR", "USA": "USA", "Moldova": "MDA", "Tanzania": "TZA", "Bolivia": "BOL",
 "Myanmar": "MMR", "Burma": "MMR", "East Timor": "TLS", "Timor-Leste": "TLS",
 "Sao Tome and Principe": "STP", "Guinea-Bissau": "GNB", "Papua New Guinea": "PNG",
 "Abu Dhabi": "ARE", "United Arab Emirates": "ARE", "Curacao": "CUW", "Aruba": "ABW"}
for k, v in EXTRA.items(): ALIASES[_norm(k)] = v

dam_rows, dam_note = [], "No ctryprem.xlsx found in InvestmentNeeds/raw_data/ - paste the data here (see note)."
if os.path.exists(DAM_F):
    xl = pd.ExcelFile(DAM_F)
    for sh in xl.sheet_names:
        raw = pd.read_excel(DAM_F, sheet_name=sh, header=None)
        hdr = None
        for r in range(min(30, len(raw))):
            vals = [str(x).lower() for x in raw.iloc[r].tolist()]
            if any("country" == v.strip() for v in vals) and any("default spread" in v for v in vals):
                hdr = r; break
        if hdr is None: continue
        df = pd.read_excel(DAM_F, sheet_name=sh, header=hdr)
        cols = {c: str(c).lower() for c in df.columns}
        pick = lambda *keys: next((c for c, l in cols.items() if all(k in l for k in keys)), None)
        c_cty, c_rat = pick("country"), pick("rating")
        c_ds, c_crp = pick("default spread"), pick("country risk premium")
        c_erp, c_tax = pick("equity risk premium"), pick("tax")
        for _, r in df.iterrows():
            if pd.isna(r[c_cty]) or pd.isna(r.get(c_ds)): continue
            dam_rows.append(dict(name=str(r[c_cty]).strip(),
                                 rating=r[c_rat] if c_rat else None,
                                 ds=r[c_ds], crp=r[c_crp] if c_crp else None,
                                 erp=r[c_erp] if c_erp else None,
                                 tax=r[c_tax] if c_tax else None))
        dam_note = f"Loaded from ctryprem.xlsx, sheet '{sh}' ({len(dam_rows)} rows)."
        break

# --- Steffen et al. median leverage ------------------------------------------
st = pd.read_csv(os.path.join(DASH, "raw_data", "SteffenEtAl2025_WACC_database.csv"))
lev = pd.to_numeric(st[st.Variable == "Leverage (debt share)"].Value.astype(str)
                    .str.replace("%", ""), errors="coerce")
DV_MED = float(np.nanmedian(lev)) / 100

# -----------------------------------------------------------------------------
# 2. STYLES
# -----------------------------------------------------------------------------
F = "Arial"
f_base = Font(name=F, size=9)
f_in = Font(name=F, size=9, color="0000FF")
f_link = Font(name=F, size=9, color="008000")
f_b = Font(name=F, size=9, bold=True)
f_h = Font(name=F, size=9, bold=True, color="FFFFFF")
f_t = Font(name=F, size=13, bold=True, color="1F3864")
f_st = Font(name=F, size=10, bold=True, color="1F3864")
f_note = Font(name=F, size=8, italic=True, color="595959")
fill_h = PatternFill("solid", fgColor="1F3864")
fill_sub = PatternFill("solid", fgColor="D9E1F2")
fill_key = PatternFill("solid", fgColor="FFFF00")
fill_out = PatternFill("solid", fgColor="E2EFDA")
fill_grey = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
PCT, PCT2 = '0.0%;(0.0%);"-"', '0.00%;(0.00%);"-"'
BN, BN1 = '#,##0;(#,##0);"-"', '#,##0.0;(#,##0.0);"-"'
DEC3 = '0.000;(0.000);"-"'
DEC2 = '0.00;(0.00);"-"'

wb = Workbook()
NAMES = {}
def name(key, ws, cell):
    col, row = re.match(r"([A-Z]+)(\d+)", cell).groups()
    ref = f"'{ws.title}'!${col}${row}"
    wb.defined_names[key] = DefinedName(key, attr_text=ref)
    NAMES[key] = ref

def put(ws, cell, v, font=f_base, fmt=None, fill=None, al=None, b=None):
    c = ws[cell]; c.value = v; c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if al: c.alignment = al
    if b: c.border = b
    return c

def header_row(ws, r, c0, labels, widths=None, height=36):
    for j, lab in enumerate(labels):
        c = ws.cell(row=r, column=c0 + j, value=lab)
        c.font, c.fill, c.alignment, c.border = f_h, fill_h, center, box
        if widths: ws.column_dimensions[CL(c0 + j)].width = widths[j]
    ws.row_dimensions[r].height = height

# -----------------------------------------------------------------------------
# 3. INPUTS
# -----------------------------------------------------------------------------
wsI = wb.active; wsI.title = "Inputs"
wsR = wb.create_sheet("README", 0)
for col, w in zip("ABCDEFG", (34, 46, 12, 12, 70, 8, 4)):
    wsI.column_dimensions[col].width = w
put(wsI, "A1", "Inputs and scenario levers", f_t)
put(wsI, "A2", "Blue = hard-coded input. Yellow fill = key assumption to review. "
    "† = not verified against the primary source, or an author's assumption. "
    "All money in USD billion, 2023 prices (IHLEG basis), unless stated.", f_note)
header_row(wsI, 4, 1, ["Key (named cell)", "Parameter", "Value", "Unit", "Source / note", "Fit"])
INPUTS = [
 ("H", "Investment pools, 2030 (mitigation only)"),
 ("p_ce_world", "Clean energy investment need, world, 2030", 4800, "USD bn/yr",
  "IHLEG 3 (Bhattacharya et al. 2024), clean energy from IEA; method note Table 4", "G", False),
 ("p_ce_emdc", "Clean energy, EMDCs other than China", 1600, "USD bn/yr",
  "IHLEG 3 (2024): EMDCs ex-China clean energy USD 1.6 tn; method note Table 4", "G", False),
 ("p_ce_chn", "Clean energy, China", 850, "USD bn/yr",
  "Derived: IEA-IFC (2023) EMDE 2.2-2.8 tn less EMDE ex-China 1.4-1.9 tn (midpoints 2.5 - 1.65). †", "M", True),
 ("p_ce_ae", "Clean energy, advanced economies (residual)", "=p_ce_world-p_ce_emdc-p_ce_chn", "USD bn/yr",
  "Residual of the three lines above", "M", False),
 ("p_nat_emdc", "Nature and agriculture (AFOLU), EMDCs ex-China", 350, "USD bn/yr",
  "IHLEG 3 (2024): nature and agriculture USD 0.3-0.4 tn, midpoint. Includes some non-mitigation nature spending. †", "M", True),
 ("p_nbs_world", "Nature-based solutions, world, 2030", 399, "USD bn/yr",
  "UNEP State of Finance for Nature (2021): NbS flows USD 133 bn/yr must triple by 2030 (133 x 3)", "M", False),
 ("p_nat_ae", "AFOLU, advanced economies", "=MAX(0,p_nbs_world-p_nat_emdc)*p_ce_ae/(p_ce_ae+p_ce_chn)", "USD bn/yr",
  "World NbS less EMDC, split AE/China pro rata to clean energy. †", "P", True),
 ("p_nat_chn", "AFOLU, China", "=MAX(0,p_nbs_world-p_nat_emdc)*p_ce_chn/(p_ce_ae+p_ce_chn)", "USD bn/yr",
  "As above. †", "P", True),
 ("H", "Sector split of clean energy (applied to every group; edit per group on Sector_Params)"),
 ("p_sh_gen", "Low-emission power generation", 0.35, "share",
  "IEA-IFC (2023), EMDEs early 2030s: 'just over a third' to low-emission generation", "M", False),
 ("p_sh_grid", "Grids and storage", 0.24, "share", "IEA-IFC (2023): 'just under a quarter'", "M", False),
 ("p_sh_end", "Efficiency and end use", 0.33, "share", "IEA-IFC (2023): 'another third'", "M", False),
 ("p_sh_fuel", "Low-emission fuels and fuel supply", 0.08, "share", "IEA-IFC (2023): 'around 8%'", "M", False),
 ("p_spl_trn", "  of end use: transport", 0.45, "share of end use",
  "Author's assumption: EVs and charging are the largest end-use item. †", "P", True),
 ("p_spl_bld", "  of end use: buildings", 0.35, "share of end use", "Author's assumption. †", "P", True),
 ("p_spl_ind", "  of end use: industry", "=1-p_spl_trn-p_spl_bld", "share of end use", "Residual", "P", False),
 ("H", "Total vs incremental"),
 ("p_inc_ae", "Incremental share of total, advanced economies", 0.38, "share",
  "McKinsey (2022): +USD 3.5 tn over today on USD 9.2 tn gross (3.5/9.2). Method note Table 5", "M", True),
 ("p_inc_chn", "Incremental share of total, China", 0.38, "share", "As AE (McKinsey global ratio). †", "M", True),
 ("p_inc_emdc", "Incremental share of total, EMDCs", 0.24, "share",
  "CCDR synthesis USD 574 bn/yr incremental vs IHLEG USD 2.4 tn total (about a quarter). Method note Table 5", "M", True),
 ("H", "Cost of capital (lens a: market)"),
 ("p_r_ae", "Market WACC, advanced economies (default)", 0.055, "rate",
  "IEA via IHLEG (2024): solar WACC 5-6% in AEs. Used where no country estimate exists", "M", True),
 ("p_r_chn", "Market WACC, China (default)", 0.06, "rate", "Method note: assumed close to AEs. †", "M", True),
 ("p_r_emdc", "Market WACC, EMDCs (default)", 0.105, "rate",
  "IEA via IHLEG (2024): 9-12% in EMDCs, midpoint", "M", True),
 ("p_rh_ae", "High case WACC, advanced economies", 0.07, "rate", "Method note lens (a) high case", "M", False),
 ("p_rh_chn", "High case WACC, China", 0.08, "rate", "Method note lens (a) high case", "M", False),
 ("p_rh_emdc", "High case WACC, EMDCs", 0.15, "rate", "Method note lens (a) high case", "M", False),
 ("p_use_obs", "Use observed country WACC (IRENA / Steffen)? 1=yes", 1, "switch",
  "Dashboard Finance 1 block. Average of the available observations", "", False),
 ("p_use_dam", "Use Damodaran-based WACC where no observation? 1=yes", 1, "switch",
  "Fills coverage gaps from the Damodaran sheet (equation 5)", "", False),
 ("p_calib", "Calibrate Damodaran WACC to observed level? 1=yes", 1, "switch",
  "Shifts the structural estimate by the mean gap to observed WACC over countries with both", "", False),
 ("p_basis", "Rate basis: 0 = nominal (as published), 1 = real", 0, "switch",
  "IHLEG, IEA, Steffen and the 6% WB rate are compared as published in the method note (0). "
  "Set 1 to deflate market rates to real terms with the inflation input", "", True),
 ("p_infl", "Inflation for real/nominal conversion (USD)", 0.02, "rate",
  "Assumption: long-run USD inflation. Also converts IRENA real WACC to nominal. †", "M", False),
 ("H", "Damodaran WACC build (equation 5) - used only where Damodaran data are present"),
 ("p_rf", "Risk-free rate (US 10-year Treasury)", 0.042, "rate",
  "Placeholder: approximate 2025-26 level. Update to the date of the Damodaran file. †", "M", True),
 ("p_erp", "Mature-market equity risk premium", 0.042, "rate",
  "Damodaran, July 2026 update (4.20%), as reported by BVWire 22 Jul 2026. Check against ctryprem.xlsx", "M", False),
 ("p_dv", "Debt share D/V, renewable projects", round(DV_MED, 3), "share",
  "Steffen et al. (2025) database, median 'Leverage (debt share)', all observations", "G", False),
 ("p_beta", "Equity beta on mature ERP", 1.0, "x", "Assumption: average-risk project equity", "M", False),
 ("p_lambda", "Exposure to country risk premium (lambda)", 1.0, "x",
  "Damodaran default: full exposure to CRP", "M", False),
 ("p_tax_def", "Corporate tax rate where Damodaran gives none", 0.25, "rate", "Assumption. †", "M", False),
 ("p_shift", "Calibration shift applied to Damodaran WACC (computed)", None, "rate",
  "Mean(observed) less mean(Damodaran) over countries with both; 0 if switch off or no overlap", "", False),
 ("H", "Lens (b) and (c): World Bank rate, savings and co-benefits"),
 ("p_r_wb", "World Bank social discount rate", 0.06, "rate",
  "Fay et al. (2016), World Bank guidance note; Beyond the Gap and IMF GFSR use 6%", "G", False),
 ("p_r_ibrd", "IBRD lending rate (memo)", 0.052, "rate",
  "Kharas & Rivard (2022): effective IBRD 10-year flexible loan ~5.2%", "M", False),
 ("p_use_ibrd", "Use IBRD lending rate instead of 6% in lens (b)/(c)? 1=yes", 0, "switch",
  "Alternative reading of 'World Bank cost of capital'", "", False),
 ("p_b", "Financial savings per USD of capital (b)", 0.133, "per yr",
  "Calibrated on SISEPUEDE LAC: BCR 1.70 at 6% x CRF 0.078. Method note Table 3", "M", True),
 ("p_wb_idx", "Weight on country fuel-cost index for b (0-1)", 1, "weight",
  "1 = scale b by country fossil supply cost per tCO2 (CPAT) relative to world; 0 = uniform b", "", False),
 ("p_c", "Local co-benefits per USD of capital (c)", 0.129, "per yr",
  "Calibrated on SISEPUEDE LAC: co-benefits 1.65 x cost at 6%. Method note Table 3", "M", True),
 ("p_wc_idx", "Weight on country co-benefit index for c (0-1)", 1, "weight",
  "1 = scale c by country local externality cost per tCO2 (CPAT) relative to world; 0 = uniform c", "", False),
 ("p_road", "Include road externalities in co-benefit index? 1=yes", 0, "switch",
  "0 = local air pollution only. 1 adds congestion, accidents and road damage (SISEPUEDE counts them; EVs alone do not cut them)", "", False),
 ("p_idx_min", "Index floor", 0.25, "x", "Guards small economies with tiny CO2 denominators", "", False),
 ("p_idx_max", "Index cap", 4, "x", "As above", "", False),
 ("p_a", "Abatement per USD 1,000 of capital (a)", 0.35, "tCO2e/yr",
  "Author's calculation, ~37 Gt over ~USD 110 tn (IEA NZE). Method note Table 3. †", "P", True),
 ("p_spc_lo", "Shadow price of carbon 2030, low", 50, "USD/tCO2e", "World Bank (2017, 2024)", "G", False),
 ("p_spc_hi", "Shadow price of carbon 2030, high", 100, "USD/tCO2e", "World Bank (2017, 2024)", "G", False),
 ("H", "Sector benefit methods (v1.1): engineering for power and transport, cost per tonne for fuels and AFOLU"),
 ("p_eng", "Use engineering savings for power and transport? 1=yes", 1, "switch",
  "1 = derive b, c, a from country fuel mix, prices and externalities (CPAT, Ember). 0 = v1.0 calibrated ratios", "", False),
 ("p_grid_credit", "Share of power-generation benefits credited to grids", 0, "weight",
  "0 = grids are enabling capital with no savings of their own (conservative). 1 = same per-$ benefits as generation", "", True),
 ("p_capex_pv", "Utility solar PV capital cost", 760, "USD/kW",
  "IRENA (2024) Renewable Power Generation Costs in 2023, global weighted average (method note Table C). †", "G", False),
 ("p_capex_wind", "Onshore wind capital cost", 1160, "USD/kW", "IRENA (2024), as above. †", "G", False),
 ("p_sh_pv", "Solar share of new generation investment", 0.5, "share", "Author's assumption. †", "M", True),
 ("p_cf_pv", "Solar PV capacity factor", 0.17, "share", "Engineering norm, global average (IRENA range 0.15-0.18). †", "M", False),
 ("p_cf_wind", "Onshore wind capacity factor", 0.35, "share", "Engineering norm, global average (IRENA range 0.3-0.4). †", "M", False),
 ("p_om_re", "Renewable O&M, share of capital a year", 0.015, "share", "Engineering norm (~USD 10-20/kW-yr). †", "M", False),
 ("p_hr_coal", "Heat rate, coal power", 10.0, "GJ/MWh", "Engineering norm, ~36% efficiency. †", "M", False),
 ("p_hr_gas", "Heat rate, gas power", 7.5, "GJ/MWh", "Engineering norm, fleet ~48% efficiency. †", "M", False),
 ("p_hr_oil", "Heat rate, oil power", 10.5, "GJ/MWh", "Engineering norm, ~34% efficiency. †", "M", False),
 ("p_ef_coal", "CO2 factor, coal", 0.0946, "tCO2/GJ", "IPCC 2006 Guidelines Vol. 2 Table 2.2, other bituminous coal", "G", False),
 ("p_ef_gas", "CO2 factor, natural gas", 0.0561, "tCO2/GJ", "IPCC 2006, Table 2.2", "G", False),
 ("p_ef_oil", "CO2 factor, residual fuel oil", 0.0774, "tCO2/GJ", "IPCC 2006, Table 2.2", "G", False),
 ("p_ef_gso", "CO2 factor, motor gasoline", 0.0693, "tCO2/GJ", "IPCC 2006, Table 2.2", "G", False),
 ("p_gj_l_oil", "Energy content, fuel oil", 0.0386, "GJ/litre", "Physical constant (approx.)", "G", False),
 ("p_gj_l_gso", "Energy content, gasoline", 0.0342, "GJ/litre", "Physical constant (approx.)", "G", False),
 ("p_ev_prem", "EV purchase premium incl. charging, 2030", 5000, "USD/vehicle",
  "Author's assumption: SISEPUEDE 2019 value USD 13,000 falling; IEA reports near parity in China by 2024. †", "P", True),
 ("p_ev_km", "Vehicle-km a year", 15000, "km", "SISEPUEDE (Baik et al. 2019)", "M", False),
 ("p_ice_l", "Gasoline car fuel use", 7.5, "l/100 km", "Engineering norm, on-road fleet average. †", "M", False),
 ("p_ev_kwh", "EV electricity use", 0.18, "kWh/km", "Engineering norm incl. charging losses. †", "M", False),
 ("p_ev_maint", "EV maintenance saving", 0.012, "USD/km", "SISEPUEDE (AAA 2019)", "M", False),
 ("p_cpt_fuel", "Abatement cost, low-emission fuels and CCS", 100, "USD/tCO2e",
  "IPCC AR6 WGIII SPM C.12: options at USD 100/t or less could halve 2019 emissions by 2030. Used as levelised cost. †", "M", True),
 ("p_cpt_afolu", "Abatement cost, AFOLU", 50, "USD/tCO2e",
  "IPCC AR6 WGIII SPM C.9.1: AFOLU potential 8-14 Gt/yr below USD 100/t, 30-50% of it below USD 20/t. 50 is an assumed average. †", "P", True),
 ("H", "Allocation of group totals to countries (equation 1 / 3)"),
 ("p_beta_alloc", "Weight on physical driver vs GDP (PPP)", 0.5, "weight",
  "Author's assumption: half current physical activity (replacement), half economic size (growth). †", "P", True),
 ("p_phi", "Pass-through of local price level to capital cost", 0.5, "elasticity",
  "Author's assumption: about half of clean-energy capex is traded equipment. 0 = none, 1 = full. †", "P", True),
 ("H", "PIMA: public investment efficiency"),
 ("p_ps_ae", "Public share of investment, AEs", 0.30, "share", "Author's assumption. †", "P", True),
 ("p_ps_chn", "Public share of investment, China", 0.50, "share", "Author's assumption (SOE-led). †", "P", True),
 ("p_ps_emdc", "Public share of investment, EMDCs", 0.40, "share",
  "IEA-IFC (2023): about 60% of EMDE clean energy investment must be private", "M", False),
 ("p_gap_ae", "Efficiency gap, advanced economies", 0.13, "share", "IMF (2015) Making Public Investment More Efficient", "G", False),
 ("p_gap_em", "Efficiency gap, emerging markets", 0.27, "share", "IMF (2015)", "G", False),
 ("p_gap_lic", "Efficiency gap, low-income developing", 0.40, "share", "IMF (2015)", "G", False),
 ("p_use_pefa", "Use PEFA PI-11 score where available? 1=yes", 1, "switch",
  "Linear map: score 1 -> LIC gap, score 4 -> AE gap (author's mapping †)", "", False),
 ("p_reform", "PIMA reform scenario? 1=yes", 0, "switch",
  "IMF (2015): stronger PIM institutions can close up to two-thirds of the gap", "", False),
 ("p_reform_close", "Share of gap closed by reform", 2 / 3, "share", "IMF (2015)", "M", False),
]
r = 5
for it in INPUTS:
    if it[0] == "H":
        c = put(wsI, f"A{r}", it[1], f_st, fill=fill_sub)
        for col in "BCDEF": wsI[f"{col}{r}"].fill = fill_sub
        r += 1; continue
    key, lab, val, unit, src, fit, keyflag = it
    put(wsI, f"A{r}", key, f_note); put(wsI, f"B{r}", lab)
    is_f = isinstance(val, str) and val.startswith("=")
    fmt = PCT if unit in ("rate", "share", "share of end use", "weight") and key not in ("p_wb_idx", "p_wc_idx", "p_beta_alloc") else \
          (DEC3 if unit in ("per yr", "tCO2e/yr", "x", "elasticity", "weight") else (BN if "USD" in unit else "0"))
    if unit == "switch": fmt = "0"
    put(wsI, f"C{r}", val, f_base if is_f else f_in, fmt, fill_key if keyflag else None)
    put(wsI, f"D{r}", unit); put(wsI, f"E{r}", src, al=wrap); put(wsI, f"F{r}", fit, al=center)
    name(key, wsI, f"C{r}")
    r += 1
put(wsI, f"A{r+1}", "Switch cells take 0 or 1. Rates and shares are stored as fractions (0.06 = 6%).", f_note)

# -----------------------------------------------------------------------------
# 4. SECTOR PARAMETERS AND GROUP POOLS
# -----------------------------------------------------------------------------
SECT = [  # code, name, pool share formula, life, driver description, life source
 ("GEN", "Power generation (low-emission)", "p_sh_gen", 25, "Electricity demand (TWh)", "Method note: 25 yrs (solar 25-30, wind 25)"),
 ("GRID", "Grids and storage", "p_sh_grid", 40, "Electricity demand (TWh)", "Method note: grids 40 yrs (storage 15 - blended up)"),
 ("TRN", "Transport (EVs, charging, transit)", "p_sh_end*p_spl_trn", 12, "Road fuel energy (EJ)", "Method note: vehicles 12 yrs"),
 ("BLD", "Buildings (efficiency, heat pumps)", "p_sh_end*p_spl_bld", 25, "Building energy use (EJ)", "Method note: 20-30 yrs"),
 ("IND", "Industry (efficiency, electrification, CCS)", "p_sh_end*p_spl_ind", 25, "Industrial energy use (EJ)", "Method note: CCS 25 yrs"),
 ("FUEL", "Low-emission fuels and fuel supply", "p_sh_fuel", 20, "Fossil energy supply (EJ)", "Assumption †"),
 ("AFOLU", "AFOLU and nature", None, 30, "Non-CO2 GHG + net LULUCF source (MtCO2e)", "Assumption †"),
]
METHOD = {
 "GEN": "Engineering: displaced fossil power (Ember mix x heat rates x CPAT fuel cost, air pollution; IPCC CO2 factors)",
 "GRID": "Enabling: no own benefits; p_grid_credit x generation benefits",
 "TRN": "Engineering: EV vs gasoline car per vehicle (CPAT fuel and power prices, gasoline air pollution)",
 "BLD": "Calibrated: SISEPUEDE b and c x CPAT country index; a = p_a",
 "IND": "Calibrated: SISEPUEDE b and c x CPAT country index; a = p_a",
 "FUEL": "Cost per tonne: no savings or co-benefits; a = 1000 x CRF(6%) / p_cpt_fuel",
 "AFOLU": "Cost per tonne: no savings; c = p_c (ecosystem co-benefits); a = 1000 x CRF(6%) / p_cpt_afolu",
}
wsS = wb.create_sheet("Sector_Params")
put(wsS, "A1", "Sector parameters and 2030 investment pools by group (USD bn/yr)", f_t)
put(wsS, "A2", "Blue cells are editable. Incremental shares default to the group values on Inputs; "
    "overwrite a cell to set a sector-specific share. Technology premium is added to the country "
    "market WACC in lens (a) only.", f_note)
hdr = ["Code", "Sector", "Share of clean energy", "Asset life (yrs)", "Tech premium (lens a)",
       "Incr. share AE", "Incr. share China", "Incr. share EMDC", "Savings multiplier (b)",
       "Co-benefit multiplier (c)", "Benefit method (v1.1)", "Allocation driver", "Life source"]
header_row(wsS, 4, 1, hdr, [8, 36, 11, 9, 10, 9, 9, 9, 10, 10, 44, 34, 34])
SROW = {}
for k, (code, nm, sh, life, drv, lsrc) in enumerate(SECT):
    rr = 5 + k; SROW[code] = rr
    put(wsS, f"A{rr}", code, f_b); put(wsS, f"B{rr}", nm)
    put(wsS, f"C{rr}", f"={sh}" if sh else "n.a.", fmt=PCT)
    put(wsS, f"D{rr}", life, f_in, "0", fill_key if code in ("FUEL", "AFOLU") else None)
    put(wsS, f"E{rr}", 0, f_in, PCT)
    put(wsS, f"F{rr}", "=p_inc_ae", fmt=PCT); put(wsS, f"G{rr}", "=p_inc_chn", fmt=PCT)
    put(wsS, f"H{rr}", "=p_inc_emdc", fmt=PCT)
    put(wsS, f"I{rr}", 1, f_in, DEC2); put(wsS, f"J{rr}", 1, f_in, DEC2)
    put(wsS, f"K{rr}", METHOD[code], al=wrap)
    put(wsS, f"L{rr}", drv); put(wsS, f"M{rr}", lsrc)
rr = 5 + len(SECT)
put(wsS, f"B{rr}", "Total (energy sectors)", f_b)
put(wsS, f"C{rr}", f"=SUM(C5:C{rr-2})", f_b, PCT)

PR0 = rr + 3
put(wsS, f"A{PR0-1}", "Investment pools, 2030 (USD bn/yr, total investment)", f_st)
header_row(wsS, PR0, 1, ["Code", "Sector", "AE", "China", "EMDC", "World"])
GCOL = {"AE": "C", "China": "D", "EMDC": "E"}
PROW = {}
for k, (code, nm, sh, *_rest) in enumerate(SECT):
    rr = PR0 + 1 + k; PROW[code] = rr
    put(wsS, f"A{rr}", code, f_b); put(wsS, f"B{rr}", nm)
    if code == "AFOLU":
        put(wsS, f"C{rr}", "=p_nat_ae", fmt=BN); put(wsS, f"D{rr}", "=p_nat_chn", fmt=BN)
        put(wsS, f"E{rr}", "=p_nat_emdc", fmt=BN)
    else:
        put(wsS, f"C{rr}", f"=p_ce_ae*C{SROW[code]}", fmt=BN)
        put(wsS, f"D{rr}", f"=p_ce_chn*C{SROW[code]}", fmt=BN)
        put(wsS, f"E{rr}", f"=p_ce_emdc*C{SROW[code]}", fmt=BN)
    put(wsS, f"F{rr}", f"=SUM(C{rr}:E{rr})", f_b, BN)
rt = PR0 + 1 + len(SECT)
put(wsS, f"B{rt}", "Total", f_b)
for col in "CDEF": put(wsS, f"{col}{rt}", f"=SUM({col}{PR0+1}:{col}{rt-1})", f_b, BN)
POOL_HDR = f"Sector_Params!$C${PR0}:$E${PR0}"

# -----------------------------------------------------------------------------
# 5. DAMODARAN SHEET
# -----------------------------------------------------------------------------
wsD = wb.create_sheet("Damodaran")
put(wsD, "A1", "Damodaran: Country Default Spreads and Risk Premiums (input)", f_t)
put(wsD, "A2", dam_note, f_note)
put(wsD, "A3", "To update by hand: download ctryprem.xlsx from pages.stern.nyu.edu/~adamodar "
    "(Data > Current data > Risk/Discount rate > Country risk), sheet 'ERPs by country'. Paste the country "
    "name into column B and the rating, rating-based default spread, country risk premium, total equity "
    "risk premium and corporate tax rate into C:G as fractions. Column A finds the ISO3 code from the alias "
    "list (Lists sheet); fill it by hand where it shows blank. Or drop the file into InvestmentNeeds/raw_data/ "
    "and re-run the build script.", f_note, al=wrap)
wsD.merge_cells("A3:H3"); wsD.row_dimensions[3].height = 48
header_row(wsD, 5, 1, ["ISO3 (lookup)", "Country (as in Damodaran)", "Moody's rating",
                       "Rating-based default spread", "Country risk premium", "Total equity risk premium",
                       "Corporate tax rate", "Matched?"], [10, 30, 10, 12, 12, 12, 12, 9])
DAM_N = max(260, len(dam_rows) + 20)
for k in range(DAM_N):
    rr = 6 + k
    put(wsD, f"A{rr}", f'=IF(B{rr}="","",IFERROR(INDEX(Lists!$B:$B,MATCH(B{rr},Lists!$A:$A,0)),""))')
    put(wsD, f"H{rr}", f'=IF(B{rr}="","",IF(A{rr}="","NO","yes"))')
    if k < len(dam_rows):
        d = dam_rows[k]
        put(wsD, f"B{rr}", d["name"], f_in)
        put(wsD, f"C{rr}", d["rating"], f_in)
        for col, key in (("D", "ds"), ("E", "crp"), ("F", "erp"), ("G", "tax")):
            v = d[key]
            if isinstance(v, str):
                v = pd.to_numeric(v.replace("%", ""), errors="coerce")
                v = v / 100 if pd.notna(v) else None
            put(wsD, f"{col}{rr}", None if v is None or (isinstance(v, float) and np.isnan(v)) else v, f_in, PCT2)
    else:
        for col in "BCDEFG": wsD[f"{col}{rr}"].font = f_in
        for col in "DEFG": wsD[f"{col}{rr}"].number_format = PCT2
DAM_RANGE = f"Damodaran!$A$6:$A${5+DAM_N}"

# Lists: name aliases for the Damodaran lookup
wsL = wb.create_sheet("Lists")
put(wsL, "A1", "Country name", f_b); put(wsL, "B1", "ISO3", f_b)
aliases_out = {}
for i in C.iso:
    for n in (name_of(i), weo_name.get(i), cpat_name.get(i)):
        if isinstance(n, str): aliases_out[n] = i
aliases_out.update(EXTRA)
for k, (n, i) in enumerate(sorted(aliases_out.items())):
    put(wsL, f"A{k+2}", n, f_in); put(wsL, f"B{k+2}", i, f_in)
wsL.column_dimensions["A"].width = 44

# -----------------------------------------------------------------------------
# 6. COUNTRY_DATA
# -----------------------------------------------------------------------------
wsC = wb.create_sheet("Country_Data")
put(wsC, "A1", "Country data (inputs from the dashboard panel, CPAT and WEO) and country parameters (formulas)", f_t)
put(wsC, "A2", "Blue = data loaded by the build script; black = formulas. Blank = not available (nothing imputed). "
    "Rates as fractions.", f_note)
cols = []  # (key, header, kind, value_fn or formula template, fmt, width)
def add(key, header, kind, src, fmt="General", width=11):
    cols.append((key, header, kind, src, fmt, width))
add("iso", "ISO3", "v", lambda r: r.iso, "General", 7)
add("name", "Country", "v", lambda r: r["name"], "General", 24)
add("region", "WB region", "v", lambda r: r.region, "General", 16)
add("income", "WB income group", "v", lambda r: r.income, "General", 14)
add("group", "IHLEG group (AE / China / EMDC)", "v", lambda r: r.group, "General", 9)
add("gdp24", "GDP 2024 (USD bn), WEO", "v", lambda r: gdp24.get(r.iso), BN1)
add("ppp24", "GDP 2024 (PPP int$ bn), WEO", "v", lambda r: ppp24.get(r.iso), BN1)
add("g25", "Real growth 2025 (%), WEO", "v", lambda r: g25.get(r.iso), DEC2)
add("g26", "Real growth 2026 (%), WEO", "v", lambda r: g26.get(r.iso), DEC2)
for h, code, u in CPAT_COLS:
    add(code, h + ", CPAT 2024", "v", (lambda code: (lambda r: CP.at[r.iso, code]
        if r.iso in CP.index and code in CP.columns and pd.notna(CP.at[r.iso, code]) else None))(code),
        DEC3 if u in ("$/l", "$/GJ", "share", "EJ") else BN1)
for h, code, blk in DASH_COLS:
    add(code, f"{h} [dashboard: {blk}]", "v", (lambda code: (lambda r: wide.at[r.iso, code]
        if r.iso in wide.index and pd.notna(wide.at[r.iso, code]) else None))(code),
        "0" if "year" in code else DEC2)
for h, c_ in EMBER_COLS:
    add("em_" + c_, h, "v", (lambda c_: (lambda r: float(EMB.at[r.iso, c_])
        if r.iso in EMB.index and pd.notna(EMB.at[r.iso, c_]) else None))(c_), "0" if c_ == "year" else BN1)
KEYS = [c[0] for c in cols]
def col_of(key): return CL([c[0] for c in cols].index(key) + 1)
# derived columns (formulas, {r} = row)
def X(key): return f"{col_of(key)}{{r}}"
def nz(key): return f'N({col_of(key)}{{r}})'          # N() turns blank into 0
add("gdp30", "GDP 2030 (USD bn, 2024 prices)", "f",
    f'={X("gdp24")}*(1+N({X("g25")})/100)*(1+N({X("g26")})/100)^5', BN1)
add("ppp_use", "GDP PPP used for weights", "f", f'=IF(N({X("ppp24")})>0,{X("ppp24")},{X("gdp24")})', BN1)
add("plevel", "Price level (GDP USD / GDP PPP)", "f", f'={X("gdp24")}/{X("ppp_use")}', DEC2)
add("has_cpat", "CPAT data? (1/0)", "f", f'=IF(ISNUMBER({X("mit.co2.tot.1")}),1,0)', "0")
add("elec", "Electricity demand (TWh)", "f",
    f'=IF({X("has_cpat")}=0,"",{nz("mit.con.ecy.ind.1")}+{nz("mit.con.ecy.res.1")}+{nz("mit.con.ecy.other.1")})', BN1)
add("d_trn", "Road fuel energy (EJ)", "f",
    f'=IF({X("has_cpat")}=0,"",({nz("mit.con.gso.all.1")}*34.2+{nz("mit.con.die.all.1")}*38.6)/1000)', DEC3)
add("d_bld", "Building energy use (EJ)", "f",
    f'=IF({X("has_cpat")}=0,"",{nz("mit.con.ecy.res.1")}*0.0036+{nz("mit.con.nga.res.1")}+{nz("mit.con.coa.res.1")}'
    f'+({nz("mit.con.lpg.all.1")}*25.3+{nz("mit.con.ker.all.1")}*37)/1000)', DEC3)
add("d_ind", "Industrial energy use (EJ)", "f",
    f'=IF({X("has_cpat")}=0,"",{nz("mit.con.ecy.ind.1")}*0.0036+{nz("mit.con.nga.ind.1")}+{nz("mit.con.coa.ind.1")}'
    f'+{nz("mit.con.nga.other.1")}+{nz("mit.con.coa.other.1")}+{nz("mit.con.oop.all.1")}*38/1000)', DEC3)
add("d_fuel", "Fossil energy supply (EJ)", "f",
    f'=IF({X("has_cpat")}=0,"",{nz("mit.con.nga.ind.1")}+{nz("mit.con.nga.res.1")}+{nz("mit.con.nga.pow.1")}+{nz("mit.con.nga.other.1")}'
    f'+{nz("mit.con.coa.ind.1")}+{nz("mit.con.coa.res.1")}+{nz("mit.con.coa.pow.1")}+{nz("mit.con.coa.other.1")}'
    f'+({nz("mit.con.gso.all.1")}*34.2+{nz("mit.con.die.all.1")}*38.6+{nz("mit.con.lpg.all.1")}*25.3'
    f'+{nz("mit.con.ker.all.1")}*37+{nz("mit.con.oop.all.1")}*38)/1000)', DEC3)
add("d_afolu", "Non-CO2 GHG + net LULUCF source (MtCO2e)", "f",
    f'=IF({X("has_cpat")}=0,"",MAX(0,{nz("mit.ghg.tot.exc.1")}-{nz("mit.co2.tot.1")})'
    f'+MAX(0,{nz("mit.ghg.tot.inc.1")}-{nz("mit.ghg.tot.exc.1")}))', BN1)
FUELPAIRS = [("mit.con.gso.all.1", "mit.sup.cost.gso.all.1"), ("mit.con.die.all.1", "mit.sup.cost.die.all.1"),
             ("mit.con.lpg.all.1", "mit.sup.cost.lpg.all.1"), ("mit.con.ker.all.1", "mit.sup.cost.ker.all.1"),
             ("mit.con.oop.all.1", "mit.sup.cost.oop.all.1"),
             ("mit.con.coa.ind.1", "mit.sup.cost.coa.ind.1"), ("mit.con.coa.res.1", "mit.sup.cost.coa.res.1"),
             ("mit.con.coa.pow.1", "mit.sup.cost.coa.pow.1"), ("mit.con.coa.other.1", "mit.sup.cost.coa.other.1"),
             ("mit.con.nga.ind.1", "mit.sup.cost.nga.ind.1"), ("mit.con.nga.res.1", "mit.sup.cost.nga.res.1"),
             ("mit.con.nga.pow.1", "mit.sup.cost.nga.pow.1"), ("mit.con.nga.other.1", "mit.sup.cost.nga.other.1")]
add("fuelexp", "Fossil fuel bill at supply cost (USD bn)", "f",
    f'=IF({X("has_cpat")}=0,"",' + "+".join(f'{nz(a)}*{nz(b)}' for a, b in FUELPAIRS) + ")", BN1)
add("co2_pos", "CO2 used in index (Mt)", "f",
    f'=IF(AND({X("has_cpat")}=1,N({X("mit.co2.tot.1")})>0),{X("mit.co2.tot.1")},0)', BN1)
add("fc_t", "Fuel bill per tCO2 (USD/t)", "f",
    f'=IF({X("co2_pos")}>0,{X("fuelexp")}*1000/{X("co2_pos")},"")', BN1)
add("lext", "Local externality cost (USD bn)", "f",
    f'=IF({X("has_cpat")}=0,"",{nz("mit.airpol.con.all.all.1")}+p_road*({nz("mit.con.con.all.all.1")}'
    f'+{nz("mit.roda.con.all.all.1")}+{nz("mit.rodd.con.all.all.1")}))', BN1)
add("cb_t", "Local co-benefit per tCO2 (USD/t)", "f",
    f'=IF({X("co2_pos")}>0,{X("lext")}*1000/{X("co2_pos")},"")', BN1)
add("idx_b", "Fuel-cost index (world = 1)", "f",
    f'=IF({X("fc_t")}="",1,MIN(p_idx_max,MAX(p_idx_min,{X("fc_t")}/p_fc_world)))', DEC2)
add("idx_c", "Co-benefit index (world = 1)", "f",
    f'=IF({X("cb_t")}="",1,MIN(p_idx_max,MAX(p_idx_min,{X("cb_t")}/p_cb_world)))', DEC2)
add("b_c", "Savings ratio b (per yr)", "f", f'=p_b*(1-p_wb_idx+p_wb_idx*{X("idx_b")})', DEC3)
add("c_c", "Co-benefit ratio c (per yr)", "f", f'=p_c*(1-p_wc_idx+p_wc_idx*{X("idx_c")})', DEC3)
# --- v1.1 engineering benefits: power generation (per USD 1,000 of capital) ---
add("fos_sh", "Fossil share of generation, Ember", "f",
    f'=IF(N({X("em_Total generation")})>0,(N({X("em_Coal")})+N({X("em_Gas")})+N({X("em_Other fossil")}))/{X("em_Total generation")},"")', PCT)
add("fos_tot", "helper: fossil generation (TWh)", "f",
    f'=N({X("em_Coal")})+N({X("em_Gas")})+N({X("em_Other fossil")})', BN1)
add("sh_coal", "Coal share of fossil generation", "f", f'=IF({X("fos_tot")}>0,N({X("em_Coal")})/{X("fos_tot")},"")', PCT)
add("sh_gas", "Gas share of fossil generation", "f", f'=IF({X("fos_tot")}>0,N({X("em_Gas")})/{X("fos_tot")},"")', PCT)
add("sh_oil", "Oil/other share of fossil generation", "f", f'=IF({X("fos_tot")}>0,N({X("em_Other fossil")})/{X("fos_tot")},"")', PCT)
add("re_kwh", "RE output per USD 1,000 (kWh/yr)", "f",
    "=1000*(p_sh_pv*p_cf_pv/p_capex_pv+(1-p_sh_pv)*p_cf_wind/p_capex_wind)*8760", BN)
add("gen_ok", "Engineering power data complete? (1/0)", "f",
    f'=IF(AND(ISNUMBER({X("fos_sh")}),ISNUMBER({X("sh_coal")}),{X("has_cpat")}=1),1,0)', "0")
add("fuel_mwh", "Fuel cost of displaced fossil power (USD/MWh)", "f",
    f'=IF({X("gen_ok")}=0,"",{X("sh_coal")}*p_hr_coal*N({X("mit.sup.cost.coa.pow.1")})+{X("sh_gas")}*p_hr_gas*N({X("mit.sup.cost.nga.pow.1")})'
    f'+{X("sh_oil")}*p_hr_oil*N({X("mit.sup.cost.oop.all.1")})/p_gj_l_oil)', BN1)
add("co2_mwh", "CO2 of displaced fossil power (t/MWh)", "f",
    f'=IF({X("gen_ok")}=0,"",{X("sh_coal")}*p_hr_coal*p_ef_coal+{X("sh_gas")}*p_hr_gas*p_ef_gas+{X("sh_oil")}*p_hr_oil*p_ef_oil)', DEC3)
add("ap_mwh", "Air pollution of displaced fossil power (USD/MWh)", "f",
    f'=IF({X("gen_ok")}=0,"",{X("sh_coal")}*p_hr_coal*N({X("mit.airpol.cost.coa.pow.1")})+{X("sh_gas")}*p_hr_gas*N({X("mit.airpol.cost.nga.pow.1")})'
    f'+{X("sh_oil")}*p_hr_oil*N({X("mit.airpol.cost.oop.1")})/p_gj_l_oil)', BN1)
add("grid_ef", "Average grid CO2 (t/MWh)", "f", f'=IF({X("gen_ok")}=0,"",{X("fos_sh")}*{X("co2_mwh")})', DEC3)
add("b_gen", "b power: fuel saved per $ less RE O&M", "f",
    f'=IF({X("gen_ok")}=0,"",{X("re_kwh")}*{X("fos_sh")}*{X("fuel_mwh")}/1000/1000-p_om_re)', DEC3)
add("c_gen", "c power: air pollution avoided per $", "f",
    f'=IF({X("gen_ok")}=0,"",{X("re_kwh")}*{X("fos_sh")}*{X("ap_mwh")}/1000/1000)', DEC3)
add("a_gen", "a power: tCO2/yr per $1,000", "f",
    f'=IF({X("gen_ok")}=0,"",{X("re_kwh")}*{X("fos_sh")}*{X("co2_mwh")}/1000)', DEC3)
# --- v1.1 engineering benefits: transport (EV premium) -----------------------
add("trn_ok", "Engineering transport data complete? (1/0)", "f",
    f'=IF(AND({X("gen_ok")}=1,ISNUMBER({X("mit.sup.cost.gso.all.1")}),ISNUMBER({X("mit.sup.cost.ecy.res.1")})),1,0)', "0")
add("ev_sav", "EV saving per vehicle (USD/yr)", "f",
    f'=IF({X("trn_ok")}=0,"",p_ev_km*p_ice_l/100*{X("mit.sup.cost.gso.all.1")}-p_ev_km*p_ev_kwh*{X("mit.sup.cost.ecy.res.1")}+p_ev_km*p_ev_maint)', BN)
add("b_trn", "b transport: saving per $ of premium", "f", f'=IF({X("trn_ok")}=0,"",{X("ev_sav")}/p_ev_prem)', DEC3)
add("c_trn", "c transport: air pollution avoided per $", "f",
    f'=IF({X("trn_ok")}=0,"",p_ev_km*p_ice_l/100*N({X("mit.airpol.cost.gso.all.1")})/p_ev_prem)', DEC3)
add("a_trn", "a transport: tCO2/yr per $1,000", "f",
    f'=IF({X("trn_ok")}=0,"",1000/p_ev_prem*(p_ev_km*p_ice_l/100*p_gj_l_gso*p_ef_gso-p_ev_km*p_ev_kwh/1000*{X("grid_ef")}))', DEC3)
# cost of capital
add("dam_ds", "Damodaran default spread", "f",
    f'=IFERROR(IF(INDEX(Damodaran!$D:$D,MATCH({X("iso")},Damodaran!$A:$A,0))="","",INDEX(Damodaran!$D:$D,MATCH({X("iso")},Damodaran!$A:$A,0))),"")', PCT2)
add("dam_crp", "Damodaran country risk premium", "f",
    f'=IFERROR(IF(INDEX(Damodaran!$E:$E,MATCH({X("iso")},Damodaran!$A:$A,0))="","",INDEX(Damodaran!$E:$E,MATCH({X("iso")},Damodaran!$A:$A,0))),"")', PCT2)
add("dam_tax", "Damodaran corporate tax", "f",
    f'=IFERROR(IF(INDEX(Damodaran!$G:$G,MATCH({X("iso")},Damodaran!$A:$A,0))="","",INDEX(Damodaran!$G:$G,MATCH({X("iso")},Damodaran!$A:$A,0))),"")', PCT2)
add("w_irena", "IRENA WACC, nominal", "f",
    f'=IF(ISNUMBER({X("wacc_solar_real_aftertax_pct")}),(1+{X("wacc_solar_real_aftertax_pct")}/100)*(1+p_infl)-1,"")', PCT2)
add("w_stef", "Steffen WACC, nominal", "f",
    f'=IF(ISNUMBER({X("wacc_steffen_nominal_aftertax_pct")}),{X("wacc_steffen_nominal_aftertax_pct")}/100,"")', PCT2)
add("w_obs", "Observed WACC (avg of available), nominal", "f",
    f'=IF(COUNT({X("w_irena")},{X("w_stef")})=0,"",AVERAGE({X("w_irena")},{X("w_stef")}))', PCT2)
add("w_dam", "Damodaran WACC (eq. 5), nominal", "f",
    f'=IF(OR({X("dam_ds")}="",{X("dam_crp")}=""),"",p_dv*(p_rf+{X("dam_ds")})*(1-IF({X("dam_tax")}="",p_tax_def,{X("dam_tax")}))'
    f'+(1-p_dv)*(p_rf+p_beta*p_erp+p_lambda*{X("dam_crp")}))', PCT2)
add("ovl", "Overlap obs & Damodaran (1/0)", "f",
    f'=IF(AND(ISNUMBER({X("w_obs")}),ISNUMBER({X("w_dam")})),1,0)', "0")
add("ovl_obs", "helper: obs if overlap", "f", f'=IF({X("ovl")}=1,{X("w_obs")},0)', PCT2)
add("ovl_dam", "helper: Damodaran if overlap", "f", f'=IF({X("ovl")}=1,{X("w_dam")},0)', PCT2)
add("w_damc", "Damodaran WACC, calibrated", "f",
    f'=IF({X("w_dam")}="","",{X("w_dam")}+p_shift)', PCT2)
add("w_grp", "Group default WACC", "f",
    f'=IF({X("group")}="AE",p_r_ae,IF({X("group")}="China",p_r_chn,p_r_emdc))', PCT2)
add("w_grph", "Group high-case WACC", "f",
    f'=IF({X("group")}="AE",p_rh_ae,IF({X("group")}="China",p_rh_chn,p_rh_emdc))', PCT2)
add("w_sel_n", "Market WACC selected, nominal", "f",
    f'=IF(AND(p_use_obs=1,ISNUMBER({X("w_obs")})),{X("w_obs")},IF(AND(p_use_dam=1,ISNUMBER({X("w_damc")})),{X("w_damc")},{X("w_grp")}))', PCT2)
add("w_src", "WACC source", "f",
    f'=IF(AND(p_use_obs=1,ISNUMBER({X("w_obs")})),"Observed (IRENA/Steffen)",IF(AND(p_use_dam=1,ISNUMBER({X("w_damc")})),"Damodaran (eq. 5)","Group default (IHLEG)"))', "General", 20)
add("r_a", "Lens (a) rate, central", "f",
    f'=IF(p_basis=1,(1+{X("w_sel_n")})/(1+p_infl)-1,{X("w_sel_n")})', PCT2)
add("r_ah", "Lens (a) rate, high", "f",
    f'={X("r_a")}+({X("w_grph")}-{X("w_grp")})', PCT2)
# PIMA
add("pim_gap", "PIM efficiency gap", "f",
    f'=IF(AND(p_use_pefa=1,ISNUMBER({X("pefa_pi11_pim_score")})),p_gap_lic-({X("pefa_pi11_pim_score")}-1)/3*(p_gap_lic-p_gap_ae),'
    f'IF({X("group")}="AE",p_gap_ae,IF({X("income")}="Low income",p_gap_lic,p_gap_em)))*(1-p_reform*p_reform_close)', PCT)
add("pim_src", "PIM gap source", "f",
    f'=IF(AND(p_use_pefa=1,ISNUMBER({X("pefa_pi11_pim_score")})),"PEFA PI-11","Income-group default (IMF 2015)")', "General", 18)
add("pub_sh", "Public share of investment", "f",
    f'=IF({X("group")}="AE",p_ps_ae,IF({X("group")}="China",p_ps_chn,p_ps_emdc))', PCT)
add("pim_mult", "PIMA multiplier on investment", "f",
    f'=(1-{X("pub_sh")})+{X("pub_sh")}/(1-{X("pim_gap")})', DEC3)

# investment weights for the world benchmarks (Country_Data and Results_Country share row order)
add("inv_tot", "Total investment I (from Results_Country)", "f", "=Results_Country!F{r}", BN1)
add("fc_wI", "helper: fuel bill per t x I", "f", f'=IF({X("fc_t")}="",0,{X("fc_t")}*{X("inv_tot")})', BN1)
add("cb_wI", "helper: co-benefit per t x I", "f", f'=IF({X("cb_t")}="",0,{X("cb_t")}*{X("inv_tot")})', BN1)
add("w_I", "helper: I where CPAT data", "f", f'=IF({X("fc_t")}="",0,{X("inv_tot")})', BN1)

KEYS = [c[0] for c in cols]
HDR_ROW, R0 = 4, 5
for j, (key, h, kind, src, fmt, width) in enumerate(cols):
    c = wsC.cell(row=HDR_ROW, column=j + 1, value=h)
    c.font, c.fill, c.alignment, c.border = f_h, fill_h, center, box
    wsC.column_dimensions[CL(j + 1)].width = width
    k = wsC.cell(row=HDR_ROW - 1, column=j + 1, value=key); k.font = f_note
wsC.row_dimensions[HDR_ROW].height = 60
for n_, row in C.iterrows():
    rr = R0 + n_
    for j, (key, h, kind, src, fmt, width) in enumerate(cols):
        c = wsC.cell(row=rr, column=j + 1)
        if kind == "v":
            v = src(row)
            if isinstance(v, (float, np.floating)) and np.isnan(v): v = None
            if isinstance(v, np.generic): v = v.item()
            c.value = v; c.font = f_in
        else:
            c.value = src.format(r=rr); c.font = f_base
        c.number_format = fmt
RL = R0 + N - 1
def CR(key): return f"Country_Data!${col_of(key)}${R0}:${col_of(key)}${RL}"
wsC.freeze_panes = "C5"

# world averages and calibration shift (named, on Inputs below the table)
rI = wsI.max_row + 3
put(wsI, f"A{rI}", "Computed world benchmarks", f_st, fill=fill_sub)
for col in "BCDEF": wsI[f"{col}{rI}"].fill = fill_sub
calc_rows = [
 ("p_fc_world", "World fuel bill per tCO2 (investment-weighted)", f'=SUM({CR("fc_wI")})/SUM({CR("w_I")})', "USD/t"),
 ("p_cb_world", "World local co-benefit per tCO2 (investment-weighted)", f'=SUM({CR("cb_wI")})/SUM({CR("w_I")})', "USD/t"),
 ("p_n_overlap", "Countries with both observed and Damodaran WACC", f'=SUM({CR("ovl")})', "count"),
]
for key, lab, fml, unit in calc_rows:
    rI += 1
    put(wsI, f"A{rI}", key, f_note); put(wsI, f"B{rI}", lab)
    put(wsI, f"C{rI}", fml, fmt=BN1 if unit != "count" else "0"); put(wsI, f"D{rI}", unit)
    name(key, wsI, f"C{rI}")
# p_shift cell defined above as None: fill formula
for row in wsI.iter_rows(min_row=5, max_row=wsI.max_row):
    if row[0].value == "p_shift":
        row[2].value = (f'=IF(AND(p_calib=1,p_n_overlap>0),SUM({CR("ovl_obs")})/p_n_overlap-SUM({CR("ovl_dam")})/p_n_overlap,0)')
        row[2].font = f_base; row[2].number_format = PCT2

# -----------------------------------------------------------------------------
# 7. CALC (country x sector)
# -----------------------------------------------------------------------------
CDK = dict(bc="b_c", cc="c_c", bg="b_gen", cg="c_gen", ag="a_gen", bt="b_trn", ct="c_trn", at="a_trn")
_eng = lambda e, cal: f'IF(AND(p_eng=1,ISNUMBER(Country_Data!${e}${{cd}})),Country_Data!${e}${{cd}},{cal})'
BCA = {  # sector -> (b, c, a) formula bodies; {bg} etc. are Country_Data columns
 "GEN": (_eng("{bg}", "Country_Data!${bc}${cd}"), _eng("{cg}", "Country_Data!${cc}${cd}"), _eng("{ag}", "p_a")),
 "GRID": ("p_grid_credit*" + _eng("{bg}", "Country_Data!${bc}${cd}"), "p_grid_credit*" + _eng("{cg}", "Country_Data!${cc}${cd}"),
          "p_grid_credit*" + _eng("{ag}", "p_a")),
 "TRN": (_eng("{bt}", "Country_Data!${bc}${cd}"), _eng("{ct}", "Country_Data!${cc}${cd}"), _eng("{at}", "p_a")),
 "BLD": ("Country_Data!${bc}${cd}", "Country_Data!${cc}${cd}", "p_a"),
 "IND": ("Country_Data!${bc}${cd}", "Country_Data!${cc}${cd}", "p_a"),
 "FUEL": ("0", "0", "1000*p_r_wb/(1-(1+p_r_wb)^-Sector_Params!$D${sr})/p_cpt_fuel"),
 "AFOLU": ("0", "p_c", "1000*p_r_wb/(1-(1+p_r_wb)^-Sector_Params!$D${sr})/p_cpt_afolu"),
}
DRV = {"GEN": "elec", "GRID": "elec", "TRN": "d_trn", "BLD": "d_bld", "IND": "d_ind",
       "FUEL": "d_fuel", "AFOLU": "d_afolu"}
wsK = wb.create_sheet("Calc")
put(wsK, "A1", "Calculation: country x sector (USD bn/yr; flows are annual over the asset life of one year's 2030 investment)", f_t)
put(wsK, "A2", "Equations follow the method note: (1)/(3) allocation, (2) incremental, (4) capital recovery factor, "
    "(6) financial and (7) economic net cost. Incremental cost > 0 = net cost; < 0 = net saving.", f_note)
KC = [  # key, header, formula template ({r} row, {cd} country row), fmt
 ("iso", "ISO3", "=Country_Data!$A${cd}", "General"),
 ("grp", "Group", "=Country_Data!${g}${cd}", "General"),
 ("sec", "Sector", None, "General"),
 ("drv", "Physical driver", "=IF(Country_Data!${d}${cd}=\"\",\"\",Country_Data!${d}${cd})", DEC3),
 ("ppp", "GDP PPP (int$ bn)", "=Country_Data!${ppp}${cd}", BN1),
 ("pl", "Price level", "=Country_Data!${pl}${cd}", DEC2),
 ("dsh", "Driver share in group", "=IF(D{r}=\"\",\"\",D{r}/SUMIFS($D$5:$D${L},$B$5:$B${L},B{r},$C$5:$C${L},C{r}))", '0.0000%'),
 ("gsh", "GDP share in group", "=E{r}/SUMIFS($E$5:$E${L},$B$5:$B${L},B{r},$C$5:$C${L},C{r})", '0.0000%'),
 ("w", "Allocation weight", "=IF(G{r}=\"\",H{r},p_beta_alloc*G{r}+(1-p_beta_alloc)*H{r})*F{r}^p_phi", '0.000000'),
 ("I", "Total investment I (USD bn)", "=INDEX(Sector_Params!$C${pr}:$E${pr},MATCH(B{r},{ph},0))*I{r}/SUMIFS($I$5:$I${L},$B$5:$B${L},B{r},$C$5:$C${L},C{r})", BN1),
 ("ish", "Incremental share", "=INDEX(Sector_Params!$F${sr}:$H${sr},MATCH(B{r},{ph},0))", PCT),
 ("dI", "Incremental investment dI (USD bn)", "=J{r}*K{r}", BN1),
 ("ra", "Rate lens (a)", "=Country_Data!${ra}${cd}+Sector_Params!$E${sr}", PCT2),
 ("rah", "Rate lens (a) high", "=Country_Data!${rah}${cd}+Sector_Params!$E${sr}", PCT2),
 ("rb", "Rate lens (b)/(c)", "=IF(p_use_ibrd=1,p_r_ibrd,p_r_wb)", PCT2),
 ("n", "Asset life", "=Sector_Params!$D${sr}", "0"),
 ("crfa", "CRF (a)", "=M{r}/(1-(1+M{r})^-P{r})", '0.0000'),
 ("crfah", "CRF (a) high", "=N{r}/(1-(1+N{r})^-P{r})", '0.0000'),
 ("crfb", "CRF (b)", "=O{r}/(1-(1+O{r})^-P{r})", '0.0000'),
 ("b", "Savings ratio b", "={bf}*Sector_Params!$I${sr}", DEC3),
 ("c", "Co-benefit ratio c", "={cf}*Sector_Params!$J${sr}", DEC3),
 ("a", "Abatement a (t/yr per $1,000)", "={af}", DEC3),
 ("sav", "Savings on dI", "=T{r}*L{r}", BN1),
 ("cb", "Local co-benefits on dI", "=U{r}*L{r}", BN1),
 ("abt", "Abatement from dI (MtCO2e/yr)", "=V{r}*L{r}", BN1),
 ("kca", "Capital charge on dI (a)", "=Q{r}*L{r}", BN1),
 ("kcah", "Capital charge on dI (a high)", "=R{r}*L{r}", BN1),
 ("kcb", "Capital charge on dI (b)", "=S{r}*L{r}", BN1),
 ("ica", "Incremental cost (a) market", "=Z{r}-W{r}", BN1),
 ("icah", "Incremental cost (a) market high", "=AA{r}-W{r}", BN1),
 ("icb", "Incremental cost (b) WB rate", "=AB{r}-W{r}", BN1),
 ("icc", "Incremental cost (c) WB + local co-benefits", "=AE{r}-X{r}", BN1),
 ("icclo", "memo: (c) + carbon at SPC low", "=AF{r}-p_spc_lo*Y{r}/1000", BN1),
 ("icchi", "memo: (c) + carbon at SPC high", "=AF{r}-p_spc_hi*Y{r}/1000", BN1),
 ("tka", "Total-lens capital charge (a)", "=Q{r}*J{r}", BN1),
 ("tkb", "Total-lens capital charge (b)", "=S{r}*J{r}", BN1),
 ("tsav", "Total-lens savings", "=T{r}*J{r}", BN1),
 ("tcb", "Total-lens co-benefits", "=U{r}*J{r}", BN1),
 ("pen", "Cost-of-capital penalty on I (a - b)", "=AI{r}-AJ{r}", BN1),
 ("pimI", "PIMA-adjusted total investment", "=J{r}*Country_Data!${pm}${cd}", BN1),
 ("pimdI", "PIMA-adjusted incremental investment", "=L{r}*Country_Data!${pm}${cd}", BN1),
]
KK = [k[0] for k in KC]
assert CL(KK.index("kca") + 1) == "Z" and CL(KK.index("icb") + 1) == "AE" and CL(KK.index("tka") + 1) == "AI"
header_row(wsK, 4, 1, [k[1] for k in KC], [7, 7, 7] + [11] * (len(KC) - 3), 48)
L = 4 + N * len(SECT)
sub = dict(g=col_of("group"), ppp=col_of("ppp_use"), pl=col_of("plevel"), ra=col_of("r_a"),
           rah=col_of("r_ah"), bc=col_of("b_c"), cc=col_of("c_c"), pm=col_of("pim_mult"), ph=POOL_HDR, L=L)
rr = 5
for n_ in range(N):
    cd = R0 + n_
    for code, *_ in SECT:
        for j, (key, h, tpl, fmt) in enumerate(KC):
            cell = wsK.cell(row=rr, column=j + 1)
            if key == "sec":
                cell.value = code; cell.font = f_in
            else:
                bf, cf, af = BCA[code]
                bf, cf, af = (x.format(cd=cd, sr=SROW[code], **{k_: col_of(v_) for k_, v_ in CDK.items()}) for x in (bf, cf, af))
                cell.value = tpl.format(r=rr, cd=cd, d=col_of(DRV[code]), pr=PROW[code], sr=SROW[code], bf=bf, cf=cf, af=af, **sub)
                cell.font = f_link if tpl.startswith("=Country_Data") or tpl.startswith("=Sector_Params") else f_base
            cell.number_format = fmt
        rr += 1
wsK.freeze_panes = "D5"
def KR(key): return f"Calc!${CL(KK.index(key)+1)}$5:${CL(KK.index(key)+1)}${L}"

# -----------------------------------------------------------------------------
# 8. RESULTS_COUNTRY
# -----------------------------------------------------------------------------
wsO = wb.create_sheet("Results_Country", 1)
put(wsO, "A1", "Results by country: mitigation investment needs and incremental cost, 2030 (USD bn/yr, 2023 prices)", f_t)
put(wsO, "A2", "Incremental cost = annual capital charge on incremental investment less fuel and operating savings "
    "(less local co-benefits in lens c), over the asset life of one year's investment. Positive = net cost. "
    "Sorted by 2024 GDP.", f_note)
OC = [  # header, formula ({i}=iso cell, {cd}=country data row), fmt, width
 ("ISO3", "=Country_Data!A{cd}", "General", 7),
 ("Country", "=Country_Data!B{cd}", "General", 22),
 ("Group", f"=Country_Data!{col_of('group')}{{cd}}", "General", 7),
 ("WB income", f"=Country_Data!{col_of('income')}{{cd}}", "General", 13),
 ("GDP 2030 (USD bn)", f"=Country_Data!{col_of('gdp30')}{{cd}}", BN, 10),
 ("Total investment I", "=SUMIFS(" + KR("I") + "," + KR("iso") + ",A{r})", BN1, 10),
 ("I, % GDP", "=IFERROR(F{r}/E{r},\"\")", PCT2, 8),
 ("Incremental investment dI", "=SUMIFS(" + KR("dI") + "," + KR("iso") + ",A{r})", BN1, 10),
 ("dI, % GDP", "=IFERROR(H{r}/E{r},\"\")", PCT2, 8),
 ("Market WACC (a)", f"=Country_Data!{col_of('r_a')}{{cd}}", PCT, 8),
 ("WACC source", f"=Country_Data!{col_of('w_src')}{{cd}}", "General", 17),
 ("Savings on dI", "=SUMIFS(" + KR("sav") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("Local co-benefits on dI", "=SUMIFS(" + KR("cb") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("Incr. cost (a) market", "=SUMIFS(" + KR("ica") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("Incr. cost (a) market high", "=SUMIFS(" + KR("icah") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("Incr. cost (b) WB rate", "=SUMIFS(" + KR("icb") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("Incr. cost (c) WB + local co-ben.", "=SUMIFS(" + KR("icc") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("(a) % GDP", "=IFERROR(N{r}/E{r},\"\")", PCT2, 8),
 ("(b) % GDP", "=IFERROR(P{r}/E{r},\"\")", PCT2, 8),
 ("(c) % GDP", "=IFERROR(Q{r}/E{r},\"\")", PCT2, 8),
 ("memo: (c) + carbon, SPC low", "=SUMIFS(" + KR("icclo") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("memo: (c) + carbon, SPC high", "=SUMIFS(" + KR("icchi") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("Abatement (MtCO2e/yr)", "=SUMIFS(" + KR("abt") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("Incr. cost per t (a) USD/t", "=IFERROR(N{r}/W{r}*1000,\"\")", BN, 8),
 ("Incr. cost per t (b) USD/t", "=IFERROR(P{r}/W{r}*1000,\"\")", BN, 8),
 ("Incr. cost per t (c) USD/t", "=IFERROR(Q{r}/W{r}*1000,\"\")", BN, 8),
 ("CoC penalty on I (a-b)", "=SUMIFS(" + KR("pen") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("PIMA-adjusted I", "=SUMIFS(" + KR("pimI") + "," + KR("iso") + ",A{r})", BN1, 9),
 ("PIM efficiency gap", f"=Country_Data!{col_of('pim_gap')}{{cd}}", PCT, 8),
 ("PEFA PI-11 (1-4)", f"=IF(Country_Data!{col_of('pefa_pi11_pim_score')}{{cd}}=\"\",\"\",Country_Data!{col_of('pefa_pi11_pim_score')}{{cd}})", DEC2, 7),
 ("Capex execution (%)", f"=IF(Country_Data!{col_of('capex_execution_rate_pct')}{{cd}}=\"\",\"\",Country_Data!{col_of('capex_execution_rate_pct')}{{cd}})", BN, 8),
 ("Official-private spread (pp)", f"=IF(Country_Data!{col_of('ids_official_private_rate_spread_pp')}{{cd}}=\"\",\"\",Country_Data!{col_of('ids_official_private_rate_spread_pp')}{{cd}})", DEC2, 8),
 ("Gross debt (% GDP)", f"=IF(Country_Data!{col_of('gross_debt_pct_gdp')}{{cd}}=\"\",\"\",Country_Data!{col_of('gross_debt_pct_gdp')}{{cd}})", BN, 8),
]
header_row(wsO, 4, 1, [o[0] for o in OC], [o[3] for o in OC], 48)
for n_ in range(N):
    rr, cd = 5 + n_, R0 + n_
    for j, (h, tpl, fmt, w) in enumerate(OC):
        c = wsO.cell(row=rr, column=j + 1, value=tpl.format(r=rr, cd=cd))
        c.font = f_link if tpl.startswith("=Country_Data") else f_base
        c.number_format = fmt
        if 13 <= j <= 16: c.fill = fill_out
OL = 4 + N
wsO.freeze_panes = "C5"
wsO.auto_filter.ref = f"A4:{CL(len(OC))}{OL}"

# -----------------------------------------------------------------------------
# 9. RESULTS_SUMMARY
# -----------------------------------------------------------------------------
wsY = wb.create_sheet("Results_Summary", 1)
for col, w in zip("ABCDEFGHIJ", (40, 12, 12, 12, 12, 12, 12, 12, 12, 12)):
    wsY.column_dimensions[col].width = w
put(wsY, "A1", "Mitigation (net zero) investment needs and incremental cost, 2030", f_t)
put(wsY, "A2", "USD bn a year, 2023 prices. Flows are annual over the asset life of one year's 2030 investment "
    "(as in the method note, Tables 6-7). Incremental cost > 0 is a net cost; < 0 a net saving.", f_note)
G3 = ["AE", "China", "EMDC"]
header_row(wsY, 4, 1, ["Item"] + G3 + ["World"])
def sgrp(key, g): return f'SUMIFS({KR(key)},{KR("grp")},"{g}")'
SUMROWS = [
 ("Total investment need, I", "I", BN),
 ("Incremental investment need, dI", "dI", BN),
 ("Savings on dI (fuel and operating)", "sav", BN),
 ("Local co-benefits on dI", "cb", BN),
 ("Capital charge on dI, lens (a) market", "kca", BN),
 ("Capital charge on dI, lens (b) WB rate", "kcb", BN),
 ("INCREMENTAL COST (a) market cost of capital", "ica", BN),
 ("INCREMENTAL COST (a) market, high case", "icah", BN),
 ("INCREMENTAL COST (b) World Bank rate", "icb", BN),
 ("INCREMENTAL COST (c) WB rate + local co-benefits", "icc", BN),
 ("memo: (c) + avoided carbon at SPC low", "icclo", BN),
 ("memo: (c) + avoided carbon at SPC high", "icchi", BN),
 ("Abatement from dI (MtCO2e/yr)", "abt", BN),
 ("Cost-of-capital penalty on I, (a) less (b)", "pen", BN),
 ("PIMA-adjusted total investment", "pimI", BN),
 ("PIMA-adjusted incremental investment", "pimdI", BN),
]
YR = {}
r = 5
for lab, key, fmt in SUMROWS:
    YR[key] = r
    bold = lab.startswith("INCREMENTAL")
    put(wsY, f"A{r}", lab, f_b if bold else f_base)
    for j, g in enumerate(G3):
        put(wsY, f"{CL(2+j)}{r}", "=" + sgrp(key, g), f_b if bold else f_base, fmt, fill_out if bold else None)
    put(wsY, f"E{r}", f"=SUM(B{r}:D{r})", f_b, fmt, fill_out if bold else None)
    r += 1
put(wsY, f"A{r}", "GDP 2030 (USD bn, 2024 prices)")
gcol = col_of("gdp30")
for j, g in enumerate(G3):
    put(wsY, f"{CL(2+j)}{r}", f'=SUMIFS(Country_Data!${gcol}${R0}:${gcol}${RL},Country_Data!${col_of("group")}${R0}:${col_of("group")}${RL},"{g}")', fmt=BN)
put(wsY, f"E{r}", f"=SUM(B{r}:D{r})", f_b, BN); YR["gdp"] = r; r += 1
for lab, key in (("I, % of GDP", "I"), ("dI, % of GDP", "dI"), ("Incremental cost (a), % of GDP", "ica"),
                 ("Incremental cost (b), % of GDP", "icb"), ("Incremental cost (c), % of GDP", "icc")):
    put(wsY, f"A{r}", lab)
    for col in "BCDE":
        put(wsY, f"{col}{r}", f"={col}{YR[key]}/{col}{YR['gdp']}", fmt=PCT2)
    r += 1
put(wsY, f"A{r}", "Financial benefit-cost ratio on dI, lens (a)")
for col in "BCDE": put(wsY, f"{col}{r}", f"={col}{YR['sav']}/{col}{YR['kca']}", fmt=DEC2)
r += 1
put(wsY, f"A{r}", "Financial benefit-cost ratio on dI, lens (b)")
for col in "BCDE": put(wsY, f"{col}{r}", f"={col}{YR['sav']}/{col}{YR['kcb']}", fmt=DEC2)
r += 1
put(wsY, f"A{r}", "Economic benefit-cost ratio on dI, lens (c)")
for col in "BCDE": put(wsY, f"{col}{r}", f"=({col}{YR['sav']}+{col}{YR['cb']})/{col}{YR['kcb']}", fmt=DEC2)
r += 1
put(wsY, f"A{r}", "Countries in group")
for j, g in enumerate(G3):
    put(wsY, f"{CL(2+j)}{r}", f'=COUNTIF(Country_Data!${col_of("group")}${R0}:${col_of("group")}${RL},"{g}")', fmt="0")
put(wsY, f"E{r}", f"=SUM(B{r}:D{r})", fmt="0")
r += 2

# Total-investment lens (replicates the method note's Tables 6-7 framing)
put(wsY, f"A{r}", "Total-investment lens (method note Tables 6-7 framing: b and c applied to total I)", f_st); r += 1
header_row(wsY, r, 1, ["Item"] + G3 + ["World"]); r += 1
TL = [("Capital charge on I, (a)", "=SUMIFS({tka},{g},\"{G}\")"),
      ("Capital charge on I, (b)", "=SUMIFS({tkb},{g},\"{G}\")"),
      ("Financial savings on I", "=SUMIFS({tsav},{g},\"{G}\")"),
      ("Co-benefits on I", "=SUMIFS({tcb},{g},\"{G}\")")]
t0 = r
for lab, tpl in TL:
    put(wsY, f"A{r}", lab)
    for j, g in enumerate(G3):
        put(wsY, f"{CL(2+j)}{r}", tpl.format(tka=KR("tka"), tkb=KR("tkb"), tsav=KR("tsav"), tcb=KR("tcb"), g=KR("grp"), G=g), fmt=BN)
    put(wsY, f"E{r}", f"=SUM(B{r}:D{r})", f_b, BN); r += 1
for lab, f in (("Net financial benefit, (a)", "={c}%d-{c}%d" % (t0 + 2, t0)),
               ("Net financial benefit, (b)", "={c}%d-{c}%d" % (t0 + 2, t0 + 1)),
               ("Net benefit with co-benefits, (c)", "={c}%d-{c}%d+{c}%d" % (t0 + 2, t0 + 1, t0 + 3))):
    put(wsY, f"A{r}", lab, f_b)
    for col in "BCDE": put(wsY, f"{col}{r}", f.format(c=col), f_b, BN)
    r += 1
r += 1

# Sector x group
put(wsY, f"A{r}", "Investment by sector and group (USD bn/yr)", f_st); r += 1
header_row(wsY, r, 1, ["Sector", "I: AE", "I: China", "I: EMDC", "I: World", "dI: World",
                       "Incr. cost (a)", "Incr. cost (b)", "Incr. cost (c)"]); r += 1
s0 = r
for code, nm, *_ in SECT:
    put(wsY, f"A{r}", nm)
    for j, g in enumerate(G3):
        put(wsY, f"{CL(2+j)}{r}", f'=SUMIFS({KR("I")},{KR("grp")},"{g}",{KR("sec")},"{code}")', fmt=BN)
    put(wsY, f"E{r}", f"=SUM(B{r}:D{r})", f_b, BN)
    for col, key in (("F", "dI"), ("G", "ica"), ("H", "icb"), ("I", "icc")):
        put(wsY, f"{col}{r}", f'=SUMIFS({KR(key)},{KR("sec")},"{code}")', fmt=BN)
    r += 1
put(wsY, f"A{r}", "Total", f_b)
for col in "BCDEFGHI": put(wsY, f"{col}{r}", f"=SUM({col}{s0}:{col}{r-1})", f_b, BN)
sec_end = r
r += 2

# Charts
ch = BarChart(); ch.type = "bar"; ch.grouping = "clustered"
ch.title = "Incremental cost by lens and group (USD bn/yr)"
for j in range(3):
    ref = Reference(wsY, min_col=2 + j, min_row=YR["ica"], max_row=YR["icc"])
    ch.add_data(ref, titles_from_data=False)
    ch.series[j].tx = None
from openpyxl.chart.series import SeriesLabel
for j, g in enumerate(G3): ch.series[j].tx = SeriesLabel(v=g)
ch.set_categories(Reference(wsY, min_col=1, min_row=YR["ica"], max_row=YR["icc"]))
ch.height, ch.width = 8, 18
wsY.add_chart(ch, "G4")
ch2 = BarChart(); ch2.type = "col"; ch2.grouping = "stacked"; ch2.overlap = 100
ch2.title = "Total investment by sector and group (USD bn/yr)"
for j in range(3):
    ch2.add_data(Reference(wsY, min_col=2 + j, min_row=s0, max_row=sec_end - 1), titles_from_data=False)
for j, g in enumerate(G3): ch2.series[j].tx = SeriesLabel(v=g)
ch2.set_categories(Reference(wsY, min_col=1, min_row=s0, max_row=sec_end - 1))
ch2.height, ch2.width = 8, 18
wsY.add_chart(ch2, f"A{r}")

# -----------------------------------------------------------------------------
# 10. CHECKS: replicate the method note, and benchmarks
# -----------------------------------------------------------------------------
wsQ = wb.create_sheet("Checks")
for col, w in zip("ABCDEFGHIJ", (34, 12, 12, 12, 12, 12, 12, 12, 12, 30)):
    wsQ.column_dimensions[col].width = w
put(wsQ, "A1", "Checks: replication of the method note (Table 7) and benchmarks", f_t)
put(wsQ, "A2", "Part 1 recomputes Table 7 of the method note from its own inputs with this workbook's equations. "
    "Differences should be rounding only.", f_note)
header_row(wsQ, 4, 1, ["Group", "Investment 2030", "WACC (a)", "(a) charge", "(a) net financial",
                       "(b) charge", "(b) net financial", "(c) net with co-benefits", "Note: (a) charge / (c) net", "Status"])
T7 = [("Advanced economies", 2750, 0.055, 205, 506), ("China", 1350, 0.06, 106, 248),
      ("EMDCs other than China", 2400, 0.105, 275, 441)]
for k, (g, inv, w, a_note, c_note) in enumerate(T7):
    rr = 5 + k
    put(wsQ, f"A{rr}", g); put(wsQ, f"B{rr}", inv, f_in, BN); put(wsQ, f"C{rr}", w, f_in, PCT)
    put(wsQ, f"D{rr}", f"=B{rr}*C{rr}/(1-(1+C{rr})^-25)", fmt=BN)
    put(wsQ, f"E{rr}", f"=p_b*B{rr}-D{rr}", fmt=BN)
    put(wsQ, f"F{rr}", f"=B{rr}*p_r_wb/(1-(1+p_r_wb)^-25)", fmt=BN)
    put(wsQ, f"G{rr}", f"=p_b*B{rr}-F{rr}", fmt=BN)
    put(wsQ, f"H{rr}", f"=G{rr}+p_c*B{rr}", fmt=BN)
    put(wsQ, f"I{rr}", f"{a_note} / {c_note}", f_in)
    put(wsQ, f"J{rr}", f'=IF(AND(ABS(D{rr}-{a_note})<=2,ABS(H{rr}-{c_note})<=3),"OK","CHECK")', f_b)
put(wsQ, "A9", "Part 2: benchmarks for incremental investment (share of GDP a year)", f_st)
header_row(wsQ, 10, 1, ["Country / group", "Model dI % GDP", "Benchmark", "Source", "", "", "", "", "", "Note"])
BM = [("BRA", "Brazil", 0.0192, "SISEPUEDE LAC (Kalra et al. 2023), Table A3"),
      ("MEX", "Mexico", 0.01, "SISEPUEDE LAC (Kalra et al. 2023): 'about 1%'"),
      ("DOM", "Dominican Republic", 0.016, "SISEPUEDE LAC (Kalra et al. 2023)")]
for k, (iso, nm, v, src) in enumerate(BM):
    rr = 11 + k
    put(wsQ, f"A{rr}", nm)
    put(wsQ, f"B{rr}", f'=INDEX(Results_Country!$I$5:$I${OL},MATCH("{iso}",Results_Country!$A$5:$A${OL},0))', fmt=PCT2)
    put(wsQ, f"C{rr}", v, f_in, PCT2); put(wsQ, f"D{rr}", src)
    put(wsQ, f"J{rr}", "SISEPUEDE covers all sectors incl. some non-capital costs; scope differs", f_note)
rr = 14
put(wsQ, f"A{rr}", "EMDCs (model, mitigation only)")
put(wsQ, f"B{rr}", f"=Results_Summary!D{YR['dI']}/Results_Summary!D{YR['gdp']}", fmt=PCT2)
put(wsQ, f"C{rr}", 0.014, f_in, PCT2); put(wsQ, f"D{rr}", "CCDR synthesis (World Bank 2023): +1.4% of GDP a year (incl. adaptation)")
rr = 15
put(wsQ, f"A{rr}", "World total investment I (model)")
put(wsQ, f"B{rr}", f"=Results_Summary!E{YR['I']}", fmt=BN)
put(wsQ, f"C{rr}", "=p_ce_world+p_nat_emdc+p_nat_ae+p_nat_chn", fmt=BN)
put(wsQ, f"D{rr}", "Pools on Inputs (must match exactly)")
put(wsQ, f"J{rr}", f'=IF(ABS(B{rr}-C{rr})<0.5,"OK","CHECK")', f_b)
rr = 16
put(wsQ, f"A{rr}", "Countries with observed WACC / Damodaran / default")
put(wsQ, f"B{rr}", f'=COUNTIF({CR("w_src")},"Observed*")', fmt="0")
put(wsQ, f"C{rr}", f'=COUNTIF({CR("w_src")},"Damodaran*")', fmt="0")
put(wsQ, f"D{rr}", f'=COUNTIF({CR("w_src")},"Group*")', fmt="0")
rr = 17
put(wsQ, f"A{rr}", "Countries with CPAT data / PEFA PI-11 score")
put(wsQ, f"B{rr}", f'=SUM({CR("has_cpat")})', fmt="0")
put(wsQ, f"C{rr}", f'=COUNTIF({CR("pim_src")},"PEFA*")', fmt="0")

# -----------------------------------------------------------------------------
# 11. SOURCES
# -----------------------------------------------------------------------------
wsZ = wb.create_sheet("Sources")
wsZ.column_dimensions["A"].width = 44; wsZ.column_dimensions["B"].width = 90
put(wsZ, "A1", "Sources", f_t)
SRC = [
 ("Method", "Global Investment Needs for Climate Action: A Method Based on SISEPUEDE (Oct 2026), equations 1-7, Tables 3-7 and Appendix C"),
 ("IHLEG 3 (2024)", "Bhattacharya, Songwe, Soubeyran & Stern, Raising Ambition and Accelerating Delivery of Climate Finance. LSE Grantham. https://www.lse.ac.uk/granthaminstitute/wp-content/uploads/2024/11/Raising-ambition-and-accelerating-delivery-of-climate-finance_Third-IHLEG-report.pdf"),
 ("IEA-IFC (2023)", "Scaling Up Private Finance for Clean Energy in Emerging and Developing Economies. https://www.iea.org/reports/scaling-up-private-finance-for-clean-energy-in-emerging-and-developing-economies (sector split of EMDE clean energy; 60% private; EMDE totals)"),
 ("McKinsey (2022)", "The Net-Zero Transition: What It Would Cost, What It Could Bring (USD 9.2 tn gross, +3.5 tn)"),
 ("CCDR synthesis (2023)", "World Bank, The Development, Climate, and Nature Crisis. https://openknowledge.worldbank.org/handle/10986/40652"),
 ("UNEP SFN (2021)", "State of Finance for Nature: Tripling Investments in Nature-based Solutions by 2030 (USD 133 bn/yr)"),
 ("Fay et al. (2016)", "Discounting Costs and Benefits in Economic Analysis of World Bank Projects (6% social rate)"),
 ("World Bank (2017, 2024)", "Shadow Price of Carbon in Economic Analysis (USD 50-100/tCO2e in 2030)"),
 ("Kharas & Rivard (2022)", "Debt, creditworthiness, and climate. Brookings GWP 180 (IBRD ~5.2%)"),
 ("IMF (2015)", "Making Public Investment More Efficient (efficiency gaps 13% AE, 27% EM, 40% LIDC; PIM can close up to two-thirds). https://infrastructuregovern.imf.org/content/dam/PIMA/Forum/IGEUR/Resources/Making%20Public%20Investment%20More%20Efficient%202015.pdf"),
 ("Damodaran", "Country Default Spreads and Risk Premiums, ctryprem.xlsx, https://pages.stern.nyu.edu/~adamodar/ (mature ERP 4.20% Jul 2026 per BVWire, https://www.bvresources.com/articles/bvwire/global-damodaran-revises-country-risk-framework-after-us-rating-downgrade)"),
 ("IMF WEO", "World Economic Outlook database (GDP USD and PPP, real growth 2025-26). File in TaxMacroFiscalDashboard/build/raw_data"),
 ("IMF CPAT / FFS", "IMF Fossil Fuel Subsidies data (CPAT engine), 2024: emissions, energy use, supply costs, externalities. https://www.imf.org/en/topics/climate-change/energy-subsidies"),
 ("Dashboard: cost of capital", "IRENA (2022) WACC appendix (real after-tax, ref. year 2021); Steffen et al. (2025) Sci. Data WACC database (nominal after-tax); IEA Cost of Capital Observatory; World Bank IDS"),
 ("Dashboard: PIMA", "World Bank BOOST (capex execution); PEFA 2016 framework (PI-11, PI-12.2); UNDP Global Climate Public Finance Review 2022 (climate-PFM diagnostic flag)"),
 ("IMF advanced economies", "WEO country classification (41 economies) used for the IHLEG AE group"),
]
for k, (a, b) in enumerate(SRC):
    put(wsZ, f"A{k+3}", a, f_b, al=wrap); put(wsZ, f"B{k+3}", b, al=wrap)

# -----------------------------------------------------------------------------
# 12. README
# -----------------------------------------------------------------------------
wsR.column_dimensions["A"].width = 3; wsR.column_dimensions["B"].width = 120
put(wsR, "B1", f"Mitigation (Net Zero) Investment Needs Model {VERSION}", f_t)
TXT = [
 ("h", "What it does"),
 ("p", "Estimates, for every economy with IMF WEO GDP (%d economies) and seven mitigation sectors, the 2030 total "
       "and incremental investment need and its incremental cost under three lenses: (a) market cost of capital, "
       "(b) the World Bank rate (6%%), and (c) the World Bank rate with local co-benefits. All cells recalculate "
       "from the Inputs sheet." % N),
 ("h", "Method (equation numbers follow the SISEPUEDE method note)"),
 ("p", "1. Totals. Group pools for 2030 (advanced economies, China, other EMDCs) come from IHLEG (2024) and IEA-IFC (2023); "
       "clean energy is split into sectors with the IEA-IFC shares; AFOLU uses IHLEG and UNEP (Inputs, Sector_Params)."),
 ("p", "2. Allocation (eq. 1/3, Tier 3 of the note). Each group-sector pool is shared across its countries in proportion to "
       "w = [beta x physical-driver share + (1 - beta) x PPP GDP share] x price level^phi. Drivers come from IMF CPAT 2024: "
       "electricity demand (power, grids), road fuel (transport), building and industrial energy use, fossil supply (fuels), "
       "non-CO2 GHG plus net LULUCF source (AFOLU). Countries without CPAT data use GDP only."),
 ("p", "3. Incremental investment (eq. 2): dI = I x incremental share (AE/China 0.38 from McKinsey 2022; EMDC 0.24 from CCDR vs IHLEG)."),
 ("p", "4. Capital charge (eq. 4): A = dI x CRF(r, n), CRF = r / (1 - (1 + r)^-n), asset life n by sector."),
 ("p", "5. Cost of capital (eq. 5). Lens (a) uses a country WACC: the average of observed RE WACC (IRENA 2021, converted to "
       "nominal; Steffen et al. 2025) where available; else a Damodaran-based WACC, D/V(rf + default spread)(1 - tax) + "
       "E/V(rf + beta x ERP + lambda x CRP), shifted to the observed level; else the IHLEG group rate (AE 5.5%, China 6%, "
       "EMDC 10.5%). The high case adds the group gap to 7% / 8% / 15%. Lens (b)/(c) use 6% (or IBRD ~5.2% by switch)."),
 ("p", "6. Savings, co-benefits and abatement (eq. 6-7), v1.1. Power generation and transport use engineering "
       "parameters: renewable output per dollar displaces the country's fossil generation mix (Ember) at CPAT 2024 fuel supply "
       "costs and air-pollution costs, with heat rates and IPCC 2006 CO2 factors; EVs replace gasoline cars (fuel, power and "
       "maintenance per vehicle over the purchase premium). Grids carry no benefits of their own (switch p_grid_credit). "
       "Low-emission fuels and AFOLU are costed per tonne (USD 100/t and 50/t). Buildings and industry keep the calibrated "
       "b = 0.133 and c = 0.129 (SISEPUEDE LAC) scaled by CPAT country indices relative to the investment-weighted world average."),
 ("p", "7. Incremental cost: (a) = CRF(WACC) x dI - savings; (b) = CRF(6%) x dI - savings; (c) = (b) - local co-benefits. "
       "Memo lines add avoided emissions (a x dI) at the World Bank shadow price of carbon. Positive = net cost."),
 ("p", "8. PIMA. PIMA-adjusted investment = I x [(1 - public share) + public share / (1 - efficiency gap)]. The gap is "
       "mapped from the PEFA PI-11 score where available, otherwise the IMF (2015) income-group gap; the reform switch "
       "closes two-thirds of it."),
 ("p", "9. Total-investment lens. Results_Summary also applies b and c to total I, which is the framing of the method "
       "note's Tables 6-7; Checks reproduces Table 7."),
 ("h", "Sheets"),
 ("p", "Results_Summary - group and world totals, lens comparison, sector split, charts.  Results_Country - one row per "
       "economy, filterable.  Inputs - every scenario lever.  Sector_Params - sector shares, asset lives, technology premia, "
       "incremental shares, pools.  Country_Data - data (blue) and country parameters (formulas): WACC, indices, PIMA.  "
       "Damodaran - country risk inputs (paste or auto-load).  Calc - country x sector engine.  Checks - replication and "
       "benchmarks.  Lists - name aliases.  Sources."),
 ("h", "Colour code"),
 ("p", "Blue text = hard-coded input or loaded data. Black = formula. Green = link to another sheet. Yellow fill = key "
       "assumption to review. † in a source note = not verified against the primary source, or an author's assumption."),
 ("h", "Data notes and caveats"),
 ("p", "- Damodaran data: " + dam_note + " Until it is supplied, countries without an observed WACC use the IHLEG group rate."),
 ("p", "- Rates are compared as published (nominal market WACC vs the 6% real World Bank rate), as in the method note. "
       "Set p_basis = 1 to deflate market rates to real terms."),
 ("p", "- The allocation is a Tier 3 estimate. It reproduces the group totals exactly, but country values depend on the "
       "drivers and on beta and phi. Replace with SISEPUEDE or CCDR country runs where they exist."),
 ("p", "- In buildings and industry, b, c and a are still calibrated on one region (SISEPUEDE LAC) and IEA NZE. "
       "Power is conservative: fossil plants avoided in growing systems are not credited, and 2024 fuel prices are used."),
 ("p", "- The Sensitivity sheet holds static results from build/run_sensitivity.py; re-run it after changing inputs."),
 ("p", "- The AE/China AFOLU pools, the end-use split, the China WACC, the public shares and the PEFA mapping are "
       "author's assumptions (yellow)."),
 ("p", "- CPAT local air pollution cost is the uncharged (implicit-subsidy) air-pollution cost of fossil fuel use, 2024."),
 ("p", "- Build: python InvestmentNeeds/build/build_investment_model.py (reads only repo files)."),
]
r = 3
for kind, t in TXT:
    if kind == "h":
        r += 1; put(wsR, f"B{r}", t, f_st)
    else:
        put(wsR, f"B{r}", t, al=wrap)
        wsR.row_dimensions[r].height = max(14, 13 * (len(t) // 140 + 1))
    r += 1

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = ws.title in ("Calc", "Country_Data", "Lists", "Damodaran")
order = ["README", "Results_Summary", "Results_Country", "Inputs", "Sector_Params", "Country_Data",
         "Damodaran", "Calc", "Checks", "Sources", "Lists"]
wb._sheets = [wb[s] for s in order]
wb.active = 0
os.makedirs(os.path.dirname(OUT_XLSX), exist_ok=True)
wb.save(OUT_XLSX)
print(f"wrote {OUT_XLSX}: {N} economies x {len(SECT)} sectors; Damodaran rows: {len(dam_rows)}; D/V median {DV_MED:.2f}")
