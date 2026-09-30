#!/usr/bin/env python3
# =============================================================================
# GLOBAL FISCAL INDICATORS — build pipeline (all countries)
# =============================================================================
# Reads the raw downloaded source files, reshapes everything into ONE long
# dataset (all years, all series), then filters down to the indicator set used
# in the dashboard and writes Data_long and Data_wide.
#
#   STAGE 1  load_all()      -> long_all : iso3, indicator, year, value, source
#   STAGE 2  ONE filter line -> long     : latest observation per indicator
#   STAGE 3  to_wide()       -> wide     : one row per economy, + __year cols
#
# Country universe: ALL countries (217, from the WDI country metadata: any entity
# with a non-empty Region, which excludes aggregates like "Euro area"). A boolean
# is_ccdr column flags the 98 published-CCDR economies so you can still subset.
#
# Run:  python build_global_indicators.py
#       python build_global_indicators.py --ccdr-only        # restrict to the 98
#       python build_global_indicators.py --emit-cpat-xlsx   # also write the R helper file
# Deps: pandas, numpy, openpyxl, pyxlsb, xlsxwriter
#
# NOTE ON THE SECOND SCRIPT: convert_cpat_xlsb.py used to be separate. It only ever
# existed because R cannot read .xlsb; Python reads it directly via pyxlsb. It is now
# the --emit-cpat-xlsx flag below, so there is one script rather than two.
#
# -----------------------------------------------------------------------------
# WHERE EVERY SOURCE FILE COMES FROM
# -----------------------------------------------------------------------------
# All World Bank WDI files below follow the same procedure:
#   go to https://data.worldbank.org/indicator/<CODE>  -> "Download CSV"
#   or直 https://api.worldbank.org/v2/country/all/indicator/<CODE>?downloadformat=csv
#   The zip contains three CSVs; the data file has 4 header rows to skip and a
#   "Last Updated Date" stamp on line 3. Keep the zip unopened - this script
#   reads straight from it.
#
#   API_NY_GDP_PETR_RT_ZS_*.zip   NY.GDP.PETR.RT.ZS  oil rents % GDP
#   API_NY_GDP_NGAS_RT_ZS_*.zip   NY.GDP.NGAS.RT.ZS  natural gas rents % GDP
#   API_NY_GDP_COAL_RT_ZS_*.zip   NY.GDP.COAL.RT.ZS  coal rents % GDP
#   API_TX_VAL_FUEL_ZS_UN_*.zip   TX.VAL.FUEL.ZS.UN  fuel exports % merch exports
#   API_TM_VAL_FUEL_ZS_UN_*.zip   TM.VAL.FUEL.ZS.UN  fuel imports % merch imports
#   API_EG_IMP_CONS_ZS_*.zip      EG.IMP.CONS.ZS     net energy imports % energy use
#
# dataset_..._IMF_RES_WEO_9_0_0.csv
#   IMF World Economic Outlook. https://www.imf.org/external/datamapper/datasets/WEO
#   Export from the IMF data portal. Wide: one row per COUNTRY x INDICATOR,
#   year columns 2017-2026 (2025-26 are projections - excluded below).
#
# world-imf2026.xlsx
#   IMF WoRLD 2026. https://www.imf.org/en/topics/fiscal-policies/world-revenue-longitudinal-database
#   Sheet "Data". NOTE: every revenue variable is ALREADY percent of GDP.
#
# imffossilfuelsubsidiesdata.xlsb
#   IMF Fossil Fuel Subsidies (CPAT engine).
#   https://www.imf.org/en/topics/climate-change/energy-subsidies -> country spreadsheet.
#   Binary .xlsb, needs pyxlsb. Sheets All_Explicit / All_Implicit hold 2024 totals
#   in US$ bn; sheet "data" holds the raw CPAT panel (year cols 9=2021 .. 12=2024).
#
# P_Data_Extract_From_International_Debt_Statistics__2_.xlsx
#   World Bank IDS via DataBank. https://databank.worldbank.org/source/international-debt-statistics
#   Select: Country = all, Counterpart-Area = World (WLD), Series = the 19 codes,
#   Time = 2015-2024. ".." means missing. NOTE: on the DT.INR.* rate series a
#   value of 0 means NO NEW BORROWING that year, not a 0% rate - hence nonzero=True.
#
# WB_BOOST_WIDEF.csv
#   World Bank BOOST harmonised. https://data360.worldbank.org/en/dataset/WB_BOOST
#   COMP_BREAKDOWN_1: ..._PARAM_1 = approved budget, ..._PARAM_2 = executed.
#
# assessments_1787749571.csv
#   PEFA 2016 framework. https://www.pefa.org/assessments/batch-downloads
#   Filter: Framework = 2016, Type = National, Status = Final, all indicators.
#
# 2024-PPI-Full-DTA.dta
#   World Bank PPI. https://ppi.worldbank.org/en/ppidata -> Stata download.
#   Project level; investment in US$ millions; FCY = financial closure year.
#
# CRDF-RP_2024.xlsx
#   OECD Climate-Related Development Finance, RECIPIENT perspective (not provider).
#   https://www.oecd.org/en/topics/sub-issues/development-finance-for-climate-and-the-environment.html
#
# release_generation_yearly_global.csv
#   Ember Yearly Electricity Data. https://ember-energy.org/data/yearly-electricity-data/
#   NOTE: there is no "Oil" category - "Other fossil" is oil plus manufactured gases.
#
# SteffenEtAl2025_WACC_database.csv
#   Steffen et al. (2025), Scientific Data. https://www.nature.com/articles/s41597-025-05912-x
#   Values are strings with a % sign. Basis is NOMINAL after-tax - do not splice with IRENA.
#
# IRENA_Cost_of_financing_renewable_power_Appendix_2023.pdf  -> transcribed inline below.
#   IRENA (2022) Renewable power generation costs in 2021, reproduced as the appendix to
#   IRENA (2023) The Cost of Financing for Renewable Power. REFERENCE YEAR IS 2021.
#   Real after-tax WACC. Table is a PDF; values transcribed into IRENA_SOLAR.
#
# WB_RISE_WIDEF.csv
#   World Bank RISE (Regulatory Indicators for Sustainable Energy), 2024 edition.
#   https://data360.worldbank.org/en/dataset/WB_RISE  -> download WIDEF csv.
#   140 countries, 2010-2023, scores 0-100. NOTE: RISE 2024 rebuilt the renewable pillar -
#   the pre-2022 sub-indicators (counterparty risk, attributes of financial incentives)
#   no longer exist. Closest current analogue is WB_RISE_RE_LVL_PLYNG_FLD.
#
# dataverse_files.zip  (contains psrt/PSRT_V1_Database.dta)
#   Power Sector Reform Tracker (PSRT) v1, ISEP / Urpelainen & Yang.
#   https://dataverse.harvard.edu/dataset.xhtml?persistentId=doi:10.7910/DVN/M7SY6X
#   142 non-OECD countries, 1982-2013, eight 0/1 reform flags. ENDS 2013: structural
#   context, not current state. High-income countries absent by design.
#   Cite: Urpelainen & Yang (2019), Energy Strategy Reviews 23: 152-162.
#
# rethinking_*.xlsx  (six files)  -- HELD, NOT USED.
#   World Bank Rethinking Power Sector Reform deep dives.
#   https://datacatalog.worldbank.org/search/dataset/0038380
#   15 countries plus 3 Indian states only. The cost-recovery workbook has a genuine
#   quasi-fiscal deficit (% GDP) series decomposed into collection losses, T&D losses
#   and underpricing - but 2010-2012 and 15 countries, so not a usable column.
#
# CofCObservatoryData.xlsx / UNDP-Global-Climate-Public-Finance-Review-2022.pdf
#   IEA https://www.iea.org/reports/cost-of-capital-observatory/dashboard  (11 countries)
#   UNDP https://www.undp.org/publications/undp-global-climate-public-finance-review
#   Both used only as country lists, transcribed inline below.
# =============================================================================

import re, sys, glob, zipfile, unicodedata, warnings
import numpy as np, pandas as pd
warnings.filterwarnings("ignore")

# Paths resolve relative to this script unless overridden by CLI flag or env var:
#   python build_global_indicators.py --raw ../raw_data --out ../output
#   RAW_DIR=... OUT_DIR=... python build_global_indicators.py
import os, argparse
_here = os.path.dirname(os.path.abspath(__file__))
_ap = argparse.ArgumentParser(add_help=True)
_ap.add_argument("--raw", default=os.environ.get("RAW_DIR", os.path.join(_here, "..", "raw_data")))
_ap.add_argument("--out", default=os.environ.get("OUT_DIR", os.path.join(_here, "..", "output")))
_ap.add_argument("--config", default=os.environ.get("CONFIG_DIR", os.path.join(_here, "..", "config")))
_ap.add_argument("--ccdr-only", action="store_true")
_ap.add_argument("--emit-cpat-xlsx", action="store_true")
_args, _ = _ap.parse_known_args()
RAW = os.path.join(os.path.abspath(_args.raw), "")
OUT = os.path.join(os.path.abspath(_args.out), "")
CONFIG = os.path.join(os.path.abspath(_args.config), "")
os.makedirs(OUT, exist_ok=True)
CCDR_ONLY      = _args.ccdr_only
EMIT_CPAT_XLSX = _args.emit_cpat_xlsx

# -----------------------------------------------------------------------------
# COUNTRY UNIVERSE — every real country, taken from the WDI country metadata.
# Any entity with a non-empty Region is a country; blank Region = aggregate.
# -----------------------------------------------------------------------------
def _wdi_meta():
    path = glob.glob(RAW + "API_*_DS2_*.zip")[0]
    with zipfile.ZipFile(path) as z:
        nm = [n for n in z.namelist() if n.startswith("Metadata_Country")][0]
        with z.open(nm) as fh: return pd.read_csv(fh)
_meta = _wdi_meta()
_real = _meta[_meta.Region.notna() & (_meta.Region.astype(str).str.strip() != "")]
NAME   = dict(zip(_real["Country Code"], _real["TableName"]))
REGION = dict(zip(_real["Country Code"], _real["Region"]))
INCOME = dict(zip(_real["Country Code"], _real["IncomeGroup"]))
ISO    = set(NAME)

# The 98 published CCDR economies, kept as a FLAG rather than a filter.
CCDR_ISO = set("""AGO BWA COM COD ETH KEN MDG MWI MUS MOZ NAM RWA SYC SOM ZAF SSD TZA UGA ZMB ZWE
BEN BFA CPV CMR CAF TCD COG CIV GAB GMB GHA GIN GNB LBR MLI MRT NER SEN SLE TGO KHM CHN FJI IDN
KIR LAO MYS MHL MNG PHL THA TUV VNM ALB ARM AZE BIH HRV GEO KAZ XKX KGZ MDA MNE MKD POL ROU SRB
TJK TUR UZB ARG BRA COL DMA DOM ECU GRD HND PRY PER LCA VCT DJI EGY IRQ JOR LBN MAR PAK TUN PSE
YEM BGD BTN MDV NPL LKA""".split())
if CCDR_ONLY: ISO = ISO & CCDR_ISO

# -----------------------------------------------------------------------------
# NAME -> ISO3 matching. Sources spell countries differently ("Türkiye" /
# "Turkiye", "Lao People's Democratic Republic" / "Lao PDR"). Normalise by
# stripping accents, parentheticals and administrative suffixes, then fall back
# to the text before the first comma. Aggregates deliberately fail to match and
# are dropped; the run prints how many labels each source lost.
# -----------------------------------------------------------------------------
SUFF=[r",?\s*(the\s+)?(islamic|arab|bolivarian|plurinational|federal democratic|federated|people s|democratic|socialist|principality|grand duchy)?\s*(republic|kingdom|state|states|union|commonwealth|federation|principality)s?\s+of(\s+the)?$",
      r",\s*(rep\.|dem\. rep\.|the|fed\. sts\.|islamic rep\.|arab rep\.|rb|prc|cr)$", r"^the\s+"]
def _n(s):
    s=unicodedata.normalize("NFKD",str(s)).encode("ascii","ignore").decode().lower().strip()
    s=re.sub(r"\([^)]*\)","",s)                      # drop parentheticals
    for _ in range(3):
        for p in SUFF: s=re.sub(p,"",s).strip()
    return re.sub(r"[^a-z0-9]+"," ",s).strip()
ALIAS={"vietnam":"VNM","viet nam":"VNM","turkiye":"TUR","turkey":"TUR","egypt":"EGY","gambia":"GMB",
"yemen":"YEM","kyrgyzstan":"KGZ","lao pdr":"LAO","laos":"LAO","lao people s democratic":"LAO",
"cape verde":"CPV","cabo verde":"CPV","west bank and gaza strip":"PSE","cote d ivoire":"CIV",
"ivory coast":"CIV","congo":"COG","dr congo":"COD","democratic republic of congo":"COD",
"democratic republic of the congo":"COD","congo dem":"COD","st lucia":"LCA","saint lucia":"LCA",
"st vincent and grenadines":"VCT","saint vincent and grenadines":"VCT","brunei":"BRN",
"russia":"RUS","syria":"SYR","iran":"IRN","venezuela":"VEN","bolivia":"BOL","tanzania":"TZA",
"moldova":"MDA","north macedonia":"MKD","macedonia":"MKD","micronesia":"FSM","bahamas":"BHS",
"hong kong":"HKG","hong kong special administrative region":"HKG","macao":"MAC","korea":"KOR",
"south korea":"KOR","republic of korea":"KOR","north korea":"PRK",
"democratic people s republic of korea":"PRK","slovak":"SVK","czechia":"CZE","czech":"CZE",
"eswatini":"SWZ","swaziland":"SWZ","timor leste":"TLS","east timor":"TLS","kosovo":"XKX",
"curacao":"CUW","chinese taipei":"TWN","taiwan":"TWN","china":"CHN","somalia":"SOM",
"guyana":"GUY","st kitts and nevis":"KNA","saint kitts and nevis":"KNA","cook islands":"COK",
"nauru":"NRU","niue":"NIU","bosnia and herzegovina":"BIH","czech republic":"CZE","lao people s democratic republic":"LAO",
"macao special administrative region":"MAC","taiwan province of china":"TWN",
"saint vincent and the grenadines":"VCT","st vincent and the grenadines":"VCT",
"montserrat":"MSR","saint helena":"SHN","tokelau":"TKL","wallis and futuna":"WLF",
"st pierre and miquelon":"SPM","saint pierre and miquelon":"SPM"}
def make_matcher(lut_names):
    """lut_names: dict {official name -> iso3}. Returns iso_of(name) -> iso3 or None."""
    lut={_n(k):v for k,v in lut_names.items()}; lut.update(ALIAS)
    def iso_of(x):
        if x is None: return None
        raw=str(x).strip()
        # Guard: stripping "Democratic Republic of the" reduces DR Congo to "congo",
        # which collides with Republic of Congo. Decide these two before normalising.
        rl=raw.lower()
        if "congo" in rl:
            return "COD" if ("dem" in rl or "drc" in rl or "kinshasa" in rl) else "COG"
        cands=[_n(raw), _n(raw.split(",")[0]), _n(raw.split(" - ")[0])]
        cands.append(re.sub(r'[^a-z0-9]+',' ',unicodedata.normalize('NFKD',str(raw)).encode('ascii','ignore').decode().lower()).strip())
        for cand in cands:
            if cand in lut: return lut[cand]
        return None
    return iso_of

iso_of_name = make_matcher({name: iso for iso, name in NAME.items()})  # name -> ISO3
UNMATCHED = {}
def to_iso(series, source):
    out = pd.Series([iso_of_name(v) for v in series], index=getattr(series, "index", None))
    miss = sorted({str(v) for v, o in zip(series, out) if o is None and pd.notna(v)})
    if miss: UNMATCHED[source] = miss
    return out

ROWS = []
def emit(iso, indicator, year, value, source):
    if iso in ISO and value is not None and not (isinstance(value, float) and np.isnan(value)):
        ROWS.append({"iso3": iso, "indicator": indicator, "year": str(year),
                     "value": round(float(value), 3), "source": source})

def wdi_zip(pattern):
    """Read a WDI download zip -> long frame (iso3, year, value). 4 header rows."""
    path = glob.glob(RAW + pattern)[0]
    with zipfile.ZipFile(path) as z:
        name = [n for n in z.namelist() if n.startswith("API_") and "Metadata" not in n][0]
        with z.open(name) as fh:
            df = pd.read_csv(fh, skiprows=4)
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    yrs = [c for c in df.columns if str(c).isdigit()]
    out = df.melt(id_vars=["Country Code"], value_vars=yrs,
                  var_name="year", value_name="value").rename(columns={"Country Code": "iso3"})
    return out.dropna(subset=["value"]), path.split("/")[-1]

# =============================================================================
# STAGE 1 — load every source into one long frame (ALL years retained)
# =============================================================================

# --- IMF WEO ----------------------------------------------------------------
WEO_F = glob.glob(RAW + "dataset_*IMF_RES_WEO*.csv")[0]
weo = pd.read_csv(WEO_F, low_memory=False)
weo["iso"] = to_iso(weo.COUNTRY, "IMF WEO")
WY = [c for c in weo.columns if str(c).strip().isdigit() and int(c) <= 2024]  # drop projections
def weo_ser(ind):
    # Two WEO labels can resolve to one ISO (e.g. a country and its SAR); keep the
    # first non-null per ISO so downstream .loc lookups return scalars, not Series.
    s = weo[(weo.INDICATOR == ind) & weo.iso.notna()]
    return s.groupby("iso")[WY].first()
G  = "Gross debt, General government, Percent of GDP"
R  = "Revenue, General government, Percent of GDP"
NB = "Net lending (+) / net borrowing (-), General government, Percent of GDP"
PB = "Primary net lending (+) / net borrowing (-), General government, Percent of GDP"
GDPU = "Gross domestic product (GDP), Current prices, US dollar"
wf = WEO_F.split("/")[-1]
for iso, row in weo_ser(G).iterrows():
    for y in WY: emit(iso, "gross_debt_pct_gdp", y, row[y], wf)
# interest = primary balance less overall balance; then / revenue
pb, nb, rev = weo_ser(PB), weo_ser(NB), weo_ser(R)
for iso in set(pb.index) & set(nb.index) & set(rev.index):
    for y in WY:
        a, b, c = pb.loc[iso, y], nb.loc[iso, y], rev.loc[iso, y]
        if pd.notna(a) and pd.notna(b) and pd.notna(c) and c > 0:
            v = (a - b) / c * 100
            if 0 <= v < 100: emit(iso, "interest_to_revenue_pct", y, v, wf + " (derived)")
GDP = {}                                   # iso -> (usd_bn, year) latest actual
for iso, row in weo_ser(GDPU).iterrows():
    for y in reversed(WY):
        if pd.notna(row[y]): GDP[iso] = (float(row[y]), y); break

# --- IMF WoRLD --------------------------------------------------------------
WORLD_F = "world-imf2026.xlsx"
wo = pd.read_excel(RAW + WORLD_F, sheet_name="Data"); wo.columns = [str(c) for c in wo.columns]
for _, r in wo.iterrows():
    if r.ISO3 in ISO and 2010 <= r.year <= 2024:
        emit(r.ISO3, "tax_revenue_pct_gdp", int(r.year), r.TaxRev, WORLD_F)   # already % of GDP
        if pd.notna(r.TaxSalExc) and pd.notna(r.TotRev) and r.TotRev > 0:
            emit(r.ISO3, "excise_revenue_pct_of_revenue", int(r.year),
                 r.TaxSalExc / r.TotRev * 100, WORLD_F)

# --- IMF Fossil Fuel Subsidies (CPAT) ---------------------------------------
FFS_F = "imffossilfuelsubsidiesdata.xlsb"
from pyxlsb import open_workbook
xb = open_workbook(RAW + FFS_F)
def ffs_total(sheet, total_col):                 # 'Total' block header column index
    out = {}
    with xb.get_sheet(sheet) as sh:
        for i, row in enumerate(sh.rows()):
            c = {x.c: x.v for x in row}
            if i < 4: continue
            iso, v = c.get(1), c.get(total_col)
            if isinstance(iso, str) and len(iso) == 3 and isinstance(v, (int, float)):
                out[iso.upper()] = float(v)
    return out
exp_bn, imp_bn = ffs_total("All_Explicit", 25), ffs_total("All_Implicit", 15)
if EMIT_CPAT_XLSX:
    # R cannot read .xlsb, so optionally dump the three sheets with explicit c0..cN
    # headers (readxl drops leading all-empty columns, which would shift positions).
    def _grid(sheet, ncol):
        rows = []
        with xb.get_sheet(sheet) as sh:
            for r in sh.rows():
                cells = {c.c: c.v for c in r}
                rows.append([cells.get(i) for i in range(ncol)])
        d = pd.DataFrame(rows); d.columns = [f"c{i}" for i in range(ncol)]; return d
    with pd.ExcelWriter(OUT + "imffossilfuelsubsidiesdata_sheets.xlsx", engine="xlsxwriter") as _x:
        _grid("All_Explicit", 26).to_excel(_x, sheet_name="All_Explicit", index=False)
        _grid("All_Implicit", 16).to_excel(_x, sheet_name="All_Implicit", index=False)
        _grid("data", 13).to_excel(_x, sheet_name="data", index=False)
    print("wrote imffossilfuelsubsidiesdata_sheets.xlsx (for the R pipeline)")
for iso in ISO:
    g = GDP.get(iso)
    if not g: continue
    e, m = exp_bn.get(iso), imp_bn.get(iso)
    if e is not None: emit(iso, "ff_subsidy_explicit_pct_gdp", 2024, e / g[0] * 100, FFS_F)
    if m is not None: emit(iso, "ff_subsidy_implicit_pct_gdp", 2024, m / g[0] * 100, FFS_F)
    if e is not None and m is not None:
        emit(iso, "ff_subsidy_total_pct_gdp", 2024, (e + m) / g[0] * 100, f"{FFS_F} + WEO")
CPAT_YR = {9:2021, 10:2022, 11:2023, 12:2024}    # column index -> year on sheet "data"
# Retail (consumer) prices. Diesel and gasoline are $/litre; gas and coal are $/GJ.
# CPAT has no COMMERCIAL sector - it splits residential / industry / power / other.
CPAT_PRICE = {
 "Retail price, die, all": "price_diesel_usd_litre",
 "Retail price, gso, all": "price_gasoline_usd_litre",
 "Retail price, nga, res": "price_gas_residential_usd_gj",
 "Retail price, nga, ind": "price_gas_industry_usd_gj",
 "Retail price, nga, pow": "price_gas_power_usd_gj",
 "Retail price, coa, res": "price_coal_residential_usd_gj",
 "Retail price, coa, ind": "price_coal_industry_usd_gj",
 "Retail price, coa, pow": "price_coal_power_usd_gj",
 "Retail price, ecy, res": "price_electricity_residential_usd_kwh",
 "Retail price, ecy, ind": "price_electricity_industry_usd_kwh",
}
with xb.get_sheet("data") as sh:                 # U1 = baseline scenario
    for row in sh.rows():
        c = {x.c: x.v for x in row}
        nm, scen = c.get(2), c.get(1)
        if scen != "U1": continue
        iso = str(c.get(4) or "").upper()
        if nm == "Share of electricity from renewables - baseline":
            for col, yr in CPAT_YR.items():
                v = c.get(col)
                if isinstance(v, (int, float)):
                    emit(iso, "renewable_share_electricity_pct", yr,
                         v * 100 if v <= 1.0001 else v, FFS_F)
        elif nm in CPAT_PRICE:
            for col, yr in CPAT_YR.items():
                v = c.get(col)
                if isinstance(v, (int, float)) and v > 0:
                    emit(iso, CPAT_PRICE[nm], yr, v, FFS_F)

# --- World Bank IDS ---------------------------------------------------------
IDS_F = "P_Data_Extract_From_International_Debt_Statistics__2_.xlsx"
ids = pd.read_excel(RAW + IDS_F, sheet_name="Data")
ids = ids[ids["Country Code"].notna()].copy()
IY = [c for c in ids.columns if str(c).startswith("20") and int(str(c)[:4]) <= 2024]
for y in IY: ids[y] = pd.to_numeric(ids[y].replace("..", np.nan), errors="coerce")
ids = ids[ids["Country Code"].isin(ISO)]
def ids_panel(code, nonzero=False):
    """iso -> {year: value}. nonzero=True for DT.INR.* where 0 = no new borrowing."""
    out = {}
    for _, r in ids[ids["Series Code"] == code].iterrows():
        d = {int(str(y)[:4]): r[y] for y in IY
             if pd.notna(r[y]) and (r[y] > 0 if nonzero else True)}
        if d: out[r["Country Code"]] = d
    return out
for iso, d in ids_panel("DT.CUR.USDL.ZS").items():
    for y, v in d.items(): emit(iso, "ids_usd_share_ppg_debt_pct", y, v, IDS_F)
for iso, d in ids_panel("DT.MAT.DPPG", nonzero=True).items():
    for y, v in d.items(): emit(iso, "ids_avg_maturity_years", y, v, IDS_F)
prv, off = ids_panel("DT.INR.PRVT", True), ids_panel("DT.INR.OFFT", True)
for iso, d in prv.items():
    for y, v in d.items(): emit(iso, "ids_private_creditor_rate_pct", y, v, IDS_F)
for iso in set(prv) & set(off):                  # spread at the latest year each side has
    yp, yo = max(prv[iso]), max(off[iso])
    emit(iso, "ids_official_private_rate_spread_pp", max(yp, yo),
         prv[iso][yp] - off[iso][yo], IDS_F)
tot = ids_panel("DT.DOD.DPPG.CD")
for code, ind in [("DT.DOD.PRVT.CD", "ids_private_creditor_share_pct"),
                  ("DT.DOD.MLAT.CD", "ids_multilateral_share_pct")]:
    part = ids_panel(code)
    for iso in set(part) & set(tot):
        for y in set(part[iso]) & set(tot[iso]):
            if tot[iso][y] > 0: emit(iso, ind, y, part[iso][y] / tot[iso][y] * 100, IDS_F)

# --- World Bank BOOST -------------------------------------------------------
BOOST_F = "WB_BOOST_WIDEF.csv"
bo = pd.read_csv(RAW + BOOST_F, low_memory=False)
BY = [c for c in bo.columns if str(c).isdigit()]
cap = bo[(bo.INDICATOR == "WB_BOOST_EXP_ECON_CAP_EXP") & (bo.UNIT_MEASURE == "XDC")]
A = cap[cap.COMP_BREAKDOWN_1 == "WB_BOOST_BUDGET_PARAM_1"].set_index("REF_AREA")[BY]
E = cap[cap.COMP_BREAKDOWN_1 == "WB_BOOST_BUDGET_PARAM_2"].set_index("REF_AREA")[BY]
for iso in set(A.index) & set(E.index) & ISO:
    for y in BY:
        a, e = A.loc[iso, y], E.loc[iso, y]
        if pd.notna(a) and pd.notna(e) and a > 0:
            emit(iso, "capex_execution_rate_pct", y, e / a * 100, BOOST_F)
env = bo[(bo.INDICATOR == "WB_BOOST_EXP_FUNC_ENV_PRO") & (bo.UNIT_MEASURE == "PT_EXP")
         & (bo.COMP_BREAKDOWN_1 == "WB_BOOST_BUDGET_PARAM_2")].set_index("REF_AREA")[BY]
for iso, row in env.iterrows():
    if iso not in ISO: continue
    if isinstance(row, pd.DataFrame): row = row.iloc[0]
    for y in BY: emit(iso, "env_protection_pct_of_expenditure", y, row[y], BOOST_F)

# --- PEFA -------------------------------------------------------------------
PEFA_F = "assessments_1787749571.csv"
pe = pd.read_csv(RAW + PEFA_F)
SCORE = {"A":4.0, "B+":3.5, "B":3.0, "C+":2.5, "C":2.0, "D+":1.5, "D":1.0, "D*":1.0}
pe["iso"] = to_iso(pe.Country, "PEFA")
for col, ind in [("PI-11","pefa_pi11_pim_score"), ("PI-12.2","pefa_pi122_asset_monitoring"),
                 ("PI-10.3","pefa_pi103_contingent_liabilities")]:
    for _, r in pe[pe.iso.notna() & pe[col].isin(SCORE)].iterrows():
        emit(r.iso, ind, int(r.Year), SCORE[r[col]], PEFA_F)

# --- World Bank PPI (5-year average, so a single derived observation) --------
PPI_F = "2024-PPI-Full-DTA.dta"
pp = pd.read_stata(RAW + PPI_F)
pp["iso"] = to_iso(pp.country.astype(str), "World Bank PPI")
agg = pp[(pp.FCY.between(2020, 2024)) & pp.iso.notna()].groupby("iso")["investment"].sum() / 5.0 / 1000.0
for iso, v in agg.items():
    if iso in GDP and v > 0:
        emit(iso, "ppi_commitments_pct_gdp", "2020-2024 avg", v / GDP[iso][0] * 100, f"{PPI_F} + WEO")

# --- OECD CRDF, recipient perspective ---------------------------------------
CRDF_F = "CRDF-RP_2024.xlsx"
cr = pd.read_excel(RAW + CRDF_F, sheet_name="2024")
MITCOL = "Mitigation-related development finance (includes overlap) - Commitment - 2024 USD thousand"
cr["iso"] = to_iso(cr["Recipient Name"], "OECD CRDF")
for iso, v in (cr[cr.iso.notna()].groupby("iso")[MITCOL].sum() / 1e6).items():
    if iso in GDP and v > 0:
        emit(iso, "mitigation_finance_received_pct_gdp", 2024, v / GDP[iso][0] * 100,
             f"{CRDF_F} + WEO")

# --- Ember ------------------------------------------------------------------
EMBER_F = "release_generation_yearly_global.csv"
em = pd.read_csv(RAW + EMBER_F, low_memory=False)
em_all = em[em["Area type"] == "Country or economy"].copy()
of = em_all[(em_all["Electricity source"] == "Other fossil") & em_all["Share of generation (%)"].notna()]
for _, r in of.iterrows():
    emit(r["ISO 3 code"], "other_fossil_gen_share_pct", int(r.Year), r["Share of generation (%)"], EMBER_F)

# Renewables: Ember publishes the year-on-year share change directly, and capacity in GW.
ren = em_all[em_all["Electricity source"] == "Renewables"]
# Ember carries a country's record forward unchanged where it has no observations
# (e.g. Kiribati: identical 0.04 TWh, 25% share, every year 2020-24). The YoY change
# then reads a spurious 0.0 pp. Detect it by the WHOLE record being constant - total
# generation flat as well as the share - which distinguishes carry-forward from a
# genuinely stable mix like Iceland or Paraguay.
_recent = em_all[em_all.Year >= em_all.Year.max() - 5]
_tg = _recent[_recent["Electricity source"] == "Total generation"].groupby("ISO 3 code")["Generation (TWh)"].nunique()
_rs = _recent[_recent["Electricity source"] == "Renewables"].groupby("ISO 3 code")["Share of generation (%)"].nunique()
CARRY_FWD = {i for i in _rs.index if _rs.get(i, 0) <= 1 and _tg.get(i, 0) <= 1}
sh_ = ren[ren["Share of generation YoY change (% points)"].notna()
          & ~ren["ISO 3 code"].isin(CARRY_FWD)]
for _, r in sh_.iterrows():
    emit(r["ISO 3 code"], "re_share_change_pp", int(r.Year),
         r["Share of generation YoY change (% points)"], EMBER_F)
cap = ren[ren["Capacity (GW)"].notna()][["ISO 3 code", "Year", "Capacity (GW)"]]
cap = cap.sort_values(["ISO 3 code", "Year"])
cap["prev"] = cap.groupby("ISO 3 code")["Capacity (GW)"].shift(1)
cap = cap[(cap.prev.notna()) & (cap.prev > 0)]
for _, r in cap.iterrows():
    emit(r["ISO 3 code"], "re_capacity_growth_pct", int(r.Year),
         (r["Capacity (GW)"] / r.prev - 1) * 100, EMBER_F)

# --- WDI energy and rents ---------------------------------------------------
for pat, ind in [("API_NY_GDP_PETR_RT_ZS*.zip", "oil_rents_pct_gdp"),
                 ("API_NY_GDP_NGAS_RT_ZS*.zip", "gas_rents_pct_gdp"),
                 ("API_NY_GDP_COAL_RT_ZS*.zip", "coal_rents_pct_gdp"),
                 ("API_TX_VAL_FUEL_ZS_UN*.zip", "fuel_exports_pct_merch"),
                 ("API_TM_VAL_FUEL_ZS_UN*.zip", "fuel_imports_pct_merch"),
                 ("API_EG_IMP_CONS_ZS*.zip",    "energy_import_dep_pct")]:
    df, fn = wdi_zip(pat)
    for _, r in df.iterrows(): emit(r.iso3, ind, r.year, r.value, fn)

# --- World Bank RISE ---------------------------------------------------------
RISE_F = "WB_RISE_WIDEF.csv"
ri = pd.read_csv(RAW + RISE_F, low_memory=False)
RY = [c for c in ri.columns if str(c).isdigit()]
RISE_KEEP = {"WB_RISE_RE_ALL": "rise_renewable_energy_score",
             "WB_RISE_RE_LVL_PLYNG_FLD": "rise_level_playing_field_score",
             "WB_RISE_EE_FMEE": "rise_ee_financing_mechanisms_score"}
for code, ind in RISE_KEEP.items():
    for _, r in ri[ri.INDICATOR == code].iterrows():
        for y in RY:
            if pd.notna(r[y]): emit(r.REF_AREA, ind, y, r[y], RISE_F)

# --- Power Sector Reform Tracker (PSRT) --------------------------------------
# Market structure coded from the reform flags:
#   3 = wholesale market   2 = single buyer (IPPs, no market)   1 = vertically integrated
PSRT_F = "dataverse_files.zip"
with zipfile.ZipFile(RAW + PSRT_F) as z:
    inner = [n for n in z.namelist() if n.endswith("PSRT_V1_Database.dta")][0]
    with z.open(inner) as fh:
        ps = pd.read_stata(fh)
ps["iso"] = to_iso(ps.cntry, "PSRT")
ps = ps[ps.iso.notna()]
for _, r in ps.iterrows():
    y = int(r.year)
    emit(r.iso, "power_market_structure", y,
         3 if r.r_wem == 1 else (2 if r.r_ipp == 1 else 1), PSRT_F)
    emit(r.iso, "independent_regulator", y, int(r.r_reg), PSRT_F)
    emit(r.iso, "generation_unbundled", y, int(r.r_und), PSRT_F)

# --- WACC: IRENA (transcribed from PDF appendix), Steffen, IEA ---------------
# IRENA (2022) real after-tax WACC, solar PV, reference year 2021.
IRENA_SOLAR = {"DZA":11.0,"ARG":13.8,"AZE":7.0,"BGD":6.8,"BIH":10.4,"BRA":6.3,"BFA":5.8,
"CHN":2.5,"COL":5.6,"HRV":5.3,"DOM":5.6,"ECU":12.2,"EGY":8.8,"ETH":8.4,"GHA":9.5,"HND":4.6,
"IDN":6.0,"IRQ":9.6,"JOR":8.2,"KAZ":6.3,"KEN":8.4,"LBN":21.0,"MYS":5.4,"MUS":4.6,"MNG":8.0,
"MNE":8.8,"MAR":6.7,"NAM":4.2,"PAK":9.2,"PER":5.2,"PHL":5.7,"POL":3.9,"ROU":5.1,"RWA":5.6,
"SEN":4.3,"ZAF":5.2,"LKA":10.3,"THA":4.5,"TUN":9.3,"TUR":7.5,"UGA":6.9,"VNM":6.0,"YEM":17.2}
for iso, v in IRENA_SOLAR.items():
    emit(iso, "wacc_solar_real_aftertax_pct", 2021, v,
         "IRENA_Cost_of_financing_renewable_power_Appendix_2023.pdf")
STEF_F = "SteffenEtAl2025_WACC_database.csv"
st = pd.read_csv(RAW + STEF_F)
st = st[st.Variable == "WACC (nominal after-tax)"].copy()
st["val"] = pd.to_numeric(st.Value.astype(str).str.replace("%", "").str.strip(), errors="coerce")
for iso, g in st[st.val.notna() & st.ISO3.isin(ISO)].groupby("ISO3"):
    ly = g["Financing year"].max()
    emit(iso, "wacc_steffen_nominal_aftertax_pct", int(ly),
         g[g["Financing year"] == ly].val.mean(), STEF_F)
IEA_COUNTRIES = ["BRA","IDN","KEN","MYS","PHL","SEN","THA","VNM","ZAF"]   # CofCObservatoryData.xlsx
for iso in IEA_COUNTRIES:
    emit(iso, "wacc_iea_observatory_available", 2024, 1, "CofCObservatoryData.xlsx")

# --- UNDP climate-PFM diagnostic flag (Table 4.2, transcribed) ---------------
UNDP_YES = ["ARM","BGD","BEN","KHM","CHN","COL","ECU","ETH","FJI","GHA","HND","IDN","KEN",
"KIR","MHL","MAR","MOZ","NPL","PAK","PHL","RWA","SYC","TZA","THA","UGA","VNM","AZE","GEO",
"HRV","GMB","MDG","MUS","PER"]   # Table 4.2 entries that fall inside the CCDR set
for iso in ISO:
    emit(iso, "climate_pfm_diagnostic_on_record", 2022, 1 if iso in UNDP_YES else 0,
         "UNDP-Global-Climate-Public-Finance-Review-2022.pdf (Table 4.2)")

long_all = pd.DataFrame(ROWS)
long_all["value"] = long_all.value.round(3)
long_all["country"] = long_all.iso3.map(NAME)
long_all["year"] = long_all.year.astype(str).str.replace(r"\.0$", "", regex=True)
long_all["region"]  = long_all.iso3.map(REGION)
long_all["is_ccdr"] = long_all.iso3.isin(CCDR_ISO)
long_all = long_all[["iso3","country","region","is_ccdr","indicator","year","value","source"]]

# =============================================================================
# STAGE 2 — THE FILTER: latest observation per economy x indicator
# =============================================================================
long_all["year"] = long_all.year.astype(str)
long = (long_all.assign(_ord=lambda d: pd.to_numeric(d.year.str[:4], errors="coerce").fillna(0))
                .sort_values(["iso3","indicator","_ord"])
                .groupby(["iso3","indicator"], as_index=False).last().drop(columns="_ord"))

# --- derived indicators that combine two filtered series ---------------------
piv = long.pivot(index="iso3", columns="indicator", values="value")
yr  = long.pivot(index="iso3", columns="indicator", values="year")
add = []
for iso in piv.index:
    o, g_, c = [piv.get(k, pd.Series(dtype=float)).get(iso) for k in
                ("oil_rents_pct_gdp","gas_rents_pct_gdp","coal_rents_pct_gdp")]
    if all(pd.notna(x) for x in (o, g_, c)):
        add.append((iso,"fossil_rents_total_pct_gdp",yr.oil_rents_pct_gdp.get(iso),o+g_+c,
                    "WDI NY.GDP.PETR+NGAS+COAL.RT.ZS"))
    x, m = piv.get("fuel_exports_pct_merch",pd.Series(dtype=float)).get(iso), \
           piv.get("fuel_imports_pct_merch",pd.Series(dtype=float)).get(iso)
    if pd.notna(x) and pd.notna(m):
        add.append((iso,"net_fuel_trade_pct_merch",yr.fuel_exports_pct_merch.get(iso),x-m,
                    "WDI TX.VAL.FUEL.ZS.UN less TM.VAL.FUEL.ZS.UN"))
    n = sum(pd.notna(piv.get(k,pd.Series(dtype=float)).get(iso)) for k in
            ("wacc_solar_real_aftertax_pct","wacc_steffen_nominal_aftertax_pct",
             "wacc_iea_observatory_available"))
    # 0 sources means no WACC evidence at all - that is a missing row, not a value of 0
    if n > 0:
        add.append((iso,"wacc_n_sources","2021-2024",n,"IRENA / Steffen et al. / IEA"))
long = pd.concat([long, pd.DataFrame(add, columns=["iso3","indicator","year","value","source"])
                  .assign(country=lambda d: d.iso3.map(NAME))], ignore_index=True)

# --- keep only the 35 dashboard indicators, in block order -------------------
# The indicator set is defined ONCE, in config/indicators.csv, and read by both this
# pipeline and build_dashboard.py. Adding an indicator means editing that file only.
_spec = pd.read_csv(CONFIG + "indicators.csv").sort_values("order")
KEEP = _spec.code.tolist()

long = (long[long.indicator.isin(KEEP)]
        .assign(indicator=lambda d: pd.Categorical(d.indicator, KEEP, ordered=True))
        .sort_values(["country","indicator"]))
long["region"]  = long.iso3.map(REGION)
long["income"]  = long.iso3.map(INCOME)
long["is_ccdr"] = long.iso3.isin(CCDR_ISO)
long = long[["iso3","country","region","income","is_ccdr","indicator","value","year","source"]]

# =============================================================================
# STAGE 3 — wide (dashboard layout: value columns then matching __year columns)
# =============================================================================
wide = long.pivot(index="iso3", columns="indicator", values="value")
long["year"] = long.year.astype(str).str.replace(r"\.0$", "", regex=True)
wyr  = long.pivot(index="iso3", columns="indicator", values="year").add_suffix("__year")
wide = pd.concat([wide, wyr], axis=1).reset_index()
wide.insert(1, "country", wide.iso3.map(NAME))
wide.insert(2, "region",  wide.iso3.map(REGION))
wide.insert(3, "income",  wide.iso3.map(INCOME))
wide.insert(4, "is_ccdr", wide.iso3.isin(CCDR_ISO))
wide = (wide.reindex(columns=["iso3","country","region","income","is_ccdr"] + KEEP + [k + "__year" for k in KEEP])
            .sort_values("country").reset_index(drop=True))

# =============================================================================
# OUTPUT
# =============================================================================
PFX = "ccdr_" if CCDR_ONLY else "global_"
with pd.ExcelWriter(OUT + PFX + "indicators.xlsx", engine="xlsxwriter") as xw:
    long_all.to_excel(xw, sheet_name="Long_all_years", index=False)
    long.to_excel(xw,     sheet_name="Data_long",      index=False)
    wide.to_excel(xw,     sheet_name="Data_wide",      index=False)
long_all.to_csv(OUT + PFX + "long_all_years.csv", index=False)
long.to_csv(OUT + PFX + "data_long.csv", index=False)
wide.to_csv(OUT + PFX + "data_wide.csv", index=False)

print(f"Stage 1  long_all : {len(long_all):>6,} rows  {long_all.indicator.nunique()} series  "
      f"{long_all.iso3.nunique()} economies")
print(f"Stage 2  long     : {len(long):>6,} rows  {long.indicator.nunique()} indicators")
print(f"Stage 3  wide     : {wide.shape[0]} economies x {len(KEEP)} indicators")
missing = [k for k in KEEP if k not in set(long.indicator.astype(str))]
if missing:
    raise SystemExit("PIPELINE/CONFIG MISMATCH - config/indicators.csv lists indicators this "
                     "pipeline never emits:\n   " + "\n   ".join(missing) +
                     "\nEither add the extraction above, or remove them from the config.")
extra = sorted(set(long_all.indicator) - set(KEEP))
print("missing indicators: none")
if extra: print(f"note: {len(extra)} series built but not in config (dropped at Stage 2)")
print(f"         of which CCDR-98 : {long[long.is_ccdr].iso3.nunique()} economies")
if UNMATCHED:
    print("\nlabels dropped (aggregates and non-countries), by source:")
    for k, v in UNMATCHED.items(): print(f"   {k:16s} {len(v):3d}  e.g. {', '.join(v[:4])}")
