#!/usr/bin/env python3
# =============================================================================
# NMIM v0.2 - data extraction
# =============================================================================
# Reads repository sources and writes the national model's input tables to
# national_model/data/ in CPAT code conventions (mit.<indicator>.<fuel>.<sector>).
#
#   countries.csv      iso, name, groups, CPAT income group, WEO GDP and growth,
#                      price level, market WACC from the v1.1 workbook
#   energy.csv         end-use energy by CPAT row code and year 2021-2024, PJ
#   prices.csv         supply cost and air-pollution cost by row code, USD/GJ (2024)
#   power.csv          Ember generation and capacity by source (latest year)
#   cbam.csv           CBAM goods output, growth and emission factors (CBAM tool)
#   emissions.csv      CPAT 2024 CO2 and GHG
#   elasticities.csv   CPAT income elasticities and AEEI by CPAT income group
#   power_market.csv   power market structure, private participation, cost recovery,
#                      RISE scores (v0.2: drives the investment typology)
#
# Sources (all in the repository, no network):
#   TaxMacroFiscalDashboard/build/output/imffossilfuelsubsidiesdata_sheets.xlsx (CPAT panel)
#   TaxMacroFiscalDashboard/build/raw_data/ WEO csv, Ember csv
#   CPAT_excel_not_master/mitigation_copypaste/data/ elasticities.csv, countries.csv, fuel_units.csv
#   CBAMPolicyTool/CBAM_Policy_Tool_global_v1.0.xlsx, sheet "Country data"
#   InvestmentNeeds/Mitigation_Investment_Needs_v1.1.xlsx, sheet Country_Data (WACC)
#   TaxMacroFiscalDashboard/build/raw_data/ dataverse_files.zip (Power Sector Reform
#     Tracker), 2024-PPI-Full-DTA.dta (World Bank PPI); output/global_data_wide.csv (RISE,
#     explicit fossil fuel subsidies)
#
# Run: python InvestmentNeeds/national_model/extract_data_v0_2.py
# =============================================================================
import os, re, unicodedata, zipfile
import numpy as np, pandas as pd, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
DASH = os.path.join(ROOT, "TaxMacroFiscalDashboard", "build")
REPO = os.path.abspath(os.path.join(ROOT, ".."))               # repository root: CPAT, CBAM
MVP = os.path.join(REPO, "CPAT_excel_not_master", "mitigation_copypaste", "data")
OUT = os.path.join(HERE, "data")
os.makedirs(OUT, exist_ok=True)
YEARS = [2021, 2022, 2023, 2024]
YC = {2021: "y21", 2022: "y22", 2023: "y23", 2024: "y24"}

# --- CPAT panel ---------------------------------------------------------------
cp = pd.read_excel(os.path.join(DASH, "output", "imffossilfuelsubsidiesdata_sheets.xlsx"), sheet_name="data")
cp.columns = ["country", "u", "label", "code", "iso", "inc", "reg", "mt", "sub", "y21", "y22", "y23", "y24"]
cp = cp.drop_duplicates(["iso", "mt"])
cp = cp[cp.iso.astype(str).str.len() == 3]
def cpv(mt, y=2024):
    s = cp[cp.mt == mt].set_index("iso")[YC[y]]
    return pd.to_numeric(s, errors="coerce")

# GJ per litre (CPAT fuel_units: gj_per_price_unit; oop is per barrel in legacy, here per litre)
fu = pd.read_csv(os.path.join(MVP, "fuel_units.csv")).set_index("fuel").gj_per_price_unit
GJL = {"gso": fu["gso"], "die": fu["die"], "lpg": fu["lpg"], "ker": fu["ker"], "oop": fu["oop"] / 158.987}

# --- end-use energy rows (PJ) -------------------------------------------------
# Oil products are reported in bn litres for all sectors; CPAT's transport and residential
# portions split them. Coal and gas are in EJ by sector; electricity in TWh.
rows = []
isos = sorted(cp.iso.unique())
for y in YEARS:
    L = {f: cpv(f"mit.con.{f}.all.1", y) for f in ["gso", "die", "lpg", "ker", "oop"]}
    trs = {f: cpv(f"mit.trs.{f}.all.a.1", y) for f in ["gso", "die"]}
    resc = {f: cpv(f"mit.rescon.{f}.all.1", 2024) for f in ["gso", "die", "lpg", "ker"]}
    E = lambda m: cpv(m, y)
    for i in isos:
        g = lambda s: float(s.get(i)) if pd.notna(s.get(i)) else 0.0
        pj = {f: g(L[f]) * GJL[f] * 1000 for f in L}             # bn l x GJ/l = EJ; x1000 = PJ
        t_g, t_d = g(trs["gso"]) or 0.0, g(trs["die"]) or 0.0
        v = {
            "mit.con.gso.rod": pj["gso"] * t_g,
            "mit.con.die.rod": pj["die"] * t_d,
            "mit.con.coa.res": g(E("mit.con.coa.res.1")) * 1000,
            "mit.con.nga.res": g(E("mit.con.nga.res.1")) * 1000,
            "mit.con.lpg.res": pj["lpg"] * g(resc["lpg"]),
            "mit.con.ker.res": pj["ker"] * g(resc["ker"]),
            "mit.con.ecy.res": g(E("mit.con.ecy.res.1")) * 3.6,
            "mit.con.coa.ind": (g(E("mit.con.coa.ind.1")) + g(E("mit.con.coa.other.1"))) * 1000,
            "mit.con.nga.ind": (g(E("mit.con.nga.ind.1")) + g(E("mit.con.nga.other.1"))) * 1000,
            # remaining oil: other oil products, non-transport gasoline/diesel not counted
            # as residential, non-residential LPG and kerosene
            "mit.con.oil.ind": pj["oop"] + pj["gso"] * max(0, 1 - t_g) + pj["die"] * max(0, 1 - t_d)
                               + pj["lpg"] * (1 - g(resc["lpg"])) + pj["ker"] * (1 - g(resc["ker"])),
            "mit.con.ecy.ind": (g(E("mit.con.ecy.ind.1")) + g(E("mit.con.ecy.other.1"))) * 3.6,
            "mit.con.coa.pow": g(E("mit.con.coa.pow.1")) * 1000,
            "mit.con.nga.pow": g(E("mit.con.nga.pow.1")) * 1000,
        }
        has = pd.notna(L["gso"].get(i)) or pd.notna(E("mit.con.ecy.res.1").get(i))
        if not has: continue
        for k, val in v.items():
            rows.append((i, k, y, val))
en = pd.DataFrame(rows, columns=["iso", "code", "year", "value_pj"])
en.to_csv(os.path.join(OUT, "energy.csv"), index=False)

# --- prices and air pollution, USD/GJ, 2024 ------------------------------------
pr = []
SC = {  # row code -> (supply cost code, air pollution code, unit factor to $/GJ)
 "gso.rod": ("mit.sup.cost.gso.all.1", "mit.airpol.cost.gso.all.1", 1 / GJL["gso"]),
 "die.rod": ("mit.sup.cost.die.all.1", "mit.airpol.cost.die.all.1", 1 / GJL["die"]),
 "coa.res": ("mit.sup.cost.coa.res.1", "mit.airpol.cost.coa.res.1", 1.0),
 "nga.res": ("mit.sup.cost.nga.res.1", "mit.airpol.cost.nga.res.1", 1.0),
 "lpg.res": ("mit.sup.cost.lpg.all.1", "mit.airpol.cost.lpg.all.1", 1 / GJL["lpg"]),
 "ker.res": ("mit.sup.cost.ker.all.1", "mit.airpol.cost.ker.all.1", 1 / GJL["ker"]),
 "coa.ind": ("mit.sup.cost.coa.ind.1", "mit.airpol.cost.coa.ind.1", 1.0),
 "nga.ind": ("mit.sup.cost.nga.ind.1", "mit.airpol.cost.nga.ind.1", 1.0),
 "oil.ind": ("mit.sup.cost.oop.all.1", "mit.airpol.cost.oop.1", 1 / GJL["oop"]),
 "coa.pow": ("mit.sup.cost.coa.pow.1", "mit.airpol.cost.coa.pow.1", 1.0),
 "nga.pow": ("mit.sup.cost.nga.pow.1", "mit.airpol.cost.nga.pow.1", 1.0),
 "oil.pow": ("mit.sup.cost.oop.all.1", "mit.airpol.cost.oop.1", 1 / GJL["oop"]),
}
for rc, (s, a, f) in SC.items():
    sv, av = cpv(s), cpv(a)
    for i in isos:
        pr.append((i, f"mit.sp.{rc}", float(sv.get(i)) * f if pd.notna(sv.get(i)) else np.nan))
        pr.append((i, f"mit.ap.{rc}", float(av.get(i)) * f if pd.notna(av.get(i)) else np.nan))
pd.DataFrame(pr, columns=["iso", "code", "usd_per_gj"]).to_csv(os.path.join(OUT, "prices.csv"), index=False)

# --- emissions -------------------------------------------------------------------
em = pd.DataFrame({"co2": cpv("mit.co2.tot.1"), "ghg_exc": cpv("mit.ghg.tot.exc.1"),
                   "ghg_inc": cpv("mit.ghg.tot.inc.1")})
em.index.name = "iso"; em.to_csv(os.path.join(OUT, "emissions.csv"))

# --- Ember power -------------------------------------------------------------
eb = pd.read_csv(os.path.join(DASH, "raw_data", "release_generation_yearly_global.csv"), low_memory=False)
eb = eb[(eb["Area type"] == "Country or economy") & eb["ISO 3 code"].notna()]
last = eb.groupby("ISO 3 code").Year.max()
cur = eb[eb.Year == eb["ISO 3 code"].map(last)]
G = cur.pivot_table(index="ISO 3 code", columns="Electricity source", values="Generation (TWh)", aggfunc="sum")
K = cur.pivot_table(index="ISO 3 code", columns="Electricity source", values="Capacity (GW)", aggfunc="sum")
prev = eb[eb.Year == eb["ISO 3 code"].map(last) - 3]
Gp = prev.pivot_table(index="ISO 3 code", columns="Electricity source", values="Generation (TWh)", aggfunc="sum")
pw = pd.DataFrame(index=G.index)
pw["year"] = last
for src, col in [("Total generation", "gen_tot"), ("Coal", "gen_coa"), ("Gas", "gen_nga"), ("Other fossil", "gen_oil"),
                 ("Solar", "gen_sol"), ("Wind", "gen_wnd"), ("Hydro", "gen_hyd"), ("Nuclear", "gen_nuc"),
                 ("Bioenergy", "gen_bio"), ("Other renewables", "gen_ore")]:
    pw[col] = G.get(src)
for src, col in [("Coal", "cap_coa"), ("Gas", "cap_nga"), ("Other fossil", "cap_oil"), ("Solar", "cap_sol"), ("Wind", "cap_wnd")]:
    pw[col] = K.get(src)
fos_now = G[["Coal", "Gas", "Other fossil"]].sum(axis=1) / G["Total generation"]
fos_3y = Gp[["Coal", "Gas", "Other fossil"]].sum(axis=1) / Gp["Total generation"]
pw["fos_trend_pp_yr"] = ((fos_3y - fos_now) / 3 * 100).reindex(pw.index)   # + = fossil share falling
pw.index.name = "iso"
pw.to_csv(os.path.join(OUT, "power.csv"))

# --- country table ------------------------------------------------------------
wide = pd.read_csv(os.path.join(DASH, "output", "global_data_wide.csv")).set_index("iso3")
weo = pd.read_csv(os.path.join(DASH, "raw_data",
                  "dataset_2026-08-26T13_41_09_874550254Z_DEFAULT_INTEGRATION_IMF_RES_WEO_9_0_0.csv"), low_memory=False)
weo["iso"] = weo.SERIES_CODE.str.split(".").str[0].replace({"KOS": "XKX", "WBG": "PSE"})
def wg(ind, y): return pd.to_numeric(weo[weo.INDICATOR == ind].groupby("iso")[y].first(), errors="coerce")
GDPU = "Gross domestic product (GDP), Current prices, US dollar"
GDPP = "Gross domestic product (GDP), Current prices, Purchasing power parity (PPP) international dollar, ICP benchmarks 2017-2021"
GRW = "Gross domestic product (GDP), Constant prices, Percent change"
mvpc = pd.read_csv(os.path.join(MVP, "countries.csv")).set_index("code")
IMF_AE = set("""AND AUS AUT BEL CAN HRV CYP CZE DNK EST FIN FRA DEU GRC HKG ISL IRL ISR ITA JPN KOR LVA LTU
LUX MAC MLT NLD NZL NOR PRT PRI SMR SGP SVK SVN ESP SWE CHE TWN GBR USA""".split())
WB2CPAT = {"Low income": "LIC", "Lower middle income": "LMIC", "Upper middle income": "UMIC", "High income": "HIC"}
# market WACC from the v1.1 workbook (cached values)
v11 = openpyxl.load_workbook(os.path.join(ROOT, "InvestmentNeeds", "Mitigation_Investment_Needs_v1.1.xlsx"),
                             read_only=True, data_only=True)["Country_Data"]
vrows = list(v11.iter_rows(values_only=True))
keys = list(vrows[2]); ki, kr, ks = keys.index("iso"), keys.index("r_a"), keys.index("w_src")
wacc = {r[ki]: (r[kr], r[ks]) for r in vrows[4:] if r[ki]}
gdp, ppp, a25, a26 = wg(GDPU, "2024"), wg(GDPP, "2024"), wg(GRW, "2025"), wg(GRW, "2026")
ct = []
for i in sorted(set(en.iso)):
    inc = mvpc["income_group"].get(i) if i in mvpc.index else None
    if inc not in ("LIC", "LMIC", "UMIC", "HIC"):
        inc = WB2CPAT.get(wide["income"].get(i) if i in wide.index else None, "UMIC")
    nm = wide["country"].get(i) if i in wide.index else cp[cp.iso == i].country.iloc[0]
    ct.append(dict(iso=i, name=nm, group="AE" if i in IMF_AE else ("China" if i == "CHN" else "EMDC"),
                   cpat_income=inc, wb_income=wide["income"].get(i) if i in wide.index else None,
                   gdp24=gdp.get(i), ppp24=ppp.get(i), g25=a25.get(i), g26=a26.get(i),
                   wacc_mkt=wacc.get(i, (None, None))[0], wacc_src=wacc.get(i, (None, None))[1]))
ct = pd.DataFrame(ct)
# fallback GDP from CPAT where WEO is missing
cg = cpv("mit.gdp.pre.lvl.1")
ct["gdp24"] = ct.gdp24.fillna(ct.iso.map(cg))
ct.to_csv(os.path.join(OUT, "countries.csv"), index=False)

# --- CBAM goods (CBAM Policy Tool, Country data) ----------------------------
def _n(s): return re.sub(r"[^a-z]", "", unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower())
name2iso = {_n(r["name"]): r["iso"] for _, r in ct.iterrows()}
name2iso.update({_n(k): v for k, v in {"Iran": "IRN", "Russia": "RUS", "South Korea": "KOR", "Taiwan": "TWN",
                 "Viet Nam": "VNM", "Venezuela": "VEN", "Egypt": "EGY", "Türkiye": "TUR", "Libya": "LBY"}.items()})
ws = openpyxl.load_workbook(os.path.join(REPO, "CBAMPolicyTool", "CBAM_Policy_Tool_global_v1.0.xlsx"),
                            read_only=True, data_only=True)["Country data"]
cr = list(ws.iter_rows(values_only=True))
hdr = cr[3]
GOODS = ["DRI-EAF steel", "Scrap-EAF steel", "Gray clinker", "Ammonia", "Urea", "Ammonium nitrate", "Primary aluminum"]
cb = []
for r in cr[4:]:
    if not r[0] or r[0] == "Egypt": continue          # Egypt uses the study row
    iso = "EGY" if r[0].startswith("Egypt (study") else name2iso.get(_n(r[0]))
    if not iso: print("CBAM: no ISO for", r[0]); continue
    d = dict(iso=iso, growth_pct=r[hdr.index("Growth 2024-30 (% a year)")])
    for gname in GOODS:
        d[f"out|{gname}"] = r[hdr.index(f"Output 2024, {gname} (kt)")]
        d[f"efp|{gname}"] = r[hdr.index(f"EF process CO2, {gname}")]
        d[f"eff|{gname}"] = r[hdr.index(f"EF fuel, {gname}")]
    cb.append(d)
cb = pd.DataFrame(cb)
cb["growth_pct"] = pd.to_numeric(cb.growth_pct, errors="coerce")
cb.to_csv(os.path.join(OUT, "cbam.csv"), index=False)

# --- CPAT elasticities (income and AEEI) ----------------------------------------
el = pd.read_csv(os.path.join(MVP, "elasticities.csv"))
el[el.type.isin(["inc", "aei"])].to_csv(os.path.join(OUT, "elasticities.csv"), index=False)

# --- power market structure (v0.2) --------------------------------------------
# pms: 1 vertically integrated utility, no IPPs; 2 single buyer with IPPs; 3 wholesale competition.
# Order: (i) curated rows for countries the Power Sector Reform Tracker (non-OECD, ends 2013) does
# not cover, (ii) curated updates where PSRT is out of date or codes a market law that never opened
# a market, (iii) PSRT 2013 (3 if r_wem, 2 if r_ipp, else 1, as the dashboard), (iv) PSRT 1 is
# raised to 2 where PPI records two or more private generation projects reaching close 2014-2024.
CURATED = {  # iso: (pms, independent regulator, unbundled, privatised, note)
 "AUS": (3, 1, 1, 1, "NEM wholesale and retail competition (WA separate)"),
 "AUT": (3, 1, 1, 0, "EU internal market; utilities majority public"),
 "BEL": (3, 1, 1, 1, "EU internal market"),
 "BRB": (1, 1, 0, 1, "Private VIU (Barbados Light & Power) under FTC regulation"),
 "CAN": (2, 1, 1, 0, "Mixed: Alberta and Ontario markets; most provinces crown VIUs buying from IPPs"),
 "CHE": (3, 1, 1, 0, "Wholesale market; retail open to large users; cantonal public utilities"),
 "CHL": (3, 1, 1, 1, "Unbundled, privatised, cost-based pool and regulated auctions"),
 "COM": (1, 0, 0, 0, "State VIU"),
 "CPV": (2, 1, 0, 0, "State utility (Electra) buying from wind IPP (Cabeolica)"),
 "CZE": (3, 1, 1, 0, "EU internal market; CEZ majority state"),
 "DEU": (3, 1, 1, 1, "EU internal market"),
 "DNK": (3, 1, 1, 0, "EU internal market (Nord Pool)"),
 "ESP": (3, 1, 1, 1, "EU internal market"),
 "EST": (3, 1, 1, 0, "EU internal market (Nord Pool); Eesti Energia state"),
 "FIN": (3, 1, 1, 0, "EU internal market (Nord Pool)"),
 "FRA": (3, 1, 1, 0, "EU internal market; EDF state-owned"),
 "GBR": (3, 1, 1, 1, "Wholesale and retail competition; private networks"),
 "GRC": (3, 1, 1, 1, "EU internal market; PPC majority private since 2021"),
 "HKG": (1, 0, 0, 1, "Two private VIUs under Scheme of Control agreements"),
 "HUN": (3, 1, 1, 0, "EU internal market; MVM state"),
 "IRL": (3, 1, 1, 0, "Single Electricity Market; ESB state"),
 "ISL": (3, 1, 1, 0, "Unbundled; Landsvirkjun state-dominant"),
 "ISR": (2, 1, 0, 0, "IEC dominant, buying from IPPs; 2018 reform"),
 "ITA": (3, 1, 1, 1, "EU internal market"),
 "JPN": (3, 1, 1, 1, "JEPX wholesale; retail fully open 2016; legal unbundling of networks 2020"),
 "KNA": (1, 0, 0, 0, "State VIU (SKELEC)"),
 "KOR": (2, 0, 1, 0, "KEPCO single buyer from its generation subsidiaries and IPPs (cost-based pool)"),
 "LUX": (3, 1, 1, 0, "EU internal market"),
 "MAC": (1, 1, 0, 1, "Private concession VIU (CEM)"),
 "MEX": (3, 1, 1, 0, "Wholesale market (MEM) since 2016; CFE dominant; 2025 law restores CFE priority"),
 "MLT": (1, 1, 0, 0, "Enemalta VIU (EU derogation)"),
 "NLD": (3, 1, 1, 1, "EU internal market; networks public"),
 "NOR": (3, 1, 1, 0, "Nord Pool; Statkraft state"),
 "NRU": (1, 0, 0, 0, "State VIU"),
 "NZL": (3, 1, 1, 0, "Wholesale market; gentailers majority state-owned"),
 "PLW": (1, 0, 0, 0, "State VIU (PPUC)"),
 "POL": (3, 1, 1, 0, "EU internal market; utilities majority state"),
 "PRT": (3, 1, 1, 1, "EU internal market (MIBEL)"),
 "SLB": (1, 0, 0, 0, "State VIU (Solomon Power)"),
 "SOM": (1, 0, 0, 1, "No national grid; private local VIUs"),
 "STP": (1, 0, 0, 0, "State VIU (EMAE)"),
 "SVK": (3, 1, 1, 0, "EU internal market"),
 "SVN": (3, 1, 1, 0, "EU internal market"),
 "SWE": (3, 1, 1, 0, "Nord Pool; Vattenfall state"),
 "TUR": (3, 1, 1, 1, "EPIAS wholesale market; privatised distribution"),
 "USA": (3, 1, 1, 1, "Mixed: ~two-thirds of load in ISO/RTO markets; rest regulated VIUs"),
 "VCT": (1, 0, 0, 0, "State VIU (VINLEC)"),
 "WSM": (2, 1, 0, 0, "State utility (EPC) buying from solar IPPs"),
}
UPDATES = {  # iso: (pms, note) - curated updates to PSRT 2013
 "VNM": (3, "Vietnam Wholesale Electricity Market operating since 2019"),
 "AFG": (2, "PSRT codes a market; DABS buys imports and IPP power (single buyer)"),
 "QAT": (2, "PSRT codes a market; Kahramaa single buyer from IWPPs"),
 "EGY": (2, "2015 Electricity Law allows a market; EETC still single buyer from IPPs"),
 "ARE": (2, "Single buyers (EWEC, DEWA) from IWPPs"),
 "OMN": (2, "Nama PWP single buyer; spot market from 2022 is small"),
 "JOR": (2, "NEPCO single buyer from IPPs"),
 "PAK": (2, "CPPA-G single buyer; competitive market (CTBCM) only starting"),
 "NGA": (2, "Unbundled and privatised, but NBET bulk trader buys from gencos"),
 "GHA": (2, "Wholesale market rules (2016) not operating; ECG buys from IPPs"),
 "MNG": (2, "Single buyer model"),
 "DMA": (1, "Private VIU (DOMLEC)"),
}
z = zipfile.ZipFile(os.path.join(DASH, "raw_data", "dataverse_files.zip"))
psrt = pd.read_stata(z.open([n for n in z.namelist() if n.endswith("PSRT_V1_Database.dta")][0]))
psrt = psrt[psrt.year == psrt.year.max()].copy()
PNAME = {"Cote d'Ivoire": "CIV", "Macedonia, FYR": "MKD", "Swaziland": "SWZ", "Vietnam": "VNM", "Cape Verde": "CPV",
         "Guyana, CR": "GUY", "Somalia": "SOM", "St. Vincent and Grenadines": "VCT", "Turkiye": "TUR"}
nm2iso = {r["name"]: r["iso"] for _, r in ct.iterrows()}
iso_of = lambda n: PNAME.get(str(n), nm2iso.get(str(n)))
psrt["iso"] = psrt.cntry.map(iso_of)
psrt = psrt.dropna(subset=["iso"]).set_index("iso")
ppi = pd.read_stata(os.path.join(DASH, "raw_data", "2024-PPI-Full-DTA.dta"))
ppi = ppi[(ppi.ssector == "Electricity") & (ppi.status_n != "Cancelled")].copy()
ppi["iso"] = ppi.country.map(iso_of)
ppi["inv24"] = pd.to_numeric(ppi.investment, errors="coerce") * pd.to_numeric(ppi.CPI2024, errors="coerce")
ppi["IY"] = pd.to_numeric(ppi.IY, errors="coerce")
seg = ppi.Segment.astype(str)
rec = ppi[ppi.IY.between(2015, 2024)]
gen = rec[seg.loc[rec.index].str.contains("generation")]
dist = ppi[seg.str.contains("distribution") & (ppi.status_n == "Active")]
tra = ppi[seg.str.contains("transmission") & (ppi.status_n == "Active")]
ipp14 = ppi[ppi.IY.between(2014, 2024) & seg.str.contains("generation")].groupby("iso").size()
wide_pm = wide.reindex(ct.iso)
cp_ = lambda m: cpv(m)
rp = {k: cp_(f"mit.rp.ecy.{k}.1") for k in ("res", "ind", "other")}
sc_ = {k: cp_(f"mit.sup.cost.ecy.{k}.1") for k in ("res", "ind", "other")}
cn = {k: cp_(f"mit.con.ecy.{k}.1") for k in ("res", "ind", "other")}
esub = cp_("mit.expsub.con.ecy.all.1"); gdpc = cp_("mit.gdp.pre.lvl.1")
pm = []
for _, c in ct.iterrows():
    i = c.iso
    if i in CURATED:
        s_, reg, und, prv, note = CURATED[i]; src = "Curated (not in PSRT)"; cos = np.nan
    elif i in psrt.index:
        r = psrt.loc[i]
        s_ = 3 if r.r_wem == 1 else (2 if r.r_ipp == 1 else 1)
        reg, und, prv, cos = int(r.r_reg), int(r.r_und), int(r.r_prv), int(r.r_cos)
        src, note = "PSRT 2013", ""
        if i in UPDATES:
            s_, note = UPDATES[i]; src = "PSRT 2013, curated update"
        elif s_ == 1 and ipp14.get(i, 0) >= 2:
            s_, src, note = 2, "PSRT 2013, PPI update", f"{int(ipp14.get(i))} private generation projects 2014-2024 (PPI)"
    else:
        reg = und = prv = cos = np.nan
        s_ = 2 if ipp14.get(i, 0) >= 2 else 1
        src, note = "PPI rule (no PSRT)", f"{int(ipp14.get(i, 0))} private generation projects 2014-2024 (PPI)"
    num = sum(float(rp[k].get(i, np.nan) or 0) * float(cn[k].get(i, np.nan) or 0) for k in rp
              if pd.notna(rp[k].get(i)) and pd.notna(sc_[k].get(i)) and pd.notna(cn[k].get(i)))
    den = sum(float(sc_[k].get(i)) * float(cn[k].get(i)) for k in rp
              if pd.notna(rp[k].get(i)) and pd.notna(sc_[k].get(i)) and pd.notna(cn[k].get(i)))
    g = gen[gen.iso == i]
    gdp_i = c.gdp24 if pd.notna(c.gdp24) else gdpc.get(i)
    pm.append(dict(
        iso=i, pms=s_, pms_src=src, pms_note=note, reg_indep=reg, unbundled=und, privatised=prv, cost_reflective_psrt=cos,
        ppi_elec_bn_yr=rec[rec.iso == i].inv24.sum() / 10 / 1000,
        ppi_gen_n_1524=len(g),
        ppi_vre_share=(g[g.technol.astype(str).str.contains("Solar|Wind")].inv24.sum() / g.inv24.sum()) if g.inv24.sum() > 0 else np.nan,
        ppi_merchant_share=(g[(g.PRS.astype(str) == "Sales to wholesale market") | (g.stype.astype(str) == "Merchant")].inv24.sum()
                            / g.inv24.sum()) if g.inv24.sum() > 0 else np.nan,
        ppi_private_dist=int((dist.iso == i).any()), ppi_private_trans=int((tra.iso == i).any()),
        cr_ecy=(num / den) if den > 0 else np.nan,
        ecy_sub_pct_gdp=(float(esub.get(i)) / float(gdp_i) * 100) if pd.notna(esub.get(i)) and gdp_i else np.nan,
        ffs_exp_pct_gdp=wide_pm.ff_subsidy_explicit_pct_gdp.get(i),
        rise_re=wide_pm.rise_renewable_energy_score.get(i), rise_lpf=wide_pm.rise_level_playing_field_score.get(i)))
pm = pd.DataFrame(pm)
pm["ppi_elec_pct_gdp"] = pm.ppi_elec_bn_yr / pm.iso.map(ct.set_index("iso").gdp24) * 100
pm.to_csv(os.path.join(OUT, "power_market.csv"), index=False)
print("power market structure:", pm.pms.value_counts().sort_index().to_dict(), pm.pms_src.value_counts().to_dict())

print(f"countries {len(ct)}, energy rows {len(en)}, power {len(pw)}, CBAM {len(cb)}")
