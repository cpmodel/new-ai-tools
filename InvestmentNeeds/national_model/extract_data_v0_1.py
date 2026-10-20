#!/usr/bin/env python3
# =============================================================================
# NMIM v0.1 - data extraction
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
#
# Sources (all in the repository, no network):
#   TaxMacroFiscalDashboard/build/output/imffossilfuelsubsidiesdata_sheets.xlsx (CPAT panel)
#   TaxMacroFiscalDashboard/build/raw_data/ WEO csv, Ember csv
#   CPAT_excel_not_master/mitigation_copypaste/data/ elasticities.csv, countries.csv, fuel_units.csv
#   CBAMPolicyTool/CBAM_Policy_Tool_global_v1.0.xlsx, sheet "Country data"
#   InvestmentNeeds/Mitigation_Investment_Needs_v1.1.xlsx, sheet Country_Data (WACC)
#
# Run: python InvestmentNeeds/national_model/extract_data_v0_1.py
# =============================================================================
import os, re, unicodedata
import numpy as np, pandas as pd, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
DASH = os.path.join(ROOT, "TaxMacroFiscalDashboard", "build")
MVP = os.path.join(ROOT, "CPAT_excel_not_master", "mitigation_copypaste", "data")
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
ws = openpyxl.load_workbook(os.path.join(ROOT, "CBAMPolicyTool", "CBAM_Policy_Tool_global_v1.0.xlsx"),
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

print(f"countries {len(ct)}, energy rows {len(en)}, power {len(pw)}, CBAM {len(cb)}")
