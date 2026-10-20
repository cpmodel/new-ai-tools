#!/usr/bin/env python3
# =============================================================================
# NMIM v0.2 - National Mitigation Investment Model: equation spec and engine
# =============================================================================
# One equation spec drives both implementations:
#   - this engine evaluates it for every country (Python), and
#   - build_workbook_v0_2.py renders the same templates as Excel formulas for the
#     country chosen on the Settings sheet (CPAT Mitigation layout).
# So the workbook and the engine cannot drift apart; check_v0_2.py confirms it.
#
# Template grammar (used in ROWS and SCALARS):
#   [code]        this year's value of row `code`
#   [code@-1]     last year's value
#   [code@N]      the base-year (2024) value
#   LAG(code,n)   value n years earlier (0 before 2025)
#   {name}        parameter (Inputs sheet, NameOfParameter `name`)
#   #name         country scalar (Mitigation column J, section 1)
#   @field        country data field (Data sheet)
#   YEAR          the year
#   Excel functions MAX, MIN, IF, CHOOSE, ABS; ^ for powers; = for equality.
#   v0.2 adds section 12, investment by type (who pays), from country power market data.
#
# Codes follow CPAT: mit.<indicator>.<fuel>.<sector>[.<scenario>], scenario 1 =
# baseline (current trends), 2 = net zero.
#
# Run: python InvestmentNeeds/national_model/nmim_v0_2.py
#      -> results/nmim_results_v0_2.csv (long), results/nmim_summary_v0_2.csv
# =============================================================================
import os, re, math
import numpy as np, pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
RES = os.path.join(HERE, "results")
VERSION = "0.2"
Y0, Y1 = 2024, 2050
YEARS = list(range(Y0, Y1 + 1))
HIST = [2021, 2022, 2023]

# -----------------------------------------------------------------------------
# PARAMETERS: name, value, unit, description, source, fit (G/M/P)
# -----------------------------------------------------------------------------
PARAMS = [
 # macro
 ("g_lr_lic", 4.5, "% a year", "Long-run real GDP growth, low income", "Assumption, SSP2-like convergence. †", "P"),
 ("g_lr_lmic", 4.0, "% a year", "Long-run real GDP growth, lower middle income", "Assumption. †", "P"),
 ("g_lr_umic", 3.0, "% a year", "Long-run real GDP growth, upper middle income", "Assumption. †", "P"),
 ("g_lr_hic", 1.5, "% a year", "Long-run real GDP growth, high income", "Assumption. †", "P"),
 ("g_conv_year", 2040, "year", "Year growth reaches its long-run rate", "Assumption", "M"),
 ("phi", 0.5, "elasticity", "Pass-through of local price level to capital cost", "As v1.1 workbook. †", "P"),
 # targets by group (AE / China / EMDC)
 ("pow_T_ae", 2035, "year", "Fossil-free power, advanced economies", "IEA Net Zero by 2050 (2021): AE power net zero 2035", "M"),
 ("pow_T_chn", 2040, "year", "Fossil-free power, China", "IEA NZE: global power net zero 2040", "M"),
 ("pow_T_emdc", 2045, "year", "Fossil-free power, other EMDCs", "Assumption: five years after the global NZE date. †", "P"),
 ("ev_T_ae", 2035, "year", "100% EV car sales, advanced economies", "IEA NZE: no new ICE car sales from 2035", "M"),
 ("ev_T_chn", 2035, "year", "100% EV car sales, China", "IEA NZE", "M"),
 ("ev_T_emdc", 2040, "year", "100% EV car sales, other EMDCs", "Assumption: five years later. †", "P"),
 ("bld_T_ae", 2045, "year", "Fossil heating fully replaced, AE", "Assumption. †", "P"),
 ("bld_T_chn", 2050, "year", "Fossil heating fully replaced, China", "Assumption. †", "P"),
 ("bld_T_emdc", 2050, "year", "Fossil heating fully replaced, EMDC", "Assumption. †", "P"),
 ("ind_a_ae", 0.90, "share", "Industrial fossil use abated by 2050, AE", "Assumption (IEA NZE heavy industry ~90%). †", "M"),
 ("ind_a_chn", 0.90, "share", "Industrial fossil use abated by 2050, China", "Assumption. †", "M"),
 ("ind_a_emdc", 0.80, "share", "Industrial fossil use abated by 2050, EMDC", "Assumption. †", "P"),
 # baseline trends
 ("ev0_ae", 0.15, "share", "EV share of car sales 2024, AE", "IEA Global EV Outlook 2024, approx. (Europe 21%, US 10%). †", "M"),
 ("ev0_chn", 0.40, "share", "EV share of car sales 2024, China", "IEA Global EV Outlook 2024, approx. †", "M"),
 ("ev0_emdc", 0.03, "share", "EV share of car sales 2024, EMDC", "IEA Global EV Outlook 2024, approx. †", "M"),
 ("evb35_ae", 0.50, "share", "Baseline EV sales share 2035, AE", "Assumption, current-policies trend. †", "P"),
 ("evb35_chn", 0.80, "share", "Baseline EV sales share 2035, China", "Assumption. †", "P"),
 ("evb35_emdc", 0.15, "share", "Baseline EV sales share 2035, EMDC", "Assumption. †", "P"),
 ("sig0_ratio", 0.33, "ratio", "EV stock share / EV sales share, 2024", "Assumption (sales rose over ~5 years). †", "P"),
 ("fos_trend_cap", 2.0, "pp a year", "Cap on baseline fall in fossil generation share", "Assumption; trend from Ember 3-year change", "M"),
 # energy service ratios
 ("ev_ratio", 0.25, "ratio", "EV final energy per unit of ICE fuel displaced", "0.18 kWh/km vs 7.5 l/100 km gasoline", "M"),
 ("hp_ratio", 0.33, "ratio", "Heat pump electricity per unit of fossil heat fuel", "COP ~3", "M"),
 ("ind_elec_sh", 0.60, "share", "Share of abated industrial fuel that is electrified", "Assumption (rest H2, CCS, bio). †", "P"),
 ("ind_ratio", 0.80, "ratio", "Electricity per unit of industrial fuel electrified", "Assumption. †", "P"),
 # vehicles
 ("veh_gj", 38.5, "GJ/vehicle-yr", "Fuel per car-equivalent vehicle a year", "15,000 km x 7.5 l/100 km x 34.2 MJ/l", "M"),
 ("veh_life", 12, "years", "Vehicle life", "SISEPUEDE (Baik et al. 2019)", "M"),
 ("ev_prem24", 10000, "USD/vehicle", "EV purchase premium 2024", "Assumption; SISEPUEDE USD 12,000 (2019). †", "P"),
 ("ev_parity", 2032, "year", "Year the EV premium reaches zero", "Assumption; IEA reports near parity in China by 2024. †", "P"),
 ("charger", 1000, "USD/vehicle", "Charging infrastructure per EV", "SISEPUEDE (Baik et al. 2019)", "M"),
 ("ev_maint", 0.012, "USD/km", "EV maintenance saving", "SISEPUEDE (AAA 2019)", "M"),
 ("veh_km", 15000, "km/yr", "Vehicle-km a year", "SISEPUEDE", "M"),
 # buildings and industry capital
 ("k_bld", 200, "USD per GJ/yr", "Capital per GJ a year of fossil heat replaced", "Assumption: heat pump ~USD 10,000 for ~50 GJ/yr. †", "P"),
 ("k_ind", 150, "USD per GJ/yr", "Capital per GJ a year of industrial fossil use abated", "Assumption. †", "P"),
 ("k_ccs", 800, "USD per tCO2/yr", "Capital per tonne a year of CBAM process CO2 abated", "Assumption, CCS on cement scale. †", "P"),
 ("cbam_a_2050", 0.90, "share", "CBAM-goods process CO2 abated by 2050 (net zero)", "Assumption. †", "M"),
 # power
 ("capex_sol24", 760, "USD/kW", "Utility solar PV capital cost 2024", "IRENA (2024) Renewable Power Generation Costs in 2023. †", "G"),
 ("capex_wnd24", 1160, "USD/kW", "Onshore wind capital cost 2024", "IRENA (2024). †", "G"),
 ("capex_gas", 1304, "USD/kW", "Gas power capital cost (new fossil build)", "CPAT legacy PowerTech (Egypt values)", "M"),
 ("lr_sol", 0.03, "a year", "Real fall in solar capital cost", "Assumption from IRENA/IEA trends. †", "M"),
 ("lr_wnd", 0.015, "a year", "Real fall in wind capital cost", "Assumption. †", "M"),
 ("cf_sol", 0.17, "share", "Solar capacity factor where Ember capacity is missing", "Engineering norm. †", "M"),
 ("cf_wnd", 0.35, "share", "Wind capacity factor where Ember capacity is missing", "Engineering norm. †", "M"),
 ("cf_fos", 0.50, "share", "Fossil capacity factor where Ember capacity is missing", "Engineering norm. †", "M"),
 ("sh_sol", 0.60, "share", "Solar share of new VRE generation", "Assumption. †", "P"),
 ("life_vre", 25, "years", "VRE life", "Method note", "G"),
 ("life_fos", 30, "years", "Fossil plant life", "CPAT PowerTech", "G"),
 ("om_vre", 0.015, "share/yr", "VRE O&M, share of capital a year", "Engineering norm. †", "M"),
 ("om_fos", 0.0105, "USD/kWh", "Fossil O&M (fixed + variable)", "CPAT PowerTech gas: 0.0065 + 0.004", "M"),
 ("grid_ratio", 0.6, "USD/USD", "Grid investment per USD of generation investment", "IEA (2023) grids: 0.6-1 per USD of renewables (low end)", "M"),
 ("st_h", 2, "hours", "Storage hours per kW of VRE added at full VRE", "CPAT PowerParams st_ratio (2 kWh per kW). †", "M"),
 ("capex_st24", 250, "USD/kWh", "Battery storage capital cost 2024", "IRENA/BNEF 250-300 (method note Table C). †", "M"),
 ("lr_st", 0.03, "a year", "Real fall in storage capital cost", "Assumption. †", "M"),
 ("hr_coa", 10.0, "GJ/MWh", "Heat rate, coal power", "Engineering norm. †", "M"),
 ("hr_nga", 7.5, "GJ/MWh", "Heat rate, gas power", "Engineering norm. †", "M"),
 ("hr_oil", 10.5, "GJ/MWh", "Heat rate, oil power", "Engineering norm. †", "M"),
 # emission factors (IPCC 2006 Vol. 2 Table 2.2, tCO2/GJ)
 ("ef_coa", 0.0946, "tCO2/GJ", "CO2 factor, coal", "IPCC 2006 Table 2.2 (other bituminous)", "G"),
 ("ef_nga", 0.0561, "tCO2/GJ", "CO2 factor, natural gas", "IPCC 2006 Table 2.2", "G"),
 ("ef_gso", 0.0693, "tCO2/GJ", "CO2 factor, gasoline", "IPCC 2006 Table 2.2", "G"),
 ("ef_die", 0.0741, "tCO2/GJ", "CO2 factor, diesel", "IPCC 2006 Table 2.2", "G"),
 ("ef_lpg", 0.0631, "tCO2/GJ", "CO2 factor, LPG", "IPCC 2006 Table 2.2", "G"),
 ("ef_ker", 0.0719, "tCO2/GJ", "CO2 factor, kerosene", "IPCC 2006 Table 2.2", "G"),
 ("ef_oil", 0.0774, "tCO2/GJ", "CO2 factor, other oil (residual fuel oil)", "IPCC 2006 Table 2.2", "G"),
 # non-CO2, process, land
 ("nc_2030", 0.30, "share", "Non-CO2 and other process GHG cut by 2030", "Global Methane Pledge (-30% methane by 2030)", "M"),
 ("nc_2050", 0.60, "share", "Non-CO2 and other process GHG cut by 2050", "Assumption. †", "P"),
 ("cpt_nc", 50, "USD/tCO2e", "Cost per tonne, non-CO2 abatement", "As v1.1 (IPCC AR6 SPM C.12). †", "M"),
 ("cpt_afolu", 50, "USD/tCO2e", "Cost per tonne, land-use emissions", "As v1.1 (IPCC AR6 SPM C.9.1). †", "P"),
 ("cpt_res", 100, "USD/tCO2", "Cost per tonne, other energy CO2 (aviation, shipping, other)", "IPCC AR6 SPM C.12 threshold. †", "P"),
 ("lu_T", 2030, "year", "Net land-use emissions brought to zero", "Glasgow Leaders' Declaration on Forests (2021)", "M"),
 # baseline calibration to current clean-energy investment (IEA World Energy Investment 2024)
 ("iea24_ae", 1005, "USD bn", "Clean energy investment 2024, advanced economies (residual)",
  "IEA WEI 2024: world ~2,000 less China 675 less EMDE ex-China ~320", "M"),
 ("iea24_chn", 675, "USD bn", "Clean energy investment 2024, China", "IEA World Energy Investment 2024", "M"),
 ("iea24_emdc", 320, "USD bn", "Clean energy investment 2024, EMDEs other than China", "IEA World Energy Investment 2024", "M"),
 # finance
 ("r_wb", 0.06, "rate", "World Bank social discount rate", "Fay et al. (2016)", "G"),
 ("life_bld", 25, "years", "Asset life, buildings", "Method note", "M"),
 ("life_ind", 25, "years", "Asset life, industry", "Method note", "M"),
 ("life_grid", 40, "years", "Asset life, grids and storage", "Method note", "M"),
 # investment typology (who pays): P1 pure public, P2 regulated utility, P3 contracted revenue,
 # P4 viability gap, P5 pure private. Shares are judgements (†) applied to country data.
 ("ty_g2_vi", 0.70, "share", "Generation, utility-built (P2), vertically integrated market", "Assumption. †", "P"),
 ("ty_g2_sb", 0.35, "share", "Generation, utility-built (P2), single buyer with IPPs", "Assumption. †", "P"),
 ("ty_g2_wm", 0.10, "share", "Generation, utility-built (P2), wholesale market", "Assumption. †", "P"),
 ("ty_g3_vi", 0.25, "share", "Generation, contracted IPPs (P3), vertically integrated market", "Assumption. †", "P"),
 ("ty_g3_sb", 0.55, "share", "Generation, contracted IPPs (P3), single buyer", "Assumption; IEA-IFC (2023) PPAs dominate EMDE renewables. †", "P"),
 ("ty_g3_wm", 0.55, "share", "Generation, contracted (P3: CfDs, auctions, corporate PPAs), wholesale market", "Assumption. †", "P"),
 ("ty_r3_vi", 0.00, "share", "Grids and storage, private contracted (P3), vertically integrated", "Assumption. †", "P"),
 ("ty_r3_sb", 0.05, "share", "Grids and storage, private contracted (P3), single buyer", "Assumption; PPI private transmission rare. †", "P"),
 ("ty_r3_wm", 0.15, "share", "Grids and storage, private contracted (P3), wholesale market", "Assumption: merchant/contracted storage, OFTO-type lines. †", "P"),
 ("ty_cr_gap", 0.50, "ratio", "Tariff under-recovery at which the utility pressure index reaches 1", "Assumption: cost recovery 50% or less. †", "P"),
 ("ty_ku1", 0.50, "share", "Share of utility generation the budget pays at full under-recovery (P2 to P1)", "Assumption. †", "P"),
 ("ty_ku4", 0.40, "share", "Share of contracted generation needing a public top-up at full under-recovery (P3 to P4)", "Assumption. †", "P"),
 ("ty_kur", 0.60, "share", "Share of utility grid investment the budget pays at full under-recovery (P2 to P1)", "Assumption. †", "P"),
 ("ty_acc_lic", 0.25, "share", "Grid investment for access paid by the budget, low income", "Assumption; access extension has no paying customers at first. †", "P"),
 ("ty_acc_lmic", 0.10, "share", "Grid investment for access paid by the budget, lower middle income", "Assumption. †", "P"),
 ("ty_h_span", 0.08, "rate", "WACC above 6% at which the capital-cost pressure index reaches 1", "Assumption: WACC 14% or more. †", "P"),
 ("ty_kh", 0.40, "share", "Share of market investment needing a top-up at full capital-cost pressure (P5 to P4)", "Assumption. †", "P"),
 ("ty_v_span", 2.0, "% GDP", "Explicit fossil fuel subsidies at which the subsidy pressure index reaches 1", "Assumption. †", "P"),
 ("ty_kv", 0.20, "share", "Share of market end-use investment needing a top-up at full subsidy pressure", "Assumption. †", "P"),
 ("ty_ci_lic", 0.30, "share", "Credit-constraint top-up share, low income", "Assumption. †", "P"),
 ("ty_ci_lmic", 0.15, "share", "Credit-constraint top-up share, lower middle income", "Assumption. †", "P"),
 ("ty_t1_lic", 0.15, "share", "Transport, public (P1: public buses, fleets), low income", "Assumption. †", "P"),
 ("ty_t1_lmic", 0.10, "share", "Transport, public (P1), lower middle income", "Assumption. †", "P"),
 ("ty_t1_umic", 0.05, "share", "Transport, public (P1), upper middle income", "Assumption. †", "P"),
 ("ty_t1_hic", 0.05, "share", "Transport, public (P1), high income", "Assumption. †", "P"),
 ("ty_t3", 0.10, "share", "Transport, contracted (P3: charging concessions, bus leasing)", "Assumption. †", "P"),
 ("ty_b1_lic", 0.40, "share", "Buildings, public (P1: public buildings, social housing), low income", "Assumption. †", "P"),
 ("ty_b1_lmic", 0.40, "share", "Buildings, public (P1), lower middle income", "Assumption. †", "P"),
 ("ty_b1_umic", 0.35, "share", "Buildings, public (P1), upper middle income", "Assumption. †", "P"),
 ("ty_b1_hic", 0.30, "share", "Buildings, public (P1), high income", "Assumption. †", "P"),
 ("ty_i4", 0.30, "share", "Industrial fuel switch needing a top-up (P4) with no country pressure", "Assumption; CCS on CBAM goods is all P4. †", "P"),
 ("ty_n1_lic", 0.30, "share", "Non-CO2, public (P1: waste, agriculture extension), low income", "Assumption. †", "P"),
 ("ty_n1_lmic", 0.25, "share", "Non-CO2, public (P1), lower middle income", "Assumption. †", "P"),
 ("ty_n1_umic", 0.20, "share", "Non-CO2, public (P1), upper middle income", "Assumption. †", "P"),
 ("ty_n1_hic", 0.20, "share", "Non-CO2, public (P1), high income", "Assumption. †", "P"),
 ("ty_n5", 0.35, "share", "Non-CO2, private (P5: oil and gas methane, captured gas sold) before pressure", "Assumption; IEA Global Methane Tracker. †", "P"),
 ("ty_l1_lic", 0.80, "share", "Land use, public (P1: forest protection), low income; rest P4", "Assumption. †", "P"),
 ("ty_l1_lmic", 0.70, "share", "Land use, public (P1), lower middle income", "Assumption. †", "P"),
 ("ty_l1_umic", 0.60, "share", "Land use, public (P1), upper middle income", "Assumption. †", "P"),
 ("ty_l1_hic", 0.50, "share", "Land use, public (P1), high income", "Assumption. †", "P"),
 ("ty_cal_pow", 0.50, "share", "Power share of unmodelled clean investment (rest end use)", "IEA WEI 2024: ~half power and grids. †", "M"),
 ("ty_p4_gap", 0.30, "share", "Public share of P4 capital (the viability gap)", "Assumption. †", "P"),
]
P = {p[0]: p[1] for p in PARAMS}

# -----------------------------------------------------------------------------
# DATA FIELDS (Data sheet columns; read by @field)
# -----------------------------------------------------------------------------
ACT = [("gso", "rod"), ("die", "rod"), ("coa", "res"), ("nga", "res"), ("lpg", "res"), ("ker", "res"), ("ecy", "res"),
       ("coa", "ind"), ("nga", "ind"), ("oil", "ind"), ("ecy", "ind")]
EL = {  # activity row -> (CPAT elasticity fuel, CPAT elasticity sector)
 ("gso", "rod"): ("gso", "tra"), ("die", "rod"): ("die", "tra"), ("coa", "res"): ("coa", "res"), ("nga", "res"): ("nga", "res"),
 ("lpg", "res"): ("oil", "res"), ("ker", "res"): ("oil", "res"), ("ecy", "res"): ("ecy", "res"),
 ("coa", "ind"): ("coa", "ind"), ("nga", "ind"): ("nga", "ind"), ("oil", "ind"): ("oil", "ind"), ("ecy", "ind"): ("ecy", "ind")}
AEI_FUEL = {"gso": "oil", "die": "oil", "lpg": "oil", "ker": "oil", "coa": "coa", "nga": "nga", "oil": "oil", "ecy": "ecy"}
PRICED = [("gso", "rod"), ("die", "rod"), ("coa", "res"), ("nga", "res"), ("lpg", "res"), ("ker", "res"),
          ("coa", "ind"), ("nga", "ind"), ("oil", "ind"), ("coa", "pow"), ("nga", "pow"), ("oil", "pow")]

def load_data():
    """Returns {iso: {field: value}} with every @field the spec uses, plus history."""
    ct = pd.read_csv(os.path.join(DATA, "countries.csv")).set_index("iso")
    en = pd.read_csv(os.path.join(DATA, "energy.csv"))
    pr = pd.read_csv(os.path.join(DATA, "prices.csv"))
    pw = pd.read_csv(os.path.join(DATA, "power.csv")).set_index("iso")
    em = pd.read_csv(os.path.join(DATA, "emissions.csv")).set_index("iso")
    cb = pd.read_csv(os.path.join(DATA, "cbam.csv")).set_index("iso")
    pm = pd.read_csv(os.path.join(DATA, "power_market.csv")).set_index("iso")
    el = pd.read_csv(os.path.join(DATA, "elasticities.csv"))
    elv = {(r.type, r.fuel_group, r.sector_group): r for r in el.itertuples()}
    ENW = en.pivot_table(index=["iso", "year"], columns="code", values="value_pj")
    PRW = pr.pivot(index="iso", columns="code", values="usd_per_gj")
    goods = [c[4:] for c in cb.columns if c.startswith("out|")]
    D = {}
    for iso, c in ct.iterrows():
        if pd.isna(c.gdp24): continue
        d = {"name": c["name"], "grp": {"AE": 1, "China": 2, "EMDC": 3}[c.group],
             "inc": {"LIC": 1, "LMIC": 2, "UMIC": 3, "HIC": 4}[c.cpat_income],
             "gdp24": c.gdp24, "plevel": (c.gdp24 / c.ppp24) if pd.notna(c.ppp24) and c.ppp24 > 0 else 1.0,
             "g25": c.g25 if pd.notna(c.g25) else 3.0, "g26": c.g26 if pd.notna(c.g26) else 3.0,
             "wacc": c.wacc_mkt if pd.notna(c.wacc_mkt) else 0.105}
        ig = c.cpat_income
        for f, s in ACT:
            for y in HIST + [Y0]:
                v = ENW.get(f"mit.con.{f}.{s}").get((iso, y)) if (iso, y) in ENW.index else np.nan
                d[f"con_{f}_{s}" + ("" if y == Y0 else f"_{y}")] = 0.0 if pd.isna(v) else float(v)
            ef, es = EL[(f, s)]
            d[f"inc_{f}_{s}"] = float(getattr(elv[("inc", ef, es)], ig))
            d[f"aei_{f}_{s}"] = float(getattr(elv[("aei", AEI_FUEL[f], es)], ig))
        for f, s in PRICED:
            for k in ("sp", "ap"):
                v = PRW.get(f"mit.{k}.{f}.{s}", pd.Series(dtype=float)).get(iso)
                d[f"{k}_{f}_{s}"] = 0.0 if pd.isna(v) else float(v)
        # power (Ember); fallback from CPAT electricity and renewable share
        dem_twh = (d["con_ecy_res"] + d["con_ecy_ind"]) / 3.6
        if iso in pw.index and pd.notna(pw.at[iso, "gen_tot"]) and pw.at[iso, "gen_tot"] > 0:
            r = pw.loc[iso].fillna(0).clip(lower=0)   # Ember carries a few small negative values
            r["fos_trend_pp_yr"] = pw.at[iso, "fos_trend_pp_yr"] if pd.notna(pw.at[iso, "fos_trend_pp_yr"]) else 0.0
            gen = r.gen_tot
            d.update(gen_tot=gen, gen_fos=r.gen_coa + r.gen_nga + r.gen_oil, gen_sol=r.gen_sol, gen_wnd=r.gen_wnd,
                     gen_cfirm=r.gen_hyd + r.gen_nuc + r.gen_bio + r.gen_ore,
                     gen_coa=r.gen_coa, gen_nga=r.gen_nga, gen_oil=r.gen_oil,
                     cap_sol=r.cap_sol, cap_wnd=r.cap_wnd, cap_fos=r.cap_coa + r.cap_nga + r.cap_oil,
                     fos_trend=r.fos_trend_pp_yr)
        else:
            gen = dem_twh * 1.1
            d.update(gen_tot=gen, gen_fos=gen * 0.8, gen_sol=0.0, gen_wnd=0.0, gen_cfirm=gen * 0.2,
                     gen_coa=0.0, gen_nga=0.0, gen_oil=gen * 0.8, cap_sol=0.0, cap_wnd=0.0, cap_fos=0.0, fos_trend=0.0)
        d["dem_twh24"] = dem_twh
        d["con_coa_pow"] = float(ENW.get("mit.con.coa.pow").get((iso, Y0), 0) or 0) if (iso, Y0) in ENW.index else 0.0
        d["con_nga_pow"] = float(ENW.get("mit.con.nga.pow").get((iso, Y0), 0) or 0) if (iso, Y0) in ENW.index else 0.0
        # emissions and CBAM goods
        d["co2_24"] = float(em.at[iso, "co2"]) if iso in em.index and pd.notna(em.at[iso, "co2"]) else 0.0
        d["ghg_exc"] = float(em.at[iso, "ghg_exc"]) if iso in em.index and pd.notna(em.at[iso, "ghg_exc"]) else d["co2_24"]
        d["ghg_inc"] = float(em.at[iso, "ghg_inc"]) if iso in em.index and pd.notna(em.at[iso, "ghg_inc"]) else d["ghg_exc"]
        if iso in cb.index:
            r = cb.loc[iso]
            d["cbam_proc24"] = sum(float(r[f"out|{g}"] or 0) * float(r[f"efp|{g}"] or 0) for g in goods) / 1000
            d["cbam_g"] = float(r.growth_pct) if pd.notna(r.growth_pct) else 3.0
        else:
            d["cbam_proc24"], d["cbam_g"] = 0.0, 0.0
        d["cal24"] = 0.0
        # power market and pricing (typology); missing cost recovery = full recovery
        q = pm.loc[iso] if iso in pm.index else None
        d["pms"] = float(q.pms) if q is not None else 2.0
        d["cr_ecy"] = float(q.cr_ecy) if q is not None and pd.notna(q.cr_ecy) else 1.0
        d["ffs_exp"] = float(q.ffs_exp_pct_gdp) if q is not None and pd.notna(q.ffs_exp_pct_gdp) else 0.0
        D[iso] = d
    return D

# -----------------------------------------------------------------------------
# SCALARS (Mitigation section 1, column J): key, description, unit, template
# -----------------------------------------------------------------------------
SCALARS = [
 ("grp", "IHLEG group (1 AE, 2 China, 3 EMDC)", "code", "@grp"),
 ("inc", "CPAT income group (1 LIC, 2 LMIC, 3 UMIC, 4 HIC)", "code", "@inc"),
 ("gdp24", "GDP 2024", "USD bn", "@gdp24"),
 ("g25", "Real GDP growth 2025 (WEO)", "%", "@g25"),
 ("g26", "Real GDP growth 2026 (WEO)", "%", "@g26"),
 ("glr", "Long-run real GDP growth", "%", "CHOOSE(#inc,{g_lr_lic},{g_lr_lmic},{g_lr_umic},{g_lr_hic})"),
 ("pf", "Local price factor on capital (price level^phi)", "x", "@plevel^{phi}"),
 ("wacc", "Market WACC (v1.1: observed, Damodaran or IHLEG default)", "rate", "@wacc"),
 ("pow_T", "Fossil-free power year", "year", "CHOOSE(#grp,{pow_T_ae},{pow_T_chn},{pow_T_emdc})"),
 ("ev_T", "100% EV sales year", "year", "CHOOSE(#grp,{ev_T_ae},{ev_T_chn},{ev_T_emdc})"),
 ("bld_T", "Fossil heating replaced by", "year", "CHOOSE(#grp,{bld_T_ae},{bld_T_chn},{bld_T_emdc})"),
 ("ind_a", "Industrial fossil use abated by 2050", "share", "CHOOSE(#grp,{ind_a_ae},{ind_a_chn},{ind_a_emdc})"),
 ("ev0", "EV sales share 2024", "share", "CHOOSE(#grp,{ev0_ae},{ev0_chn},{ev0_emdc})"),
 ("evb35", "Baseline EV sales share 2035", "share", "CHOOSE(#grp,{evb35_ae},{evb35_chn},{evb35_emdc})"),
 ("sig0", "EV stock share 2024", "share", "#ev0*{sig0_ratio}"),
 ("gcr", "Generation / end-use electricity, 2024", "ratio", "MIN(1.4,MAX(0.9,@gen_tot/MAX(0.001,@dem_twh24)))"),
 ("f0", "Fossil share of generation 2024", "share", "MIN(1,@gen_fos/MAX(0.001,@gen_tot))"),
 ("ftrend", "Baseline fall in fossil share", "share/yr", "MIN({fos_trend_cap},MAX(0,@fos_trend))/100"),
 ("cfirm", "Clean firm generation (hydro, nuclear, bio, other)", "TWh", "@gen_cfirm"),
 ("sh_coa", "Coal share of fossil generation", "share", "IF(@gen_fos>0,@gen_coa/@gen_fos,0)"),
 ("sh_nga", "Gas share of fossil generation", "share", "IF(@gen_fos>0,@gen_nga/@gen_fos,0)"),
 ("sh_oil", "Oil share of fossil generation", "share", "IF(@gen_fos>0,@gen_oil/@gen_fos,1)"),
 ("hr_coa", "Coal heat rate, calibrated to CPAT power fuel use", "GJ/MWh",
  "IF(@gen_coa>0.1,MIN(14,MAX(7,@con_coa_pow/@gen_coa)),{hr_coa})"),
 ("hr_nga", "Gas heat rate, calibrated to CPAT power fuel use", "GJ/MWh",
  "IF(@gen_nga>0.1,MIN(12,MAX(5.5,@con_nga_pow/@gen_nga)),{hr_nga})"),
 ("cf_sol", "Solar capacity factor", "share", "IF(@cap_sol>0.05,MIN(0.3,MAX(0.1,@gen_sol/(@cap_sol*8.76))),{cf_sol})"),
 ("cf_wnd", "Wind capacity factor", "share", "IF(@cap_wnd>0.05,MIN(0.5,MAX(0.15,@gen_wnd/(@cap_wnd*8.76))),{cf_wnd})"),
 ("cf_fos", "Fossil capacity factor", "share", "IF(@cap_fos>0.05,MIN(0.8,MAX(0.2,@gen_fos/(@cap_fos*8.76))),{cf_fos})"),
 ("nc24", "Non-CO2 and other process GHG 2024", "MtCO2e", "MAX(0,@ghg_exc-@co2_24-@cbam_proc24)"),
 ("lu24", "Net land-use source 2024", "MtCO2e", "MAX(0,@ghg_inc-@ghg_exc)"),
 ("cbam24", "CBAM-goods process CO2 2024", "MtCO2", "@cbam_proc24"),
 ("cbam_g", "CBAM-goods output growth to 2030", "%", "@cbam_g"),
 ("crf_a_grid", "CRF at market WACC, grids", "x", "#wacc/(1-(1+#wacc)^-{life_grid})"),
 ("crf_a_vre", "CRF at market WACC, generation", "x", "#wacc/(1-(1+#wacc)^-{life_vre})"),
 ("crf_a_veh", "CRF at market WACC, vehicles", "x", "#wacc/(1-(1+#wacc)^-{veh_life})"),
 ("crf_a_bld", "CRF at market WACC, buildings", "x", "#wacc/(1-(1+#wacc)^-{life_bld})"),
 ("crf_a_ind", "CRF at market WACC, industry", "x", "#wacc/(1-(1+#wacc)^-{life_ind})"),
 ("crf_b_grid", "CRF at 6%, grids", "x", "{r_wb}/(1-(1+{r_wb})^-{life_grid})"),
 ("crf_b_vre", "CRF at 6%, generation", "x", "{r_wb}/(1-(1+{r_wb})^-{life_vre})"),
 ("crf_b_veh", "CRF at 6%, vehicles", "x", "{r_wb}/(1-(1+{r_wb})^-{veh_life})"),
 ("crf_b_bld", "CRF at 6%, buildings", "x", "{r_wb}/(1-(1+{r_wb})^-{life_bld})"),
 ("crf_b_ind", "CRF at 6%, industry", "x", "{r_wb}/(1-(1+{r_wb})^-{life_ind})"),
 # investment typology (who pays): pressure indices, then type shares by sector
 ("pms", "Power market structure (1 integrated, 2 single buyer, 3 wholesale)", "code", "@pms"),
 ("ty_u", "Utility pressure index: electricity tariff under-recovery", "0-1", "MIN(1,MAX(0,(1-@cr_ecy)/{ty_cr_gap}))"),
 ("ty_h", "Capital-cost pressure index: market WACC above 6%", "0-1", "MIN(1,MAX(0,(#wacc-{r_wb})/{ty_h_span}))"),
 ("ty_v", "Subsidy pressure index: explicit fossil fuel subsidies", "0-1", "MIN(1,MAX(0,@ffs_exp/{ty_v_span}))"),
 ("ty_m", "Top-up share of market end-use investment (P5 to P4)", "share",
  "MIN(1,#ty_h*{ty_kh}+#ty_v*{ty_kv}+CHOOSE(#inc,{ty_ci_lic},{ty_ci_lmic},0,0))"),
 ("g2o", "Generation, utility-built share before pressure", "share", "CHOOSE(#pms,{ty_g2_vi},{ty_g2_sb},{ty_g2_wm})"),
 ("g3o", "Generation, contracted share before pressure", "share", "CHOOSE(#pms,{ty_g3_vi},{ty_g3_sb},{ty_g3_wm})"),
 ("g5o", "Generation, merchant and distributed share before pressure", "share", "1-#g2o-#g3o"),
 ("ty_gen_p1", "Type share, generation: P1 pure public", "share", "#g2o*#ty_u*{ty_ku1}"),
 ("ty_gen_p2", "Type share, generation: P2 regulated utility", "share", "#g2o-#ty_gen_p1"),
 ("ty_gen_p3", "Type share, generation: P3 contracted revenue", "share", "#g3o*(1-#ty_u*{ty_ku4})"),
 ("ty_gen_p4", "Type share, generation: P4 viability gap", "share", "#g3o*#ty_u*{ty_ku4}+#g5o*MIN(1,#ty_h*{ty_kh})"),
 ("ty_gen_p5", "Type share, generation: P5 pure private", "share", "#g5o*(1-MIN(1,#ty_h*{ty_kh}))"),
 ("gr3", "Grids and storage, private contracted share", "share", "CHOOSE(#pms,{ty_r3_vi},{ty_r3_sb},{ty_r3_wm})"),
 ("ty_grid_p1", "Type share, grids: P1 pure public", "share",
  "(1-#gr3)*MIN(1,#ty_u*{ty_kur}+CHOOSE(#inc,{ty_acc_lic},{ty_acc_lmic},0,0))"),
 ("ty_grid_p2", "Type share, grids: P2 regulated utility", "share", "1-#gr3-#ty_grid_p1"),
 ("ty_grid_p3", "Type share, grids: P3 contracted revenue", "share", "#gr3"),
 ("ty_grid_p4", "Type share, grids: P4 viability gap", "share", "0"),
 ("ty_grid_p5", "Type share, grids: P5 pure private", "share", "0"),
 ("ty_trn_p1", "Type share, transport: P1 pure public", "share",
  "CHOOSE(#inc,{ty_t1_lic},{ty_t1_lmic},{ty_t1_umic},{ty_t1_hic})"),
 ("ty_trn_p2", "Type share, transport: P2 regulated utility", "share", "0"),
 ("ty_trn_p3", "Type share, transport: P3 contracted revenue", "share", "{ty_t3}"),
 ("ty_trn_p4", "Type share, transport: P4 viability gap", "share", "(1-#ty_trn_p1-#ty_trn_p3)*#ty_m"),
 ("ty_trn_p5", "Type share, transport: P5 pure private", "share", "(1-#ty_trn_p1-#ty_trn_p3)*(1-#ty_m)"),
 ("ty_bld_p1", "Type share, buildings: P1 pure public", "share",
  "CHOOSE(#inc,{ty_b1_lic},{ty_b1_lmic},{ty_b1_umic},{ty_b1_hic})"),
 ("ty_bld_p2", "Type share, buildings: P2 regulated utility", "share", "0"),
 ("ty_bld_p3", "Type share, buildings: P3 contracted revenue", "share", "0"),
 ("ty_bld_p4", "Type share, buildings: P4 viability gap", "share", "(1-#ty_bld_p1)*#ty_m"),
 ("ty_bld_p5", "Type share, buildings: P5 pure private", "share", "(1-#ty_bld_p1)*(1-#ty_m)"),
 ("ty_ih_p4", "Type share, industrial fuel switch: P4 viability gap (rest P5)", "share", "MIN(1,{ty_i4}+(1-{ty_i4})*#ty_m)"),
 ("ty_nc_p1", "Type share, non-CO2: P1 pure public", "share", "CHOOSE(#inc,{ty_n1_lic},{ty_n1_lmic},{ty_n1_umic},{ty_n1_hic})"),
 ("ty_nc_p5", "Type share, non-CO2: P5 pure private (rest P4)", "share", "{ty_n5}*(1-#ty_m)"),
 ("ty_lu_p1", "Type share, land use: P1 pure public (rest P4)", "share",
  "CHOOSE(#inc,{ty_l1_lic},{ty_l1_lmic},{ty_l1_umic},{ty_l1_hic})"),
]
for _k in range(1, 6):   # unmodelled clean investment: power and end-use mix (IEA WEI 2024)
    SCALARS.append((f"ty_cal_p{_k}", f"Type share, unmodelled clean investment: P{_k}", "share",
                    f"{{ty_cal_pow}}*#ty_gen_p{_k}+(1-{{ty_cal_pow}})*#ty_bld_p{_k}"))

# -----------------------------------------------------------------------------
# ROWS: (section, code, fuel, sector, description, unit, base template, projection template)
# base = 2024 column (N); proj = 2025-2050. History 2021-2023 shown for data rows.
# -----------------------------------------------------------------------------
ROWS = []
def R(sec, code, fuel, sector, desc, unit, base, proj, hist=None):
    ROWS.append(dict(sec=sec, code=code, fuel=fuel, sector=sector, desc=desc, unit=unit, base=base, proj=proj, hist=hist))

S2 = "2. Macro"
R(S2, "mit.gdp.gr", "", "", "Real GDP growth", "%", "0",
  "IF(YEAR=2025,#g25,IF(YEAR=2026,#g26,IF(YEAR<{g_conv_year},#g26+(#glr-#g26)*(YEAR-2026)/({g_conv_year}-2026),#glr)))")
R(S2, "mit.gdp.lvl", "", "", "Real GDP (2024 prices)", "USD bn", "#gdp24", "[mit.gdp.lvl@-1]*(1+[mit.gdp.gr]/100)")

S3 = "3. Activity (baseline end use, CPAT equation, prices fixed real)"
for f, s in ACT:
    base = f"@con_{f}_{s}" if s != "rod" else f"@con_{f}_{s}/(1-#sig0)"
    R(S3, f"mit.act.{f}.{s}", f, s, "Energy service demand" + (" (ICE-equivalent fuel)" if s == "rod" else ""), "PJ",
      base, f"[mit.act.{f}.{s}@-1]*(1+[mit.gdp.gr]/100)^@inc_{f}_{s}/(1+@aei_{f}_{s})",
      hist=[f"@con_{f}_{s}_{y}" for y in HIST])

S4 = "4. Road transport"
R(S4, "mit.veh.stock", "", "rod", "Vehicle stock (car-equivalent)", "million",
  "([mit.act.gso.rod]+[mit.act.die.rod])/{veh_gj}", "([mit.act.gso.rod]+[mit.act.die.rod])/{veh_gj}")
R(S4, "mit.veh.sales", "", "rod", "Vehicle sales", "million",
  "[mit.veh.stock]/{veh_life}", "[mit.veh.stock]/{veh_life}+MAX(0,[mit.veh.stock]-[mit.veh.stock@-1])")
R(S4, "mit.ev.prem", "", "rod", "EV premium incl. nothing for charging", "USD/vehicle",
  "{ev_prem24}", "MAX(0,{ev_prem24}*(1-(YEAR-2024)/({ev_parity}-2024)))")
for sc in (1, 2):
    sh = ("IF(YEAR>=2035,#evb35,#ev0+(#evb35-#ev0)*(YEAR-2024)/11)" if sc == 1
          else "MIN(1,#ev0+(1-#ev0)*(YEAR-2024)/(#ev_T-2024))")
    R(S4, f"mit.ev.shs.{sc}", "ecy", "rod", "EV share of sales", "share", "#ev0", sh)
    R(S4, f"mit.ev.stk.{sc}", "ecy", "rod", "EV share of stock", "share", "#sig0",
      f"[mit.ev.stk.{sc}@-1]*(1-1/{{veh_life}})+[mit.ev.shs.{sc}]/{{veh_life}}")
    for f in ("gso", "die"):
        R(S4, f"mit.con.{f}.rod.{sc}", f, "rod", "Fossil fuel use", "PJ",
          f"[mit.act.{f}.rod]*(1-[mit.ev.stk.{sc}])", f"[mit.act.{f}.rod]*(1-[mit.ev.stk.{sc}])")
    R(S4, f"mit.con.ecy.ev.{sc}", "ecy", "rod", "EV electricity", "PJ",
      f"([mit.act.gso.rod]+[mit.act.die.rod])*[mit.ev.stk.{sc}]*{{ev_ratio}}",
      f"([mit.act.gso.rod]+[mit.act.die.rod])*[mit.ev.stk.{sc}]*{{ev_ratio}}")
    R(S4, f"mit.inv.trn.{sc}", "", "rod", "Investment: EV premium and chargers", "USD bn", "0",
      f"[mit.veh.sales]*[mit.ev.shs.{sc}]*([mit.ev.prem]+{{charger}})*#pf/1000")
    R(S4, f"mit.om.trn.{sc}", "", "rod", "EV maintenance saving (negative cost)", "USD bn",
      f"-[mit.ev.stk.{sc}]*[mit.veh.stock]*{{veh_km}}*{{ev_maint}}/1000",
      f"-[mit.ev.stk.{sc}]*[mit.veh.stock]*{{veh_km}}*{{ev_maint}}/1000")

S5 = "5. Buildings"
R(S5, "mit.bld.h.2", "", "res", "Share of fossil heat replaced (net zero)", "share", "0",
  "MIN(1,(YEAR-2024)/(#bld_T-2024))")
BF = ("coa", "nga", "lpg", "ker")
for sc in (1, 2):
    for f in BF:
        R(S5, f"mit.con.{f}.res.{sc}", f, "res", "Fossil fuel use", "PJ", f"[mit.act.{f}.res]",
          f"[mit.act.{f}.res]" if sc == 1 else f"[mit.act.{f}.res]*(1-[mit.bld.h.2])")
R(S5, "mit.bld.disp.2", "", "res", "Fossil heat fuel displaced", "PJ", "0",
  "+".join(f"[mit.act.{f}.res]" for f in BF).join(["(", ")"]) + "*[mit.bld.h.2]")
R(S5, "mit.con.ecy.hp.2", "ecy", "res", "Heat pump electricity", "PJ", "0", "[mit.bld.disp.2]*{hp_ratio}")
R(S5, "mit.inv.bld.2", "", "res", "Investment: heat pumps and efficiency", "USD bn", "0",
  "{k_bld}*MAX(0,[mit.bld.disp.2]-[mit.bld.disp.2@-1])*#pf/1000")

S6 = "6. Industry"
R(S6, "mit.ind.a.2", "", "ind", "Share of industrial fossil use abated (net zero)", "share", "0",
  "#ind_a*(YEAR-2024)/26")
IF_ = ("coa", "nga", "oil")
for sc in (1, 2):
    for f in IF_:
        R(S6, f"mit.con.{f}.ind.{sc}", f, "ind", "Fossil fuel use", "PJ", f"[mit.act.{f}.ind]",
          f"[mit.act.{f}.ind]" if sc == 1 else f"[mit.act.{f}.ind]*(1-[mit.ind.a.2])")
R(S6, "mit.ind.abt.2", "", "ind", "Industrial fossil use abated", "PJ", "0",
  "([mit.act.coa.ind]+[mit.act.nga.ind]+[mit.act.oil.ind])*[mit.ind.a.2]")
R(S6, "mit.con.ecy.ix.2", "ecy", "ind", "Electricity for electrified industry", "PJ", "0",
  "[mit.ind.abt.2]*{ind_elec_sh}*{ind_ratio}")
R(S6, "mit.cbam.co2.1", "", "ind", "CBAM-goods process CO2 (CBAM tool output)", "MtCO2", "#cbam24",
  "#cbam24*(1+#cbam_g/100)^MIN(YEAR-2024,6)")
R(S6, "mit.cbam.abt.2", "", "ind", "CBAM-goods process CO2 abated (net zero)", "MtCO2", "0",
  "[mit.cbam.co2.1]*{cbam_a_2050}*(YEAR-2024)/26")
R(S6, "mit.inv.ixh.2", "", "ind", "Investment: industrial fuel switch", "USD bn", "0",
  "{k_ind}*MAX(0,[mit.ind.abt.2]-[mit.ind.abt.2@-1])*#pf/1000")
R(S6, "mit.inv.ccs.2", "", "ind", "Investment: CCS on CBAM-goods process CO2", "USD bn", "0",
  "{k_ccs}*MAX(0,[mit.cbam.abt.2]-[mit.cbam.abt.2@-1])*#pf/1000")
R(S6, "mit.inv.ind.2", "", "ind", "Investment: industrial fuel switch and CCS", "USD bn", "0",
  "[mit.inv.ixh.2]+[mit.inv.ccs.2]")

S7 = "7. Power"
R(S7, "mit.cap.sol.cost", "", "pow", "Solar capital cost", "USD/kW", "{capex_sol24}", "{capex_sol24}*(1-{lr_sol})^(YEAR-2024)")
R(S7, "mit.cap.wnd.cost", "", "pow", "Wind capital cost", "USD/kW", "{capex_wnd24}", "{capex_wnd24}*(1-{lr_wnd})^(YEAR-2024)")
R(S7, "mit.cap.st.cost", "", "pow", "Storage capital cost", "USD/kWh", "{capex_st24}", "{capex_st24}*(1-{lr_st})^(YEAR-2024)")
for sc in (1, 2):
    extra = "" if sc == 1 else "+[mit.con.ecy.hp.2]+[mit.con.ecy.ix.2]"
    R(S7, f"mit.dem.ecy.{sc}", "ecy", "pow", "End-use electricity demand", "PJ",
      f"[mit.act.ecy.res]+[mit.act.ecy.ind]+[mit.con.ecy.ev.{sc}]",
      f"[mit.act.ecy.res]+[mit.act.ecy.ind]+[mit.con.ecy.ev.{sc}]{extra}")
    R(S7, f"mit.gen.tot.{sc}", "ecy", "pow", "Generation", "TWh", f"#gcr*[mit.dem.ecy.{sc}]/3.6",
      f"#gcr*[mit.dem.ecy.{sc}]/3.6")
    fs = ("MAX(0,[mit.gen.fsh.1@-1]-#ftrend)" if sc == 1
          else "MIN([mit.gen.fsh.1],#f0*MAX(0,1-(YEAR-2024)/(#pow_T-2024)))")
    R(S7, f"mit.gen.fsh.{sc}", "", "pow", "Fossil share of generation", "share", "#f0", fs)
    R(S7, f"mit.gen.fos.{sc}", "", "pow", "Fossil generation", "TWh",
      f"MIN([mit.gen.tot.{sc}]*[mit.gen.fsh.{sc}],MAX(0,[mit.gen.tot.{sc}]-#cfirm))",
      f"MIN([mit.gen.tot.{sc}]*[mit.gen.fsh.{sc}],MAX(0,[mit.gen.tot.{sc}]-#cfirm))")
    R(S7, f"mit.gen.vre.{sc}", "", "pow", "VRE generation needed", "TWh",
      f"MAX(0,[mit.gen.tot.{sc}]-#cfirm-[mit.gen.fos.{sc}])", f"MAX(0,[mit.gen.tot.{sc}]-#cfirm-[mit.gen.fos.{sc}])")
    R(S7, f"mit.gen.sol.{sc}", "", "pow", "Solar generation", "TWh", "@gen_sol",
      f"[mit.gen.sol.{sc}@-1]+{{sh_sol}}*MAX(0,[mit.gen.vre.{sc}]-[mit.gen.vre.{sc}@-1])")
    R(S7, f"mit.gen.wnd.{sc}", "", "pow", "Wind generation", "TWh", "@gen_wnd",
      f"[mit.gen.wnd.{sc}@-1]+(1-{{sh_sol}})*MAX(0,[mit.gen.vre.{sc}]-[mit.gen.vre.{sc}@-1])")
    for t, cf in (("sol", "cf_sol"), ("wnd", "cf_wnd")):
        R(S7, f"mit.cap.{t}.{sc}", "", "pow", f"{'Solar' if t == 'sol' else 'Wind'} capacity", "GW",
          f"[mit.gen.{t}.{sc}]/(#{cf}*8.76)", f"[mit.gen.{t}.{sc}]/(#{cf}*8.76)")
        R(S7, f"mit.add.{t}.{sc}", "", "pow", f"{'Solar' if t == 'sol' else 'Wind'} additions incl. replacement", "GW", "0",
          f"MAX(0,[mit.cap.{t}.{sc}]-[mit.cap.{t}.{sc}@-1])+[mit.cap.{t}.{sc}@-1]/{{life_vre}}")
    R(S7, f"mit.cap.fos.need.{sc}", "", "pow", "Fossil capacity needed", "GW",
      f"[mit.gen.fos.{sc}]/(#cf_fos*8.76)", f"[mit.gen.fos.{sc}]/(#cf_fos*8.76)")
    R(S7, f"mit.add.fos.{sc}", "", "pow", "Fossil additions (gas)", "GW", "0",
      f"MAX(0,[mit.cap.fos.need.{sc}]-[mit.cap.fos.stk.{sc}@-1]*(1-1/{{life_fos}}))")
    R(S7, f"mit.cap.fos.stk.{sc}", "", "pow", "Fossil capacity in service", "GW", f"[mit.cap.fos.need.{sc}]",
      f"[mit.cap.fos.stk.{sc}@-1]*(1-1/{{life_fos}})+[mit.add.fos.{sc}]")
    R(S7, f"mit.inv.gen.{sc}", "", "pow", "Investment: generation", "USD bn", "0",
      f"([mit.add.sol.{sc}]*[mit.cap.sol.cost]+[mit.add.wnd.{sc}]*[mit.cap.wnd.cost]+[mit.add.fos.{sc}]*{{capex_gas}})*#pf/1000")
    R(S7, f"mit.gen.vsh.{sc}", "", "pow", "VRE share of generation", "share",
      f"IF([mit.gen.tot.{sc}]>0,([mit.gen.sol.{sc}]+[mit.gen.wnd.{sc}])/[mit.gen.tot.{sc}],0)",
      f"IF([mit.gen.tot.{sc}]>0,([mit.gen.sol.{sc}]+[mit.gen.wnd.{sc}])/[mit.gen.tot.{sc}],0)")
    R(S7, f"mit.inv.grid.{sc}", "", "pow", "Investment: grids and storage", "USD bn", "0",
      f"{{grid_ratio}}*[mit.inv.gen.{sc}]+([mit.add.sol.{sc}]+[mit.add.wnd.{sc}])*{{st_h}}"
      f"*MIN(1,MAX(0,([mit.gen.vsh.{sc}]-0.2)/0.6))*[mit.cap.st.cost]*#pf/1000")
    for f, hr in (("coa", "#hr_coa"), ("nga", "#hr_nga"), ("oil", "{hr_oil}")):
        R(S7, f"mit.con.{f}.pow.{sc}", f, "pow", "Power fuel use", "PJ",
          f"[mit.gen.fos.{sc}]*#sh_{f}*{hr}", f"[mit.gen.fos.{sc}]*#sh_{f}*{hr}")
    R(S7, f"mit.om.pow.{sc}", "", "pow", "Power O&M", "USD bn",
      f"{{om_vre}}*([mit.cap.sol.{sc}]*{{capex_sol24}}+[mit.cap.wnd.{sc}]*{{capex_wnd24}})/1000+[mit.gen.fos.{sc}]*{{om_fos}}",
      f"{{om_vre}}*([mit.cap.sol.{sc}]*{{capex_sol24}}+[mit.cap.wnd.{sc}]*{{capex_wnd24}})/1000+[mit.gen.fos.{sc}]*{{om_fos}}")

S8 = "8. Non-CO2, process and land use"
R(S8, "mit.nc.abt.2", "", "", "Non-CO2 and process GHG abated (net zero)", "MtCO2e", "0",
  "#nc24*IF(YEAR<=2030,{nc_2030}*(YEAR-2024)/6,{nc_2030}+({nc_2050}-{nc_2030})*(YEAR-2030)/20)")
R(S8, "mit.lu.abt.2", "", "", "Net land-use emissions abated (net zero)", "MtCO2e", "0",
  "#lu24*MIN(1,(YEAR-2024)/({lu_T}-2024))")
R(S8, "mit.spd.nc.2", "", "", "Spending: non-CO2 and process abatement", "USD bn", "0", "[mit.nc.abt.2]*{cpt_nc}/1000")
R(S8, "mit.spd.lu.2", "", "", "Spending: land-use abatement", "USD bn", "0", "[mit.lu.abt.2]*{cpt_afolu}/1000")
R(S8, "mit.inv.nco.2", "", "", "Spending: non-CO2 and land-use abatement", "USD bn", "0", "[mit.spd.nc.2]+[mit.spd.lu.2]")

S9 = "9. Emissions, fuel bills and air pollution"
EFK = {"gso": "ef_gso", "die": "ef_die", "coa": "ef_coa", "nga": "ef_nga", "lpg": "ef_lpg", "ker": "ef_ker", "oil": "ef_oil"}
def fuel_terms(sc):
    return ([("gso", "rod"), ("die", "rod")] + [(f, "res") for f in BF] + [(f, "ind") for f in IF_]
            + [("coa", "pow"), ("nga", "pow"), ("oil", "pow")])
for sc in (1, 2):
    co2 = "+".join(f"[mit.con.{f}.{s}.{sc}]*{{{EFK[f]}}}" for f, s in fuel_terms(sc))
    R(S9, f"mit.co2.en.{sc}", "", "", "Energy CO2", "MtCO2", co2, co2)
    if sc == 1:
        R(S9, "mit.co2.res.1", "", "", "Other energy CO2 in CPAT not in the rows above (held flat)", "MtCO2",
          "MAX(0,@co2_24-[mit.co2.en.1])", "[mit.co2.res.1@-1]")
    else:
        R(S9, "mit.co2.res.2", "", "", "Other energy CO2, net zero (abated with industry)", "MtCO2",
          "[mit.co2.res.1]", "[mit.co2.res.1]*(1-[mit.ind.a.2])")
        R(S9, "mit.inv.oth.2", "", "", "Spending: other energy CO2 abated (aviation, shipping, other)", "USD bn", "0",
          "([mit.co2.res.1]-[mit.co2.res.2])*{cpt_res}/1000")
    ghg = (f"[mit.co2.en.{sc}]+[mit.co2.res.{sc}]+[mit.cbam.co2.1]+#nc24+#lu24" if sc == 1 else
           f"[mit.co2.en.{sc}]+[mit.co2.res.{sc}]+[mit.cbam.co2.1]-[mit.cbam.abt.2]+#nc24+#lu24-[mit.nc.abt.2]-[mit.lu.abt.2]")
    R(S9, f"mit.ghg.tot.{sc}", "", "", "Total GHG incl. land use", "MtCO2e", ghg, ghg)
    fb = "(" + "+".join(f"[mit.con.{f}.{s}.{sc}]*@sp_{f}_{s}" for f, s in fuel_terms(sc)) + ")/1000"
    R(S9, f"mit.fuel.bill.{sc}", "", "", "Fossil fuel bill at 2024 supply costs", "USD bn", fb, fb)
    ap = "(" + "+".join(f"[mit.con.{f}.{s}.{sc}]*@ap_{f}_{s}" for f, s in fuel_terms(sc)) + ")/1000"
    R(S9, f"mit.airpol.{sc}", "", "", "Air pollution cost", "USD bn", ap, ap)

S10 = "10. Investment and incremental investment"
for sc in (1, 2):
    R(S10, f"mit.inv.cal.{sc}", "", "", "Clean investment not modelled bottom-up (IEA 2024 calibration; both scenarios)",
      "USD bn", "0", "@cal24*[mit.gdp.lvl]/#gdp24")
SECT = [("gen", "Power generation", "vre"), ("grid", "Grids and storage", "grid"), ("trn", "Transport", "veh"),
        ("bld", "Buildings", "bld"), ("ind", "Industry", "ind"), ("nco", "Non-CO2 and land use", None),
        ("oth", "Other energy CO2", None), ("cal", "Unmodelled clean investment (calibration)", "vre")]
HAS = lambda c: any(r["code"] == c for r in ROWS)
for k, nm, life in SECT:
    t = "-".join(x for x in ([f"[mit.inv.{k}.2]"] if HAS(f"mit.inv.{k}.2") else ["0"]) +
                 ([f"[mit.inv.{k}.1]"] if HAS(f"mit.inv.{k}.1") else []))
    R(S10, f"mit.dinv.{k}", "", "", f"Incremental investment: {nm.lower()}", "USD bn", "0", t)
R(S10, "mit.inv.tot.1", "", "", "Total investment, baseline", "USD bn", "0",
  "+".join(f"[mit.inv.{k}.1]" for k, *_ in SECT if any(r["code"] == f"mit.inv.{k}.1" for r in ROWS)))
R(S10, "mit.inv.tot.2", "", "", "Total investment, net zero", "USD bn", "0",
  "+".join(f"[mit.inv.{k}.2]" for k, *_ in SECT if HAS(f"mit.inv.{k}.2")))
R(S10, "mit.dinv.tot", "", "", "Incremental investment, total", "USD bn", "0", "[mit.inv.tot.2]-[mit.inv.tot.1]")

S11 = "11. Incremental cost by lens"
for k, nm, life in SECT:
    if life is None: continue
    lp = {"vre": "life_vre", "grid": "life_grid", "veh": "veh_life", "bld": "life_bld", "ind": "life_ind"}[life]
    R(S11, f"mit.alive.{k}", "", "", f"Incremental capital in service: {nm.lower()}", "USD bn", "0",
      f"[mit.alive.{k}@-1]+[mit.dinv.{k}]-LAG(mit.dinv.{k},{{{lp}}})")
for lens in ("a", "b"):
    terms = "+".join(f"[mit.alive.{k}]*#crf_{lens}_{l}" for k, nm, l in SECT if l)
    R(S11, f"mit.kchg.{lens}", "", "", f"Capital charge on incremental capital, lens ({lens})", "USD bn", "0",
      f"{terms}+[mit.dinv.nco]+[mit.dinv.oth]")
R(S11, "mit.dom", "", "", "Change in O&M (net zero less baseline)", "USD bn", "0",
  "[mit.om.pow.2]+[mit.om.trn.2]-[mit.om.pow.1]-[mit.om.trn.1]")
R(S11, "mit.dfuel", "", "", "Fuel savings (baseline less net zero)", "USD bn", "0", "[mit.fuel.bill.1]-[mit.fuel.bill.2]")
R(S11, "mit.dap", "", "", "Air pollution avoided", "USD bn", "0", "[mit.airpol.1]-[mit.airpol.2]")
R(S11, "mit.ic.a", "", "", "INCREMENTAL COST (a) market WACC", "USD bn", "0", "[mit.kchg.a]+[mit.dom]-[mit.dfuel]")
R(S11, "mit.ic.b", "", "", "INCREMENTAL COST (b) World Bank 6%", "USD bn", "0", "[mit.kchg.b]+[mit.dom]-[mit.dfuel]")
R(S11, "mit.ic.c", "", "", "INCREMENTAL COST (c) 6% + local co-benefits", "USD bn", "0", "[mit.ic.b]-[mit.dap]")
R(S11, "mit.ghg.abt", "", "", "GHG abated (baseline less net zero)", "MtCO2e", "0", "[mit.ghg.tot.1]-[mit.ghg.tot.2]")
R(S11, "mit.df.a", "", "", "Discount factor, market WACC", "x", "1", "(1+#wacc)^-(YEAR-2024)")
R(S11, "mit.df.b", "", "", "Discount factor, 6%", "x", "1", "(1+{r_wb})^-(YEAR-2024)")

S12 = "12. Investment by type: who pays (P1 pure public ... P5 pure private)"
TYPES = [(1, "P1 pure public"), (2, "P2 regulated utility"), (3, "P3 contracted revenue"),
         (4, "P4 viability gap"), (5, "P5 pure private")]
def typ_terms(k, sc):
    """Investment rows of scenario sc times their type-k share."""
    t = [f"[mit.inv.{x}.{sc}]*#ty_{x}_p{k}" for x in ("gen", "grid", "trn", "bld", "cal") if HAS(f"mit.inv.{x}.{sc}")]
    if sc == 2:
        if k == 4: t += ["[mit.inv.ixh.2]*#ty_ih_p4", "[mit.inv.ccs.2]", "[mit.spd.nc.2]*(1-#ty_nc_p1-#ty_nc_p5)",
                         "[mit.spd.lu.2]*(1-#ty_lu_p1)", "[mit.inv.oth.2]"]
        if k == 5: t += ["[mit.inv.ixh.2]*(1-#ty_ih_p4)", "[mit.spd.nc.2]*#ty_nc_p5"]
        if k == 1: t += ["[mit.spd.nc.2]*#ty_nc_p1", "[mit.spd.lu.2]*#ty_lu_p1"]
    return "+".join(t) if t else "0"
for sc in (1, 2):
    for k, nm in TYPES:
        R(S12, f"mit.typ.p{k}.{sc}", "", "", f"Investment, {nm}, {'baseline' if sc == 1 else 'net zero'}", "USD bn",
          "0", typ_terms(k, sc))
for k, nm in TYPES:
    R(S12, f"mit.dtyp.p{k}", "", "", f"Incremental investment, {nm}", "USD bn", "0", f"[mit.typ.p{k}.2]-[mit.typ.p{k}.1]")
R(S12, "mit.typ.bud.2", "", "", "Direct budget need, net zero: P1 + public gap share of P4", "USD bn", "0",
  "[mit.typ.p1.2]+{ty_p4_gap}*[mit.typ.p4.2]")
R(S12, "mit.dtyp.bud", "", "", "Direct budget need, incremental: P1 + public gap share of P4", "USD bn", "0",
  "[mit.dtyp.p1]+{ty_p4_gap}*[mit.dtyp.p4]")
R(S12, "mit.typ.cnt.2", "", "", "Revenue resting on a public tariff or buyer, net zero (P2 + P3)", "USD bn", "0",
  "[mit.typ.p2.2]+[mit.typ.p3.2]")
R(S12, "mit.typ.prv.2", "", "", "Private capital, net zero: P3 + P5 + private part of P4", "USD bn", "0",
  "[mit.typ.p3.2]+[mit.typ.p5.2]+(1-{ty_p4_gap})*[mit.typ.p4.2]")

ROWIDX = {r["code"]: i for i, r in enumerate(ROWS)}
assert len(ROWIDX) == len(ROWS), "duplicate codes"
for _i, _r in enumerate(ROWS):   # same-year references must point to rows evaluated earlier
    for _t in (_r["base"], _r["proj"]):
        for _c in re.findall(r"\[([\w.]+)\]", _t):
            assert ROWIDX[_c] < _i, f"{_r['code']} reads {_c} before it is computed"

# -----------------------------------------------------------------------------
# PYTHON EVALUATION
# -----------------------------------------------------------------------------
def MAX(*a): return max(a)
def MIN(*a): return min(a)
def IF(c, a, b): return a if c else b
def CHOOSE(i, *a): return a[int(round(i)) - 1]
ABS = abs

def to_py(t):
    s = t
    s = re.sub(r"LAG\(([\w.]+),\s*([^)]+)\)", lambda m: f'LAGF("{m.group(1)}",i,{m.group(2)})', s)
    s = re.sub(r"\[([\w.]+)@-1\]", r'V["\1"][i-1]', s)
    s = re.sub(r"\[([\w.]+)@N\]", r'V["\1"][0]', s)
    s = re.sub(r"\[([\w.]+)\]", r'V["\1"][i]', s)
    s = re.sub(r"\{(\w+)\}", r'P["\1"]', s)
    s = re.sub(r"#(\w+)", r'S["\1"]', s)
    s = re.sub(r"@(\w+)", r'D["\1"]', s)
    s = s.replace("YEAR", "Y").replace("^", "**")
    s = re.sub(r"(?<![<>!=])=(?!=)", "==", s)
    return s

COMPILED_S = [(k, compile(to_py(t), k, "eval")) for k, _, _, t in SCALARS]
COMPILED_R = [(r["code"], compile(to_py(r["base"]), r["code"] + ".b", "eval"),
               compile(to_py(r["proj"]), r["code"] + ".p", "eval")) for r in ROWS]

np.seterr(all="ignore")   # Python IF() is eager: a branch Excel never evaluates may divide by zero

def run_country(d, params=None):
    PP = dict(P); PP.update(params or {})
    d = {k: (np.float64(v) if isinstance(v, (int, float, np.floating)) else v) for k, v in d.items()}
    S = {}
    env = dict(MAX=MAX, MIN=MIN, IF=IF, CHOOSE=CHOOSE, ABS=ABS, P=PP, D=d, S=S)
    for k, c in COMPILED_S:
        S[k] = eval(c, env)
    n = len(YEARS)
    V = {code: [0.0] * n for code, *_ in COMPILED_R}
    def LAGF(code, i, lag):
        lag = int(round(lag))
        return V[code][i - lag] if i - lag >= 1 else 0.0
    env.update(V=V, LAGF=LAGF)
    for i, y in enumerate(YEARS):
        env["i"], env["Y"] = i, y
        for code, cb, cp in COMPILED_R:
            V[code][i] = float(eval(cb if i == 0 else cp, env))
    return S, V

def summarise(iso, d, S, V):
    yi = {y: i for i, y in enumerate(YEARS)}
    out = dict(iso=iso, name=d["name"], grp=["AE", "China", "EMDC"][d["grp"] - 1], wacc=S["wacc"])
    for y in (2030, 2040, 2050):
        i = yi[y]
        out[f"inv1_{y}"] = V["mit.inv.tot.1"][i]; out[f"inv2_{y}"] = V["mit.inv.tot.2"][i]
        out[f"dinv_{y}"] = V["mit.dinv.tot"][i]
        for L in "abc": out[f"ic{L}_{y}"] = V[f"mit.ic.{L}"][i]
        out[f"ghg1_{y}"] = V["mit.ghg.tot.1"][i]; out[f"ghg2_{y}"] = V["mit.ghg.tot.2"][i]
        out[f"gdp_{y}"] = V["mit.gdp.lvl"][i]
    rng = range(yi[2025], yi[2050] + 1)
    out["cum_inv2"] = sum(V["mit.inv.tot.2"][i] for i in rng)
    out["cum_dinv"] = sum(V["mit.dinv.tot"][i] for i in rng)
    out["pv_ic_a"] = sum(V["mit.ic.a"][i] * V["mit.df.a"][i] for i in rng)
    out["pv_ic_b"] = sum(V["mit.ic.b"][i] * V["mit.df.b"][i] for i in rng)
    out["pv_ic_c"] = sum(V["mit.ic.c"][i] * V["mit.df.b"][i] for i in rng)
    out["cum_abt"] = sum(V["mit.ghg.abt"][i] for i in rng)
    out["co2_model_2024"] = V["mit.co2.en.1"][0]; out["co2_cpat_2024"] = d["co2_24"]
    for k in range(1, 6):
        for y in (2030, 2040):
            out[f"typ{k}_{y}"] = V[f"mit.typ.p{k}.2"][yi[y]]
        out[f"dtyp{k}_2030"] = V[f"mit.dtyp.p{k}"][yi[2030]]
        out[f"typ{k}_cum"] = sum(V[f"mit.typ.p{k}.2"][i] for i in rng)
    out["bud_2030"] = V["mit.typ.bud.2"][yi[2030]]; out["dbud_2030"] = V["mit.dtyp.bud"][yi[2030]]
    out["prv_2030"] = V["mit.typ.prv.2"][yi[2030]]
    out["pms"] = d["pms"]; out["cr_ecy"] = d["cr_ecy"]
    for x in ("ty_u", "ty_h", "ty_v", "ty_m"): out[x] = S[x]
    for k, *_ in SECT:
        out[f"dinv_{k}_2030"] = V[f"mit.dinv.{k}"][yi[2030]]
        out[f"inv2_{k}_2030"] = V[f"mit.inv.{k}.2"][yi[2030]] if f"mit.inv.{k}.2" in V else 0.0
    return out

def calibrate(D):
    """Two passes: the baseline's 2025 investment by group is compared with IEA 2024 clean-energy
    investment; the shortfall is clean investment the bottom-up rows do not model (hydro, nuclear,
    efficiency beyond the rows, other). It is allocated to countries by 2024 GDP share (cal24), grows
    with GDP and is added to BOTH scenarios, so it calibrates levels but not incremental investment."""
    tgt = {1: P["iea24_ae"], 2: P["iea24_chn"], 3: P["iea24_emdc"]}
    base, gdp = {1: 0.0, 2: 0.0, 3: 0.0}, {1: 0.0, 2: 0.0, 3: 0.0}
    for iso, d in D.items():
        d["cal24"] = 0.0
        S, V = run_country(d)
        base[d["grp"]] += V["mit.inv.tot.1"][1]; gdp[d["grp"]] += d["gdp24"]
    gap = {g: max(0.0, tgt[g] - base[g]) for g in tgt}
    for iso, d in D.items():
        d["cal24"] = gap[d["grp"]] * d["gdp24"] / gdp[d["grp"]]
    return {g: (base[g], tgt[g], gap[g]) for g in tgt}

def load_calibrated():
    D = load_data(); cal = calibrate(D); return D, cal

def main():
    os.makedirs(RES, exist_ok=True)
    D, cal = load_calibrated()
    print("baseline calibration (model 2025, IEA 2024, gap):", {k: tuple(round(x) for x in v) for k, v in cal.items()})
    long, summ = [], []
    for iso, d in D.items():
        S, V = run_country(d)
        for code in V:
            long.append([iso, f"{iso.lower()}.{code}", ROWS[ROWIDX[code]]["unit"]] + V[code])
        summ.append(summarise(iso, d, S, V))
    cols = ["iso", "code", "unit"] + YEARS
    pd.DataFrame(long, columns=cols).to_csv(os.path.join(RES, "nmim_results_v0_2.csv"), index=False, float_format="%.10g")
    sm = pd.DataFrame(summ)
    sm.to_csv(os.path.join(RES, "nmim_summary_v0_2.csv"), index=False, float_format="%.10g")
    g = sm.groupby("grp")[[c for c in sm.columns if c.endswith("_2030") and not c.startswith(("gdp", "dinv_", "inv2_"))]
                          + ["cum_inv2", "cum_dinv", "pv_ic_a", "pv_ic_b", "pv_ic_c", "cum_abt"]].sum()
    g.loc["World"] = g.sum()
    pd.set_option("display.width", 220)
    print(g.round(0).T)
    print(sm[[f"inv2_{k}_2030" for k, *_ in SECT]].sum().round(0))
    t = sm.groupby("grp")[[f"typ{k}_2030" for k in range(1, 6)] + ["bud_2030", "prv_2030"]].sum()
    t.loc["World"] = t.sum()
    print(t.round(0)); print((t[[f"typ{k}_2030" for k in range(1, 6)]].div(t[[f"typ{k}_2030" for k in range(1, 6)]].sum(axis=1), axis=0) * 100).round(1))
    return sm

if __name__ == "__main__":
    main()
