#!/usr/bin/env python3
"""Build the Integrated MF-C (macro-fiscal climate) tool.

The tool is the Global Fiscal Indicators dashboard with extra modules spliced in as blocks.
Each module turns a repository model's results into rows of the dashboard's long format
(iso3, country, region, income, is_ccdr, indicator, value, year, source) and brings its own
indicator and source rows (same columns as the dashboard's config). This script writes the
merged panel and config to build/output/merged/ and runs the dashboard's own builder on them,
so the integrated workbook reads, scores and looks exactly like the dashboard.

Inputs : TaxMacroFiscalDashboard/build/{output/global_data_long.csv, config/}
         per module, see MODULES (now: InvestmentNeeds NMIM v0.2 summary)
Output : IntegratedMFC/Integrated_MFC_Tool_v<version>.xlsx (version in config/tool.json)
         build/output/<module>_long.csv (the module's rows, committed)

To add a module: write an extract function (given the dashboard's long panel) returning long rows, add
config/<name>_indicators.csv and config/<name>_sources.csv, and register it in MODULES.
"""
import os, sys, shutil, subprocess
import numpy as np, pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
DASH = os.path.join(ROOT, "TaxMacroFiscalDashboard", "build")
INV = os.path.join(ROOT, "InvestmentNeeds")
CONFIG = os.path.join(HERE, "config")
OUT = os.path.join(HERE, "output")
MERGED = os.path.join(OUT, "merged")
SHIP = os.path.abspath(os.path.join(HERE, ".."))


# ---------------------------------------------------------------- module: investment needs
NMIM = os.path.join(INV, "national_model", "results", "nmim_summary_v0_2.csv")
SRC = "nmim_summary_v0_2.csv (InvestmentNeeds NMIM v0.2, model output)"
SRC_PIMA = "nmim_summary_v0_2.csv + dashboard PEFA PI-11 (InvestmentNeeds NMIM v0.2, model output)"
# Public investment efficiency gap, as in the v1.1 workbook (IMF 2015, Making Public
# Investment More Efficient): AE 13%, low income 40%, others 27%; where the dashboard has a
# PEFA PI-11 score (1-4) the gap runs linearly from 40% at 1 to 13% at 4.
GAP_AE, GAP_EM, GAP_LIC = 0.13, 0.27, 0.40


def extract_investment_needs(long):
    nm = pd.read_csv(NMIM).set_index("iso")
    g30, g50 = nm.gdp_2030, nm.gdp_2050
    # NMIM's budget need is P1 + a public share of P4 (ty_p4_gap); recover it from the
    # results rather than restating it, and check it is one number for every country.
    p4 = nm.typ4_2030.where(nm.typ4_2030 > 1e-9)
    gap_share = ((nm.bud_2030 - nm.typ1_2030) / p4).dropna()
    if gap_share.max() - gap_share.min() > 1e-6:
        sys.exit("NMIM budget need is not P1 + a fixed share of P4; check ty_p4_gap in nmim_v0_2.py")
    k4 = float(gap_share.median())
    out = []

    def emit(series, code, src=SRC, year=2030):
        for iso, v in series.items():
            if pd.notna(v) and np.isfinite(float(v)):
                out.append((iso, code, float(v), str(year), src))

    pct = lambda x, g=g30: x / g * 100
    # Investment 1: need and cost
    emit(pct(nm.inv2_2030), "nmim_nz_inv_2030_pct_gdp")
    emit(pct(nm.dinv_2030), "nmim_incr_inv_2030_pct_gdp")
    for k in ["gen", "grid", "trn", "bld", "ind", "nco", "oth"]:
        emit(pct(nm[f"dinv_{k}_2030"]), f"nmim_incr_inv_{k}_2030_pct_gdp")
    emit(pct(nm.dinv_2050, g50), "nmim_incr_inv_2050_pct_gdp", year=2050)
    emit(pct(nm.ica_2030), "nmim_incr_cost_a_2030_pct_gdp")
    emit(pct(nm.icb_2030), "nmim_incr_cost_b_2030_pct_gdp")
    emit(pct(nm.icc_2030), "nmim_incr_cost_c_2030_pct_gdp")
    emit(pct(nm.ica_2030 - nm.icb_2030), "nmim_coc_penalty_2030_pct_gdp")
    abt = nm.ghg1_2030 - nm.ghg2_2030                       # MtCO2e; costs in USD bn
    emit((nm.icb_2030 * 1000 / abt).where(abt > 1e-6), "nmim_cost_per_t_b_2030_usd")
    emit(abt / nm.ghg1_2030 * 100, "nmim_abatement_2030_pct")
    emit(pct(nm.pv_ic_b), "nmim_pv_incr_cost_b_pct_gdp", year="2025-2050")

    # Investment 2: public investment needs
    emit(pct(nm.bud_2030), "nmim_budget_need_2030_pct_gdp")
    emit(pct(nm.typ1_2030), "nmim_typ1_2030_pct_gdp")
    emit(pct(k4 * nm.typ4_2030), "nmim_p4_gap_2030_pct_gdp")
    emit(pct(nm.dbud_2030), "nmim_budget_need_incr_2030_pct_gdp")
    pefa = long[long.indicator == "pefa_pi11_pim_score"].set_index("iso3").value
    income = long.drop_duplicates("iso3").set_index("iso3").income
    gap = pd.Series({iso: GAP_LIC - (pefa[iso] - 1) / 3 * (GAP_LIC - GAP_AE) if iso in pefa.index
                     else GAP_AE if nm.grp[iso] == "AE"
                     else GAP_LIC if income.get(iso) == "Low income" else GAP_EM for iso in nm.index})
    emit(pct(nm.bud_2030 / (1 - gap)), "nmim_budget_need_pima_2030_pct_gdp", src=SRC_PIMA)
    emit(pct(nm.typ1_2030 + nm.typ2_2030 + k4 * nm.typ4_2030), "nmim_budget_plus_util_2030_pct_gdp")
    emit(pct(nm.typ1_2030 + nm.typ2_2030 + nm.typ4_2030), "nmim_public_all_2030_pct_gdp")
    emit(pct(nm.prv_2030), "nmim_private_2030_pct_gdp")
    emit(pct(nm.typ3_2030), "nmim_typ3_2030_pct_gdp")
    emit(pct(nm.typ5_2030), "nmim_typ5_2030_pct_gdp")
    emit(pct((1 - k4) * nm.typ4_2030), "nmim_p4_private_2030_pct_gdp")
    emit(nm.prv_2030 / nm.inv2_2030 * 100, "nmim_private_share_2030_pct")
    return pd.DataFrame(out, columns=["iso3", "indicator", "value", "year", "source"])


# name, extract function, block its blocks follow in reading order
MODULES = [("investment_needs", extract_investment_needs, "Context 3")]


# ---------------------------------------------------------------- merge and build
def main():
    os.makedirs(MERGED, exist_ok=True)
    long = pd.read_csv(os.path.join(DASH, "output", "global_data_long.csv"), dtype={"year": str})
    cfg = pd.read_csv(os.path.join(DASH, "config", "indicators.csv")).sort_values("order")
    src = pd.read_csv(os.path.join(DASH, "config", "sources.csv"))
    meta = long.drop_duplicates("iso3").set_index("iso3")[["country", "region", "income", "is_ccdr"]]

    for name, extract, after in MODULES:
        rows = extract(long)
        dropped = sorted(set(rows.iso3) - set(meta.index))
        if dropped:
            print(f"{name}: {len(dropped)} economies not in the dashboard, left out: {', '.join(dropped)}")
        rows = rows[rows.iso3.isin(meta.index)]
        rows = rows.join(meta, on="iso3")[long.columns]
        rows.to_csv(os.path.join(OUT, f"{name}_long.csv"), index=False)
        mcfg = pd.read_csv(os.path.join(CONFIG, f"{name}_indicators.csv"))
        missing = set(mcfg.code) - set(rows.indicator)
        if missing:
            sys.exit(f"MODULE/CONFIG MISMATCH ({name}): config names indicators the module never emits: {sorted(missing)}")
        clash = set(mcfg.code) & set(cfg.code)
        if clash:
            sys.exit(f"{name}: indicator codes already in the dashboard: {sorted(clash)}")
        at = cfg.index[cfg.block == after].max()
        pos = list(cfg.index).index(at) + 1
        cfg = pd.concat([cfg.iloc[:pos], mcfg, cfg.iloc[pos:]], ignore_index=True)
        src = pd.concat([src, pd.read_csv(os.path.join(CONFIG, f"{name}_sources.csv"))], ignore_index=True)
        long = pd.concat([long, rows], ignore_index=True)
        print(f"{name}: {len(rows)} rows, {rows.iso3.nunique()} economies, {len(mcfg)} indicators")

    cfg["order"] = range(1, len(cfg) + 1)
    long.to_csv(os.path.join(MERGED, "global_data_long.csv"), index=False)
    cfg.to_csv(os.path.join(MERGED, "indicators.csv"), index=False)
    src.to_csv(os.path.join(MERGED, "sources.csv"), index=False)
    shutil.copy(os.path.join(CONFIG, "tool.json"), os.path.join(MERGED, "tool.json"))

    subprocess.run([sys.executable, os.path.join(DASH, "scripts", "build_dashboard.py"),
                    "--out", MERGED, "--config", MERGED, "--ship", SHIP], check=True)


if __name__ == "__main__":
    main()
