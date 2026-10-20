#!/usr/bin/env python3
"""Build the Integrated MF-C (macro-fiscal climate) tool.

The tool is the Global Fiscal Indicators dashboard with extra modules spliced in as blocks.
Each module turns a repository model's results into rows of the dashboard's long format
(iso3, country, region, income, is_ccdr, indicator, value, year, source) and brings its own
indicator and source rows (same columns as the dashboard's config). This script writes the
merged panel and config to build/output/merged/ and runs the dashboard's own builder on them,
so the integrated workbook reads, scores and looks exactly like the dashboard.

Inputs : TaxMacroFiscalDashboard/build/{output/global_data_long.csv, config/}
         per module, see MODULES (now: InvestmentNeeds v1.1 workbook and NMIM v0.2 summary)
Output : IntegratedMFC/Integrated_MFC_Tool_v<version>.xlsx (version in config/tool.json)
         build/output/<module>_long.csv (the module's rows, committed)

To add a module: write an extract function returning long rows, add
config/<name>_indicators.csv and config/<name>_sources.csv, and register it in MODULES.
"""
import os, sys, shutil, subprocess
import numpy as np, pandas as pd, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
DASH = os.path.join(ROOT, "TaxMacroFiscalDashboard", "build")
INV = os.path.join(ROOT, "InvestmentNeeds")
CONFIG = os.path.join(HERE, "config")
OUT = os.path.join(HERE, "output")
MERGED = os.path.join(OUT, "merged")
SHIP = os.path.abspath(os.path.join(HERE, ".."))


# ---------------------------------------------------------------- module: investment needs
V11 = os.path.join(INV, "Mitigation_Investment_Needs_v1.1.xlsx")
NMIM = os.path.join(INV, "national_model", "results", "nmim_summary_v0_2.csv")
SRC_V11 = "Mitigation_Investment_Needs_v1.1.xlsx (InvestmentNeeds v1.1, model output)"
SRC_NMIM = "nmim_summary_v0_2.csv (InvestmentNeeds NMIM v0.2, model output)"


def _sheet(wb, name, header_row=4):
    """A results sheet of the v1.1 workbook as a frame. Reads the values Excel or LibreOffice
    cached on the last save; the workbook must have been recalculated (the build does that)."""
    rows = list(wb[name].iter_rows(min_row=header_row, values_only=True))
    df = pd.DataFrame(rows[1:], columns=rows[0]).dropna(subset=["ISO3"])
    return df


def extract_investment_needs():
    wb = openpyxl.load_workbook(V11, read_only=True, data_only=True)
    rc = _sheet(wb, "Results_Country")
    calc = _sheet(wb, "Calc")
    if rc["Total investment I"].isna().all():
        sys.exit(f"{V11} has no cached values: open and save it in Excel, or recalculate with LibreOffice.")
    rc = rc.set_index("ISO3")
    gdp = rc["GDP 2030 (USD bn)"].astype(float)
    out = []

    def emit(series, code, src, year=2030):
        for iso, v in series.items():
            if v is not None and pd.notna(v) and np.isfinite(float(v)):
                out.append((iso, code, float(v), str(year), src))

    # Investment 1: v1.1, 2030 snapshot. Shares of GDP come as fractions; report percent.
    emit(rc["I, % GDP"].astype(float) * 100, "inn_total_inv_2030_pct_gdp", SRC_V11)
    emit(rc["dI, % GDP"].astype(float) * 100, "inn_incr_inv_2030_pct_gdp", SRC_V11)
    sec = calc.pivot_table(index="ISO3", columns="Sector",
                           values="Incremental investment dI (USD bn)", aggfunc="sum")
    for k in ["GEN", "GRID", "TRN", "BLD", "IND", "FUEL", "AFOLU"]:
        emit(sec[k] / gdp.reindex(sec.index) * 100, f"inn_incr_inv_{k.lower()}_2030_pct_gdp", SRC_V11)
    emit(rc["(a) % GDP"].astype(float) * 100, "inn_incr_cost_a_2030_pct_gdp", SRC_V11)
    emit(rc["(b) % GDP"].astype(float) * 100, "inn_incr_cost_b_2030_pct_gdp", SRC_V11)
    emit(rc["(c) % GDP"].astype(float) * 100, "inn_incr_cost_c_2030_pct_gdp", SRC_V11)
    emit(rc["Incr. cost per t (b) USD/t"].astype(float), "inn_cost_per_t_b_2030_usd", SRC_V11)
    emit(rc["CoC penalty on I (a-b)"].astype(float) / gdp * 100, "inn_coc_penalty_2030_pct_gdp", SRC_V11)
    emit(rc["PIMA-adjusted I"].astype(float) / gdp * 100, "inn_pima_adj_inv_2030_pct_gdp", SRC_V11)

    # Investment 2: NMIM v0.2. USD bn over the model's own GDP for the year.
    nm = pd.read_csv(NMIM).set_index("iso")
    g30, g50 = nm.gdp_2030, nm.gdp_2050
    emit(nm.inv2_2030 / g30 * 100, "nmim_nz_inv_2030_pct_gdp", SRC_NMIM)
    for k in range(1, 6):
        emit(nm[f"typ{k}_2030"] / g30 * 100, f"nmim_typ{k}_2030_pct_gdp", SRC_NMIM)
    emit(nm.bud_2030 / g30 * 100, "nmim_budget_need_2030_pct_gdp", SRC_NMIM)
    emit(nm.prv_2030 / nm.inv2_2030 * 100, "nmim_private_share_2030_pct", SRC_NMIM)
    emit(nm.dinv_2030 / g30 * 100, "nmim_incr_inv_2030_pct_gdp", SRC_NMIM)
    emit(nm.dinv_2050 / g50 * 100, "nmim_incr_inv_2050_pct_gdp", SRC_NMIM, year=2050)
    emit(nm.pv_ic_b / g30 * 100, "nmim_pv_incr_cost_b_pct_gdp", SRC_NMIM, year="2025-2050")
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
        rows = extract()
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
