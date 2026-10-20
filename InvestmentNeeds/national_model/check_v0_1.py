#!/usr/bin/env python3
# =============================================================================
# NMIM v0.1 - checks (LibreOffice)
# =============================================================================
# 1. Workbook = Python: for each test country, set Settings!C4, recalculate the
#    workbook headless and read the Check sheet's max abs difference (key rows,
#    2024-2050 sample years). Pass if relative difference < 1e-6.
# 2. Identities in the Python results for every country: generation balance,
#    no negative capacity or investment, net zero emissions below baseline in 2050,
#    total investment = sum of sectors.
# 3. Benchmarks: world net zero investment 2030 vs IHLEG/IEA, baseline GHG 2030,
#    energy CO2 2024 vs CPAT.
# Writes check_report_v0_1.md.
#
# Run: python InvestmentNeeds/national_model/check_v0_1.py  (needs soffice; set RECALC_PY
#      to a recalc helper or it converts with soffice directly)
# =============================================================================
import os, sys, shutil, subprocess, tempfile
import openpyxl, pandas as pd, numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, "NMIM_v0.1.xlsx")
TEST = ["IND", "USA", "CHN", "BRA", "KEN", "EGY", "DEU", "IDN"]

def recalc(path):
    rc = os.environ.get("RECALC_PY")
    if rc:
        subprocess.run([sys.executable, rc, path, "300"], check=True, capture_output=True)
    else:
        out = tempfile.mkdtemp()
        subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out, path],
                       check=True, capture_output=True, timeout=600)
        shutil.move(os.path.join(out, os.path.basename(path)), path)

lines = ["# NMIM v0.1 check report", ""]
ok_all = True
lines += ["## 1. Workbook = Python (live formulas vs engine, key rows)", "", "| Country | Max abs diff | Max abs value | Relative | Result |", "| --- | --- | --- | --- | --- |"]
tmp = tempfile.mkdtemp()
for iso in TEST:
    f = os.path.join(tmp, f"chk_{iso}.xlsx")
    wb = openpyxl.load_workbook(WB); wb["Settings"]["C4"] = iso; wb.save(f)
    recalc(f)
    c = openpyxl.load_workbook(f, data_only=True)["Check"]
    diff = c["B3"].value
    vals = [abs(x) for row in c.iter_rows(min_row=6, max_row=21, min_col=2, max_col=6, values_only=True) for x in row if isinstance(x, (int, float))]
    mx = max(vals) if vals else 1.0
    rel = (diff or 0) / max(mx, 1e-9)
    res = "PASS" if diff is not None and rel < 1e-6 else "FAIL"
    ok_all &= res == "PASS"
    lines.append(f"| {iso} | {diff:.2e} | {mx:,.1f} | {rel:.1e} | {res} |")
    print(iso, diff, res, flush=True)

r = pd.read_csv(os.path.join(HERE, "results", "nmim_results_v0_1.csv"))
yrs = [str(y) for y in range(2024, 2051)]
def get(code):
    return r[r.code.str.split(".", n=1).str[1] == code].set_index("iso")[yrs]
lines += ["", "## 2. Identities (all countries, all years)", ""]
checks = []
for sc in (1, 2):
    bal = get(f"mit.gen.fos.{sc}") + get(f"mit.gen.vre.{sc}")
    # generation >= fossil + vre (clean firm fills the rest); vre and fossil non-negative
    checks.append((f"fossil + VRE <= generation, scenario {sc}", bool((bal <= get(f"mit.gen.tot.{sc}") + 1e-6).all().all())))
    for code in (f"mit.inv.gen.{sc}", f"mit.inv.grid.{sc}", f"mit.cap.sol.{sc}", f"mit.cap.fos.stk.{sc}", f"mit.inv.trn.{sc}"):
        checks.append((f"{code} >= 0", bool((get(code) >= -1e-9).all().all())))
tot = get("mit.inv.tot.2"); parts = sum(get(f"mit.inv.{k}.2") for k in ("gen", "grid", "trn", "bld", "ind", "nco", "oth", "cal"))
checks.append(("total net zero investment = sum of sectors", bool(((tot - parts).abs() < 1e-6).all().all())))
checks.append(("net zero GHG 2050 below baseline (every country)", bool((get("mit.ghg.tot.2")["2050"] <= get("mit.ghg.tot.1")["2050"] + 1e-6).all())))
for name, ok in checks:
    ok_all &= ok
    lines.append(f"- {'PASS' if ok else 'FAIL'}: {name}")

s = pd.read_csv(os.path.join(HERE, "results", "nmim_summary_v0_1.csv"))
w = s[["inv1_2030", "inv2_2030", "dinv_2030", "ghg1_2030", "ghg2_2030", "ghg2_2050", "co2_model_2024", "co2_cpat_2024"]].sum()
emdc = s[s.grp == "EMDC"][["inv2_2030"]].sum()
lines += ["", "## 3. Benchmarks (world)", "", "| Item | Model | Benchmark | Source |", "| --- | --- | --- | --- |",
          f"| Net zero investment 2030, USD bn | {w.inv2_2030:,.0f} | 4,800 clean energy + 400 nature | IHLEG (2024), IEA |",
          f"| of which EMDCs ex-China | {emdc.inv2_2030:,.0f} | 1,600 clean energy + 350 nature | IHLEG (2024) |",
          f"| Baseline investment 2030 (calibrated to 2024 by group) | {w.inv1_2030:,.0f} | ~2,000 clean energy in 2024 | IEA World Energy Investment 2024 |",
          f"| Baseline GHG 2030, MtCO2e | {w.ghg1_2030:,.0f} | ~53,000-57,000 | UNEP Emissions Gap (current policies, approx.) |",
          f"| Energy CO2 2024, 11 end-use rows + power vs CPAT | {w.co2_model_2024:,.0f} | {w.co2_cpat_2024:,.0f} | CPAT 2024 panel; the gap is carried as other energy CO2 (mit.co2.res), so totals tie to CPAT |",
          f"| Net zero GHG 2050 (residual) | {w.ghg2_2050:,.0f} | ~0 net with removals | Residual to be met by removals (not modelled) |"]
lines += ["", f"**Overall: {'PASS' if ok_all else 'FAIL'}**"]
open(os.path.join(HERE, "check_report_v0_1.md"), "w").write("\n".join(lines) + "\n")
print("\n".join(lines))
