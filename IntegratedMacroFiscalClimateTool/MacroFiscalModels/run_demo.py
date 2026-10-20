"""Demo: (A) DIGNAD-lite disaster scenarios; (B) a country's NMIM net zero investment programme in
GIMF-lite and DIGNAD-lite; (C) the macro path fed into LIC-DSF-lite.

    python3 run_demo.py [ISO3]        (default KEN)

The macro calibrations and the DSF baseline ratios are illustrative placeholders, not the
country's: only the investment programme comes from NMIM. Outputs go to results/.
"""
import os
import sys
from dataclasses import replace
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import dignad_lite as dg
import gimf_lite as gm
from lic_dsf import Inputs, assess, print_assessment, composite_indicator, THRESHOLDS
from nmim_bridge import programme, run_gimf, run_dignad, dsf_overlay

ISO = sys.argv[1] if len(sys.argv) > 1 else "KEN"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results")
os.makedirs(OUT, exist_ok=True)
Y0 = 2025


def grid(panels, series, title, fname, x, xlabel):
    fig, axs = plt.subplots(2, 3, figsize=(14, 7.5))
    for ax, (ttl, f) in zip(axs.flat, panels):
        for name, s in series.items():
            ax.plot(x, f(s), label=name, lw=1.8)
        if "dev" in ttl:
            ax.axhline(0, color="grey", lw=0.6)
        ax.set_title(ttl, fontsize=10); ax.set_xlabel(xlabel)
    axs[0, 0].legend(fontsize=8)
    fig.suptitle(title); fig.tight_layout(); fig.savefig(os.path.join(OUT, fname), dpi=130); plt.close(fig)


# ---------------------------------------------------------------- A. DIGNAD-lite: disaster
H = 20
base = dg.Params()
scen = {
    "No reconstruction": replace(base, rec=0.0, sh_grant=0, sh_conc=0),
    "Reconstruction (standard)": base,
    "Reconstruction, build back better": replace(base, bbb=1.0),
    "Pre-disaster adaptation + BBB": replace(base, ig1=0.04, ig2=0.03, bbb=1.0),
}
dres = {n: dg.solve(p)[:2] for n, p in scen.items()}
dev = lambda o, s, v: 100 * (o[v][:H] / s[v] - 1)
grid([("Real GDP, % dev.", lambda r: dev(*r, "y")),
      ("Private capital, % dev.", lambda r: dev(*r, "k")),
      ("Public capital, % dev.", lambda r: 100 * ((r[0]["kg1"] + r[0]["kg2"])[:H] / (r[1]["kg1"] + r[1]["kg2"]) - 1)),
      ("Hand-to-mouth consumption, % dev.", lambda r: dev(*r, "ch")),
      ("Public debt, % GDP", lambda r: 100 * r[0]["debt_gdp"][:H]),
      ("Commercial interest rate, %", lambda r: 100 * r[0]["rc"][:H])],
     dres, "DIGNAD-lite: disaster (10% of capital destroyed, 3% TFP loss) under alternative responses",
     "dignad_disaster.png", np.arange(1, H + 1), "years after disaster")
print("A. DIGNAD-lite disaster: GDP % dev. in years 1 / 5 / 10; debt/GDP in year 10")
for n, (o, s) in dres.items():
    g = dev(o, s, "y")
    print(f"   {n:36s} {g[0]:6.2f} {g[4]:6.2f} {g[9]:6.2f}   {100 * o['debt_gdp'][9]:5.1f}")

# ---------------------------------------------------------------- B. NMIM programme
prog = programme(ISO)
yrs = prog["years"]
n = len(yrs)
print(f"\nB. NMIM programme for {prog['name']} ({ISO}), % of GDP:")
for y in (2025, 2030, 2040, 2050):
    i = y - Y0
    print(f"   {y}: public {100 * prog['pub'][i]:.2f}  private {100 * prog['prv'][i]:.2f}  energy saving {100 * prog['sav'][i]:.2f}")

gp = gm.Params()
gres = {
    "Full programme, public part debt financed": run_gimf(prog, gp),
    "Full programme, public part tax financed": run_gimf(prog, replace(gp, dfs=0.0, instrument="tau_l")),
    "Full programme, no energy savings": run_gimf(prog, gp, savings=False),
    "Public part only": run_gimf(prog, gp, private=False, savings=False),
}
T = n
yrsx = np.arange(Y0, Y0 + T)
gdev = lambda r, v: 100 * (r[0][v][:T] / r[1][v] - 1)
grid([("Real GDP, % dev.", lambda r: gdev(r, "y")),
      ("Consumption, % dev.", lambda r: 100 * (r[0]["c"][:T] / (r[1]["co"] + r[1]["cl"]) - 1)),
      ("Private investment (excl. green), % dev.", lambda r: gdev(r, "inv")),
      ("Public debt, % GDP", lambda r: 100 * r[0]["debt_gdp"][:T]),
      ("Inflation, %", lambda r: 100 * r[0]["pi"][:T]),
      ("Real exchange rate, % dev. (+ = depreciation)", lambda r: gdev(r, "q"))],
     gres, f"GIMF-lite: NMIM net zero investment programme, {prog['name']} (illustrative macro calibration)",
     f"gimf_{ISO}.png", yrsx, "year")
print("\n   GIMF-lite: GDP % dev. 2026 / 2030 / 2040 / 2050; public debt % GDP 2050")
for nm, r in gres.items():
    g = gdev(r, "y")
    print(f"   {nm:44s} {g[1]:6.2f} {g[5]:6.2f} {g[15]:6.2f} {g[25]:6.2f}   {100 * r[0]['debt_gdp'][25]:5.1f}")

dres2 = {
    "30% grants, 30% concessional, 40% commercial": run_dignad(prog),
    "All commercial": run_dignad(prog, dg.Params(eps_k=0, eps_g=0, eps_a=0, rec=0, sh_grant=0, sh_conc=0)),
    "30/30/40, low efficiency (s = 0.4)": run_dignad(prog, dg.Params(eps_k=0, eps_g=0, eps_a=0, rec=0, s=0.4)),
}
print("\n   DIGNAD-lite (public part only): GDP % dev. 2030 / 2040 / 2050; public debt % GDP 2050")
for nm, (o, s, c) in dres2.items():
    g = 100 * (o["y"] / s["y"] - 1)
    print(f"   {nm:44s} {g[5]:6.2f} {g[15]:6.2f} {g[25]:6.2f}   {100 * o['debt_gdp'][25]:5.1f}")

# paths to CSV
rows = {"year": yrsx, "nmim_pub_gdp": prog["pub"], "nmim_prv_gdp": prog["prv"], "nmim_sav_gdp": prog["sav"]}
for nm, r in gres.items():
    k = "gimf_" + nm.split(",")[0].replace(" ", "_").lower() + ("_tax" if "tax" in nm else "") + ("_nosav" if "no energy" in nm else "")
    rows[k + "_gdp_dev"] = gdev(r, "y")
    rows[k + "_debt_gdp"] = r[0]["debt_gdp"][:T]
pd.DataFrame(rows).to_csv(os.path.join(OUT, f"paths_{ISO}.csv"), index=False, float_format="%.6g")

# ---------------------------------------------------------------- C. DSF
ci, cls = composite_indicator(cpia=3.3, growth=0.05, reserves_cov=0.30, remittances=0.04, world_growth=0.033)
scale = prog["gdp"][0] / Inputs().gdp0 / 1.0
base_inp = Inputs()
base_inp = replace(base_inp, gdp0=base_inp.gdp0 * scale, exports0=base_inp.exports0 * scale,
                   ext_stock0=base_inp.ext_stock0 * scale, dom_stock0=base_inp.dom_stock0 * scale)
o, s0 = gres["Full programme, public part debt financed"][:2]
pd_ss = gm.solve(gp)[0]["prim_def"][0]
nz_inp = dsf_overlay(base_inp, o["y"], o["prim_def"], pd_ss)
print(f"\nC. LIC-DSF-lite (illustrative baseline ratios, scaled to {ISO}'s GDP; CI {ci:.2f} -> {cls})")
print("== baseline =="); a0 = assess(base_inp, cls); print_assessment(a0)
print("== baseline + net zero programme (GIMF-lite, debt financed) =="); a1 = assess(nz_inp, cls); print_assessment(a1)

fig, axs = plt.subplots(1, 5, figsize=(18, 3.8))
for ax, k in zip(axs, THRESHOLDS):
    yr = Y0 - 1 + a0["base"]["years"]
    ax.plot(yr, a0["base"]["ind"][k], label="Baseline", lw=2)
    ax.plot(yr, a1["base"]["ind"][k], label="Baseline + net zero programme", lw=2)
    w = a1["rows"][k]["worst_test"]
    ax.plot(yr, a1["tests"][w]["ind"][k], "--", label="Programme, most extreme shock", lw=1.5)
    ax.axhline(a0["thresholds"][k], color="red", lw=1, ls=":", label="Threshold")
    ax.set_title(f"{k}\n(worst: {w})", fontsize=9)
axs[0].legend(fontsize=7)
fig.suptitle(f"LIC-DSF-lite ({cls} capacity, illustrative baseline). Risk: baseline {a0['overall']}, "
             f"with programme {a1['overall']}")
fig.tight_layout(); fig.savefig(os.path.join(OUT, f"dsf_{ISO}.png"), dpi=130); plt.close(fig)
print(f"\nSaved results/dignad_disaster.png, gimf_{ISO}.png, dsf_{ISO}.png, paths_{ISO}.csv")
