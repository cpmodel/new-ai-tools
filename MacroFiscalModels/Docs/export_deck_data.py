"""Writes Docs/deck_data.json (the numbers in the pitch deck charts and tables). Run from MacroFiscalModels/."""
import json, os, sys
from dataclasses import replace
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import dignad_lite as dg
import gimf_lite as gm
from nmim_bridge import programme, run_gimf, dsf_overlay
from lic_dsf import Inputs, assess, composite_indicator

ISO = sys.argv[1] if len(sys.argv) > 1 else "KEN"
d = {"iso": ISO}
prog = programme(ISO); gp = gm.Params()
sc = {"Debt financed": run_gimf(prog, gp),
      "Tax financed": run_gimf(prog, replace(gp, dfs=0.0, instrument="tau_l")),
      "No energy savings": run_gimf(prog, gp, savings=False),
      "Public part only": run_gimf(prog, gp, private=False, savings=False)}
d["name"] = prog["name"]
d["years"] = list(range(2025, 2051))
d["gimf"] = {k: [round(float(100 * (r[0]["y"][t] / r[1]["y"] - 1)), 2) for t in range(26)] for k, r in sc.items()}
d["gimf_debt"] = {k: [round(float(100 * r[0]["debt_gdp"][t]), 1) for t in range(26)] for k, r in sc.items()}
d["prog"] = {k: [round(float(100 * prog[k][t]), 2) for t in range(26)] for k in ("pub", "prv", "sav")}
b = dg.Params()
ds = {"No reconstruction": replace(b, rec=0.0, sh_grant=0, sh_conc=0), "Reconstruction": b,
      "Build back better": replace(b, bbb=1.0), "Adaptation first + BBB": replace(b, ig1=0.04, ig2=0.03, bbb=1.0)}
d["dignad"] = {}
for k, p in ds.items():
    o, s, c = dg.solve(p)
    d["dignad"][k] = [round(float(100 * (o["y"][t] / s["y"] - 1)), 2) for t in range(20)]
ci, cls = composite_indicator(cpia=3.3, growth=0.05, reserves_cov=0.30, remittances=0.04, world_growth=0.033)
f = prog["gdp"][0] / Inputs().gdp0
bi = Inputs(); bi = replace(bi, gdp0=bi.gdp0 * f, exports0=bi.exports0 * f, ext_stock0=bi.ext_stock0 * f, dom_stock0=bi.dom_stock0 * f)
o = sc["Debt financed"][0]; pss = gm.solve(gp)[0]["prim_def"][0]
a0 = assess(bi, cls); a1 = assess(dsf_overlay(bi, o["y"], o["prim_def"], pss), cls)
d["dsf"] = [[k, r["threshold"], round(float(r["baseline_max"]), 1), round(float(a1["rows"][k]["baseline_max"]), 1),
             round(float(a1["rows"][k]["stress_max"]), 1), a1["rows"][k]["worst_test"]] for k, r in a0["rows"].items()]
d["dsf_rating"] = [a0["overall"], a1["overall"]]
d["ci"] = [round(float(ci), 2), cls]
json.dump(d, open(os.path.join(HERE, "deck_data.json"), "w"), indent=1)
print("wrote deck_data.json")
