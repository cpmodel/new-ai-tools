"""Integrated country run: dashboard -> investment needs (NMIM) -> pricing feasibility -> macro model
(GIMF-lite, DIGNAD-lite or DIGNAR-lite) -> LIC-DSF-lite -> results for the dashboard.

    python3 integrated.py KEN SLE UGA LBR      # print the chain for some countries
    python3 integrated.py --all                # every NMIM country -> results/integrated_results.csv

Then rebuild the dashboard panel (TaxMacroFiscalDashboard/build/scripts/build_global_indicators.py),
which reads results/integrated_results.csv into Appendix D. This script reads only the panel's source
indicators, never Appendix D, so the order build panel -> integrated.py -> build panel is stable.

Steps for one country:
 1. Calibration (calibration.py): shared settings from the dashboard panel, income-group defaults
    where data are missing; model choice.
 2. Programme (nmim_bridge.py): incremental budget need split into capital and spending, private
    investment, energy savings, fossil power CO2 on the net zero path (% GDP, 2025-2050).
 3. Pricing (pricing_bridge.py): the feasible power-sector mix from PriceFeasibility; its tax part
    times NMIM's power CO2 is revenue earmarked for the budget need.
 4. Macro model: GIMF-lite takes the whole programme; DIGNAD-lite and DIGNAR-lite the public part
    (DIGNAR-lite adds falling fossil rents). If the fiscal rule cannot stabilise debt, the rule is
    tightened (twice); if it still cannot, the status says so.
 5. LIC-DSF-lite (low and lower-middle income only): a baseline from the panel, shifted by the
    macro model's growth and primary deficit; mechanical risk rating.
"""
import os
import sys
import warnings
from dataclasses import replace
import numpy as np
import pandas as pd

import calibration as cb
from nmim_bridge import programme, run_gimf, run_dignad, dsf_overlay, _nmim
from pricing_bridge import pricing_revenue
from lic_dsf import Inputs, Terms, assess

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "results", "integrated_results.csv")
Y0 = 2025
RISK_CODE = {"Low": 1, "Moderate": 2, "High": 3}
# DSF-only defaults (author's assumptions), LIC / LMIC
EXT_SHARE = (0.60, 0.50)      # external share of public debt
EXPORTS = (0.25, 0.30)        # exports of goods and services, % GDP (WEO extract has no country trade data)


def _panel_value(iso, col):
    df = cb._panel()
    if iso in df.index and col in df.columns and pd.notna(df.loc[iso, col]):
        return float(df.loc[iso, col])
    return None


def dsf_baseline(iso, prog, sh):
    """LIC-DSF-lite inputs from the panel and NMIM; returns (Inputs, notes on assumptions)."""
    v = cb.values(sh)
    k = 0 if v["income"] == "Low income" else 1
    notes = []
    gdp = prog["gdp"]
    g = gdp[1:11] / gdp[:10] - 1
    rev = _panel_value(iso, "revenue_pct_gdp")
    if rev is None:
        rev = 100 * (v["tax"] + 0.03); notes.append("revenue = tax + 3% of GDP")
    pb = _panel_value(iso, "primary_balance_pct_gdp")
    if pb is None:
        pb = -2.0; notes.append("primary deficit 2% of GDP")
    notes += ["external share of debt", "exports / GDP", "capacity medium (no CPIA)"]
    ext = v["debt"] * EXT_SHARE[k] * gdp[0]
    priv = v["rstar"] + 0.03                                   # nominal commercial rate
    conc = min(max(v["conc_share"], 0.0), 1.0)
    inp = Inputs(gdp0=gdp[0], g=g, exports0=EXPORTS[k] * gdp[0], gx=g + 0.02,
                 rev_gdp=np.full(10, rev / 100), pd_gdp=np.full(10, -pb / 100),
                 ext_stock0=ext, ext_rate0=conc * 0.01 + (1 - conc) * priv,
                 dom_stock0=v["debt"] * gdp[0] - ext,
                 new_terms=[Terms(conc, 0.01, 40, 10), Terms(1 - conc, priv, 10, 3)] if 0 < conc < 1
                 else [Terms(1.0, 0.01 if conc >= 1 else priv, 40 if conc >= 1 else 10, 10 if conc >= 1 else 3)])
    return inp, notes


def _run_macro(model, prog, sh, rev):
    """Run the chosen model, tightening the fiscal rule if debt cannot be stabilised."""
    tries = [{}, dict(k=0.2), dict(k=0.4)]
    last = None
    for t in tries:
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                if model == "GIMF-lite":
                    p = cb.to_gimf(sh, **({"lam_b": t["k"]} if t else {}))
                    out, ss0, _, _ = run_gimf(prog, p, rev=rev)
                    base = run_gimf({**prog, **{k: np.zeros_like(prog["pub"]) for k in
                                                ("pub", "pub_cap", "pub_spend", "prv", "sav")}}, p)[0]
                    y = out["y"] / ss0["y"]
                    return dict(y=y, debt=out["debt_gdp"], pdef=out["prim_def"], pdef0=base["prim_def"][0],
                                y_base=base["y"] / ss0["y"], debt_base=base["debt_gdp"],
                                rule=p.lam_b, tightened=bool(t))
                over = {"lam_tau": t["k"], "lam_tr": t["k"]} if t else {}
                p = (cb.to_dignar(sh, **over) if model == "DIGNAR-lite" else cb.to_dignad(sh, **over))
                out, ss, _ = run_dignad(prog, p, rev=rev)
                gdp0 = 1 + p.rents0
                base = run_dignad({**prog, **{k: np.zeros_like(prog["pub"]) for k in ("pub", "pub_cap", "pub_spend")}},
                                  replace(p, rents_path=()))[0]            # no programme, no rent decline
                # baseline for the charts: no programme, but rents fall as in the programme run (global transition)
                bl = run_dignad({**prog, **{k: np.zeros_like(prog["pub"]) for k in ("pub", "pub_cap", "pub_spend")}}, p)[0]
                return dict(y=out["gdp"] / gdp0, debt=out["debt_gdp"], pdef=-out["prim_bal"],
                            pdef0=-base["prim_bal"][0], y_base=bl["gdp"] / gdp0, debt_base=bl["debt_gdp"],
                            rule=p.lam_tr, tightened=bool(t))
        except (RuntimeError, ValueError, np.linalg.LinAlgError) as e:
            last = e
    raise RuntimeError(f"not financeable under the fiscal rule ({last})")


def run(iso, mode="hybrid", paths=None):
    """One country. If `paths` is a dict, it is filled with annual series (2025-2050, % of GDP or %)
    and the calibration settings, for the country tool."""
    sh = cb.shared(iso)
    v = cb.values(sh)
    model, why = cb.choose_model(iso)
    prog = programme(iso)
    rev, mix = pricing_revenue(prog, mode)
    i30, i50 = 2030 - Y0, 2050 - Y0
    r = dict(iso3=iso, name=prog["name"], income=v["income"], model=model, model_reason=why,
             assumptions=sum(1 for k in sh if str(sh[k][1]).startswith("default")),
             budget_need_2030_pct_gdp=100 * prog["pub"][i30], budget_capital_2030_pct_gdp=100 * prog["pub_cap"][i30],
             budget_spending_2030_pct_gdp=100 * prog["pub_spend"][i30], private_2030_pct_gdp=100 * prog["prv"][i30],
             budget_need_cum_pct_gdp=100 * prog["pub"].sum(), energy_saving_2050_pct_gdp=100 * prog["sav"][i50],
             pricing_tau_t=mix["tau_t"] if mix else np.nan, pricing_tau_f=mix["tau_f"] if mix else np.nan,
             pricing_binding=mix["binding"] if mix else "no data",
             pricing_F1=mix["F1"] if mix else np.nan, pricing_F2=mix["F2"] if mix else np.nan,
             pricing_F3=mix["F3"] if mix else np.nan, pricing_abatement_pct=mix["E"] if mix else np.nan,
             pricing_revenue_2030_pct_gdp=100 * rev[i30],
             pricing_share_of_need_pct=100 * rev.sum() / prog["pub"].sum() if prog["pub"].sum() > 0 else np.nan,
             financing_gap_2030_pct_gdp=100 * (prog["pub"][i30] - rev[i30]),
             rents_pct_gdp=100 * v["rents"], debt_2024_pct_gdp=100 * v["debt"])
    try:
        m = _run_macro(model, prog, sh, rev)
        r.update(status=f"ok, fiscal rule tightened to {m['rule']}" if m["tightened"] else "ok",
                 gdp_dev_2030_pct=100 * (m["y"][i30] - 1), gdp_dev_2050_pct=100 * (m["y"][i50] - 1),
                 debt_2030_pct_gdp=100 * m["debt"][i30], debt_2050_pct_gdp=100 * m["debt"][i50])
        r["debt_change_2050_pp"] = r["debt_2050_pct_gdp"] - r["debt_2024_pct_gdp"]
        if paths is not None:
            n = len(prog["years"])
            paths["gdp_dev_pct"] = 100 * (m["y"][:n] - 1)
            paths["debt_pct_gdp"] = 100 * m["debt"][:n]
            paths["gdp_dev_base_pct"] = 100 * (m["y_base"][:n] - 1)              # baseline: no programme
            paths["debt_base_pct_gdp"] = 100 * m["debt_base"][:n]
    except RuntimeError as e:
        r.update(status=str(e).split(" (")[0])
        m = None
    if v["income"] in ("Low income", "Lower middle income") and m is not None:
        inp, notes = dsf_baseline(iso, prog, sh)
        a0 = assess(inp, "medium")
        a1 = assess(dsf_overlay(inp, m["y"], m["pdef"], m["pdef0"]), "medium")
        r.update(dsf_rating_base=a0["overall"], dsf_rating_programme=a1["overall"],
                 dsf_risk_code=RISK_CODE[a1["overall"]], dsf_assumptions="; ".join(notes))
        if paths is not None:
            for k in a0["thresholds"]:
                paths[f"dsf_base|{k}"] = a0["base"]["ind"][k][1:]          # projection years 1-10 = 2025-2034
                paths[f"dsf_prog|{k}"] = a1["base"]["ind"][k][1:]
                paths[f"dsf_thr|{k}"] = a0["thresholds"][k]
    if paths is not None:
        for k in ("pub", "pub_cap", "pub_spend", "prv", "sav"):
            paths[k] = 100 * prog[k]
        paths["rev"] = 100 * rev
        paths["settings"] = {k: sh[k] for k in sh}
        paths["mix"] = mix
    return r


def show(r):
    print(f"\n=== {r['name']} ({r['iso3']}): {r['income']} -> {r['model']} ({r['model_reason']}); {r['assumptions']} settings are defaults")
    print(f"   Budget need 2030: {r['budget_need_2030_pct_gdp']:.2f}% GDP (capital {r['budget_capital_2030_pct_gdp']:.2f}, "
          f"spending {r['budget_spending_2030_pct_gdp']:.2f}); private {r['private_2030_pct_gdp']:.2f}%")
    print(f"   Pricing ({r['pricing_binding']} binds): tax ${r['pricing_tau_t']:.0f}/t + feebate ${r['pricing_tau_f']:.0f}/t "
          f"-> revenue {r['pricing_revenue_2030_pct_gdp']:.2f}% GDP in 2030, {r['pricing_share_of_need_pct']:.0f}% of the need 2025-2050")
    print(f"   Financing gap 2030: {r['financing_gap_2030_pct_gdp']:.2f}% GDP;  status: {r['status']}")
    if "gdp_dev_2050_pct" in r:
        print(f"   GDP vs today: {r['gdp_dev_2030_pct']:+.2f}% (2030), {r['gdp_dev_2050_pct']:+.2f}% (2050); "
              f"debt {r['debt_2024_pct_gdp']:.0f}% -> {r['debt_2030_pct_gdp']:.0f}% (2030) -> {r['debt_2050_pct_gdp']:.0f}% (2050)")
    if "dsf_rating_base" in r:
        print(f"   LIC-DSF-lite: {r['dsf_rating_base']} -> {r['dsf_rating_programme']}  (assumed: {r['dsf_assumptions']})")


PATHS_OUT = os.path.join(HERE, "results", "integrated_paths.csv")
SETTINGS_OUT = os.path.join(HERE, "results", "integrated_settings.csv")
YEARS = list(range(Y0, 2051))


def run_all(isos=None):
    """All countries -> integrated_results.csv (one row each), integrated_paths.csv (annual series,
    long: iso3, series, 2025..2050) and integrated_settings.csv (calibration value and source)."""
    m, D = _nmim()
    rows, prow, srow = [], [], []
    for iso in isos or sorted(D):
        p = {}
        try:
            rows.append(run(iso, paths=p))
        except Exception as e:                               # keep going; record the failure
            rows.append(dict(iso3=iso, status=f"error: {str(e)[:80]}"))
        for k, val in p.items():
            if k in ("settings", "mix"):
                continue
            a = np.atleast_1d(np.asarray(val, float))
            vals = list(a) + [np.nan] * (len(YEARS) - len(a)) if a.size > 1 else [float(a[0])] + [np.nan] * (len(YEARS) - 1)
            prow.append([iso, k] + vals[:len(YEARS)])
        for k, (val, src) in p.get("settings", {}).items():
            srow.append([iso, k, val, src])
    df = pd.DataFrame(rows)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    df.to_csv(OUT, index=False, float_format="%.4g")
    pd.DataFrame(prow, columns=["iso3", "series"] + YEARS).to_csv(PATHS_OUT, index=False, float_format="%.4g")
    pd.DataFrame(srow, columns=["iso3", "setting", "value", "source"]).to_csv(SETTINGS_OUT, index=False, float_format="%.4g")
    return df


if __name__ == "__main__":
    args = sys.argv[1:]
    if args == ["--all"]:
        df = run_all()
        print(f"wrote {OUT}: {len(df)} countries")
        print(df.status.value_counts().to_string())
        print(df.model.value_counts().to_string())
    else:
        for iso in args or ["KEN"]:
            show(run(iso))
