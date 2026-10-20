"""Standalone feasibility-constrained effectiveness metrics for power-sector
carbon pricing: carbon tax / ETS, feebate (= output-based allocation /
tradable performance standard) and tax + feebate hybrids.

For each country it computes, for a 2030-style medium-term response:

  E   abatement, % of baseline power-sector CO2
  F1  visible price: first-year rise in the residential tariff (%)
  F2  household loss: net burden on the bottom 40%, % of their consumption,
      after rebating a share of tax revenue to them
  F3  concentrated losers: net payment by coal generators, % of their revenue

and the best policy mix that keeps F1-F3 under user thresholds, which
constraint binds, and its shadow price.

Country data come from the TaxMacroFiscalDashboard panel and Ember (see
build_calibration). Every parameter that is NOT country data is in PRIORS and
is an assumption to be replaced (ideally by CPAT runs). See README.md.

Usage:
    python3 price_feasibility.py                      # default thresholds
    python3 price_feasibility.py --f1 5 --f2 0.3 --f3 10 --rebate 0.1
"""
import argparse
import os

import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
PANEL = os.path.join(REPO, "TaxMacroFiscalDashboard", "build", "output", "global_data_wide.csv")
EMBER = os.path.join(REPO, "TaxMacroFiscalDashboard", "build", "raw_data",
                     "release_generation_yearly_global.csv")

# --- Priors: assumptions, not data. Replace with country evidence / CPAT. ----
PRIORS = {
    "hr_coal": 10.0,          # GJ fuel per MWh, subcritical coal (~36% efficiency)
    "hr_gas": 7.0,            # GJ fuel per MWh, CCGT (~51% efficiency)
    "elasticity": -0.3,       # long-run price elasticity of electricity demand
    # tariff pass-through by power market structure (PSRT):
    # 3 wholesale market, 2 single buyer, 1 vertically integrated
    "pass_through": {3: 1.0, 2: 0.6, 1: 0.4},
    # when PSRT has no entry (it covers mostly developing countries)
    "pass_through_missing": {"High income": 0.9, "other": 0.6},
    "switch_missing": 50.0,   # $/t when no gas price at all
    "kappa": 0.4,             # max share of coal+oil generation displaceable by 2030
    "rise_scale": (0.5, 1.0), # kappa multiplier at RISE score 0 and 100
    "rise_missing": 35.0,     # median RISE RE score, used when missing
    "e_replace_share_gas": 0.3,  # share of displaced fossil output replaced by gas (rest clean)
    "switch_floor": 15.0,     # $/t: lowest switching price (non-price barriers)
    "wholesale_share": 0.6,   # wholesale price as share of residential tariff
    "res_share": 0.35,        # residential share of electricity sales
    "q40": 0.2,               # bottom-40% share of residential electricity use
    # electricity budget share of bottom 40%, by income group (placeholder)
    "b40_share": {"Low income": 0.02, "Lower middle income": 0.035,
                  "Upper middle income": 0.04, "High income": 0.03},
    "phase_in_years": 1,      # tax reaches its level in this many equal steps
}

DEFAULT_THRESHOLDS = {"f1": 10.0, "f2": 0.5, "f3": 15.0}
# $/tCO2: fine steps at low rates, where low tariffs make F1 move fast
TAU_GRID = np.concatenate([np.arange(0, 20, 0.2), np.arange(20, 201, 1.0)])


def build_calibration(panel_path=PANEL, ember_path=EMBER, year=None):
    """One row per country: data used by the model, with their sources."""
    em = pd.read_csv(ember_path)
    em = em[em["Area type"] == "Country or economy"]
    if year is None:
        year = int(em.loc[em["Generation (TWh)"].notna(), "Year"].max()) - 1
    em = em[em.Year == year]
    gen = em.pivot_table(index="ISO 3 code", columns="Electricity source",
                         values="Generation (TWh)", aggfunc="sum")
    emi = em.pivot_table(index="ISO 3 code", columns="Electricity source",
                         values="Emissions (MtCO2e)", aggfunc="sum")
    cal = pd.DataFrame({
        "gen_coal_twh": gen.get("Coal"), "gen_gas_twh": gen.get("Gas"),
        "gen_oil_twh": gen.get("Other fossil"), "gen_total_twh": gen.get("Total generation"),
        "em_coal_mt": emi.get("Coal"), "em_gas_mt": emi.get("Gas"),
        "em_oil_mt": emi.get("Other fossil"),
    })
    cal.index.name = "iso3"
    cal["ember_year"] = year

    p = pd.read_csv(panel_path).set_index("iso3")
    keep = ["country", "region", "income",
            "price_electricity_residential_usd_kwh", "price_coal_power_usd_gj",
            "price_gas_power_usd_gj", "price_gas_industry_usd_gj", "power_market_structure",
            "rise_renewable_energy_score", "coal_rents_pct_gdp",
            "ff_subsidy_explicit_pct_gdp", "fuel_exports_pct_merch"]
    cal = p[keep].join(cal, how="inner")
    cal = cal[(cal.gen_total_twh > 0) & cal.price_electricity_residential_usd_kwh.notna()]
    return cal.fillna({"gen_coal_twh": 0, "gen_gas_twh": 0, "gen_oil_twh": 0,
                       "em_coal_mt": 0, "em_gas_mt": 0, "em_oil_mt": 0})


def country_parameters(cal, priors=PRIORS):
    """Derive model parameters per country from calibration data and priors."""
    d = pd.DataFrame(index=cal.index)
    fossil_em = cal.em_coal_mt + cal.em_gas_mt + cal.em_oil_mt
    d["e0_mt"] = fossil_em
    d["e_avg"] = fossil_em / cal.gen_total_twh                     # tCO2/MWh
    d["e_coal"] = (cal.em_coal_mt / cal.gen_coal_twh).where(cal.gen_coal_twh > 0, 0.85)
    e_gas = (cal.em_gas_mt / cal.gen_gas_twh).where(cal.gen_gas_twh > 0, 0.40)
    e_oil = (cal.em_oil_mt / cal.gen_oil_twh).where(cal.gen_oil_twh > 0, 0.70)
    d["p_res"] = cal.price_electricity_residential_usd_kwh * 1000     # $/MWh
    miss = priors["pass_through_missing"]
    d["phi"] = cal.power_market_structure.map(priors["pass_through"]).fillna(
        cal.income.map(lambda g: miss.get(g, miss["other"])))

    # Substitution: displace up to kappa of coal+oil output, scaled by RISE.
    rise = cal.rise_renewable_energy_score.fillna(priors["rise_missing"]) / 100
    lo, hi = priors["rise_scale"]
    disp_twh = priors["kappa"] * (lo + (hi - lo) * rise) * (cal.gen_coal_twh + cal.gen_oil_twh)
    hi_em = cal.gen_coal_twh + cal.gen_oil_twh
    e_hi = ((cal.em_coal_mt + cal.em_oil_mt) / hi_em).where(hi_em > 0, 0)
    e_rep = priors["e_replace_share_gas"] * e_gas
    d["s_max"] = (disp_twh * (e_hi - e_rep).clip(lower=0) / fossil_em).where(fossil_em > 0, 0)

    # Coal-to-gas switching price from country fuel prices and heat rates.
    p_gas = cal.price_gas_power_usd_gj.fillna(cal.price_gas_industry_usd_gj)
    gap = (p_gas * priors["hr_gas"]
           - cal.price_coal_power_usd_gj * priors["hr_coal"])          # $/MWh
    tau_switch = gap / (d.e_coal - e_gas).where(d.e_coal - e_gas > 0.1, 0.1)
    d["tau_switch"] = tau_switch.fillna(priors["switch_missing"]).clip(lower=priors["switch_floor"])

    d["e_hi"] = e_hi
    d["b40"] = cal.income.map(priors["b40_share"]).fillna(0.03)
    d["hi_gen_share"] = (hi_em / cal.gen_total_twh).fillna(0)
    return d


def evaluate(par, tau_t, tau_f, rebate=0.0, priors=PRIORS):
    """Metrics for tax tau_t and feebate tau_f ($/t, scalars or broadcastable
    arrays). par is one row (Series) of country_parameters."""
    tau_t = np.asarray(tau_t, float)
    tau_f = np.asarray(tau_f, float)
    tau_m = tau_t + tau_f
    a_s = par.s_max * (1 - np.exp(-tau_m / par.tau_switch))
    dp = par.phi * par.e_avg * tau_t / par.p_res                 # tariff rise, fraction
    a_o = np.minimum(-priors["elasticity"] * dp, 1.0)
    e = 1 - (1 - a_s) * (1 - a_o)
    f1 = 100 * dp / priors["phase_in_years"]
    # B40 needs phi*res_share*q40 of revenue to be fully compensated.
    need = par.phi * priors["res_share"] * priors["q40"]
    cover = np.minimum(rebate / need, 1.0) if need > 0 else 1.0
    f2 = 100 * par.b40 * dp * (1 - cover)
    # Coal and oil generators: unpassed tax cost plus net feebate payment.
    p_whole = priors["wholesale_share"] * par.p_res
    f3 = 100 * (tau_t * par.e_hi * (1 - par.phi) + tau_f * max(par.e_hi - par.e_avg, 0)) / p_whole
    f3 = f3 * (par.hi_gen_share > 0)
    return {"E": 100 * e, "A_sub": 100 * a_s, "A_dem": 100 * a_o, "F1": f1, "F2": f2, "F3": f3}


def best(par, thresholds, rebate, mode, priors=PRIORS):
    """Highest-abatement (tau_t, tau_f) on the grid meeting all thresholds."""
    g = TAU_GRID
    if mode == "tax":
        tt, tf = g, np.zeros_like(g)
    elif mode == "feebate":
        tt, tf = np.zeros_like(g), g
    else:
        tt, tf = [a.ravel() for a in np.meshgrid(g, g)]
    m = evaluate(par, tt, tf, rebate, priors)
    ok = ((m["F1"] <= thresholds["f1"]) & (m["F2"] <= thresholds["f2"])
          & (m["F3"] <= thresholds["f3"]))
    if not ok.any():
        return None
    score = np.where(ok, m["E"] - 1e-6 * (tt + tf), -np.inf)  # tie-break: lower rates
    i = int(np.argmax(score))
    out = {k: float(np.asarray(v).ravel()[i] if np.ndim(v) else v) for k, v in m.items()}
    out.update(tau_t=float(tt[i]), tau_f=float(tf[i]))
    out["binding"] = _binding(par, out, thresholds, rebate, mode, priors)
    return out


def _binding(par, out, thresholds, rebate, mode, priors):
    """Constraints violated by one more grid step in a rate the mode can raise."""
    nxt = TAU_GRID[np.searchsorted(TAU_GRID, max(out["tau_t"], out["tau_f"]), "right"):]
    step = (nxt[0] - max(out["tau_t"], out["tau_f"])) if len(nxt) else 1.0
    moves = {"tax": [(step, 0)], "feebate": [(0, step)], "hybrid": [(step, 0), (0, step)]}[mode]
    hit = set()
    capped = True
    for dt, df in moves:
        tt, tf = out["tau_t"] + dt, out["tau_f"] + df
        if max(tt, tf) > TAU_GRID.max():
            continue
        capped = False
        m = evaluate(par, tt, tf, rebate, priors)
        hit |= {k for k in ("F1", "F2", "F3") if float(m[k]) > thresholds[k.lower()]}
    if hit:
        return "+".join(sorted(hit))
    return "rate cap" if capped else "none"


def country_metrics(cal, thresholds=DEFAULT_THRESHOLDS, rebate=0.0, priors=PRIORS):
    par = country_parameters(cal, priors)
    rows = []
    for iso, p in par.iterrows():
        if p.e0_mt <= 0:
            continue
        r = {"iso3": iso, "country": cal.at[iso, "country"], "income": cal.at[iso, "income"],
             "baseline_mt": p.e0_mt, "e_avg_t_mwh": p.e_avg, "pass_through": p.phi,
             "tau_switch": p.tau_switch, "s_max_pct": 100 * p.s_max,
             "hi_gen_share": p.hi_gen_share}
        ref_tax = evaluate(p, 50, 0, rebate, priors)
        ref_fb = evaluate(p, 0, 50, rebate, priors)
        r["E_tax50"] = float(ref_tax["E"])
        r["F1_tax50"] = float(ref_tax["F1"])
        r["E_feebate50"] = float(ref_fb["E"])
        r["F3_feebate50"] = float(ref_fb["F3"])
        r["s_at50"] = float(ref_tax["A_sub"] / (ref_tax["A_sub"] + ref_tax["A_dem"])) \
            if ref_tax["E"] > 0 else np.nan
        r["feebate_tax_ratio50"] = r["E_feebate50"] / r["E_tax50"] if r["E_tax50"] > 0 else np.nan
        r["E_per_pp_F1_tax50"] = r["E_tax50"] / r["F1_tax50"] if r["F1_tax50"] > 0 else np.nan
        for mode in ("tax", "feebate", "hybrid"):
            b = best(p, thresholds, rebate, mode, priors)
            for k in ("E", "tau_t", "tau_f", "F1", "F2", "F3", "binding"):
                r[f"{mode}_{k}"] = b[k] if b else (np.nan if k != "binding" else "infeasible")
        relaxed = dict(thresholds, f1=thresholds["f1"] + 1)
        b1 = best(p, relaxed, rebate, "hybrid", priors)
        r["hybrid_shadow_F1_pp"] = (b1["E"] - r["hybrid_E"]) if b1 and r["hybrid_E"] == r["hybrid_E"] else np.nan
        b1t = best(p, relaxed, rebate, "tax", priors)
        r["tax_shadow_F1_pp"] = (b1t["E"] - r["tax_E"]) if b1t and r["tax_E"] == r["tax_E"] else np.nan
        r["hybrid_mt_saved"] = p.e0_mt * r["hybrid_E"] / 100 if r["hybrid_E"] == r["hybrid_E"] else np.nan
        rows.append(r)
    return pd.DataFrame(rows).set_index("iso3")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--f1", type=float, default=DEFAULT_THRESHOLDS["f1"], help="max tariff rise, %%")
    ap.add_argument("--f2", type=float, default=DEFAULT_THRESHOLDS["f2"], help="max B40 net burden, %% of consumption")
    ap.add_argument("--f3", type=float, default=DEFAULT_THRESHOLDS["f3"], help="max coal generator net payment, %% of revenue")
    ap.add_argument("--rebate", type=float, default=0.0, help="share of tax revenue rebated to bottom 40%%")
    ap.add_argument("--out", default=os.path.join(HERE, "output"))
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    cal = build_calibration()
    cal.to_csv(os.path.join(a.out, "calibration_inputs.csv"))
    th = {"f1": a.f1, "f2": a.f2, "f3": a.f3}
    m = country_metrics(cal, th, a.rebate)
    m.round(3).to_csv(os.path.join(a.out, "country_metrics.csv"))
    big = m.sort_values("baseline_mt", ascending=False).head(15)
    cols = ["baseline_mt", "pass_through", "tau_switch", "s_at50", "tax_E", "feebate_E",
            "hybrid_E", "hybrid_tau_t", "hybrid_tau_f", "hybrid_binding", "tax_shadow_F1_pp",
            "hybrid_shadow_F1_pp"]
    print(f"Thresholds {th}, rebate {a.rebate}; {len(m)} countries -> {a.out}")
    print(big[cols].round(2).to_string())


if __name__ == "__main__":
    main()
