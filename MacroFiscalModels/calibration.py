"""One calibration for GIMF-lite and DIGNAD-lite, from the Tax and Macro-Fiscal Dashboard panel.

The settings the two models share (efficiency, return on public capital, depreciation, capital
share, public investment, taxes, debt, constrained households, borrowing cost) are set once here,
from dashboard data where it exists and from income-group defaults where it does not. Every
value carries its source, so data and assumptions can be told apart. The rest of each model's
parameters keep their defaults.

    python3 calibration.py KEN SLE UGA LBR      # calibration, model choice, cross-check

Functions: shared(iso) -> dict of (value, source); to_gimf(sh) / to_dignad(sh) -> Params;
choose_model(iso, disaster_prone) -> (name, reason); cross_check(iso) -> both models on the
public part of the NMIM programme.
"""
import os
import sys
from dataclasses import replace
import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
PANEL = os.path.join(HERE, "..", "TaxMacroFiscalDashboard", "build", "output", "global_data_wide.csv")

INCOME = ["Low income", "Lower middle income", "Upper middle income", "High income"]
# income-group defaults (author's assumptions unless a source is named), LIC / LMIC / UMIC / HIC
DEFAULTS = {
    "pim_gap": ((0.40, 0.27, 0.27, 0.13), "IMF (2015) efficiency gap by income group"),
    "Rg": ((0.20, 0.17, 0.15, 0.10), "assumption: return on public capital higher where capital is scarce (IMF literature 10-25%)"),
    "ig": ((0.06, 0.05, 0.045, 0.035), "assumption: public investment % GDP by income group (IMF ICSD orders of magnitude)"),
    "constrained": ((0.60, 0.50, 0.35, 0.25), "assumption: credit-constrained share (Findex orders of magnitude)"),
    "conc_share": ((0.70, 0.45, 0.15, 0.0), "assumption: concessional share of public debt"),
    "grants": ((0.02, 0.01, 0.0, 0.0), "assumption: grants % GDP"),
    "nfd": ((0.30, 0.30, 0.20, 0.0), "assumption: net foreign debt % GDP"),
    "tax": ((0.12, 0.15, 0.20, 0.25), "assumption: tax revenue % GDP"),
    "debt": ((0.50, 0.55, 0.60, 0.70), "assumption: gross public debt % GDP"),
}
COMMON = dict(delta=0.07, delta_g=0.04, alpha=0.33)  # the same in both models (assumption)
TAX_SPLIT = dict(cons=0.40, lab=0.45, cap=0.15)       # shares of tax revenue (assumption)
BASES = dict(cons=0.65, lab=0.61, cap=0.39)           # approximate tax bases, % GDP
TR_TARGET = 0.02                                      # transfers in the starting steady state, % GDP


def _panel():
    if not hasattr(_panel, "df"):
        _panel.df = pd.read_csv(PANEL).set_index("iso3")
    return _panel.df


def _num(row, col):
    v = row.get(col, np.nan) if row is not None else np.nan
    return float(v) if pd.notna(v) else None


def shared(iso):
    """Shared settings for one country: {name: (value, source)}."""
    df = _panel()
    row = df.loc[iso] if iso in df.index else None
    inc = row["income"] if row is not None and row["income"] in INCOME else "Lower middle income"
    k = INCOME.index(inc)
    out = {"income": (inc, "dashboard (World Bank income group)" if row is not None else "default")}

    def dflt(name):
        vals, src = DEFAULTS[name]
        return vals[k], f"default: {src}"

    # efficiency: PEFA PI-11 mapped as in InvestmentNeeds v1.1 (score 1 -> LIC gap, 4 -> AE gap)
    pefa = _num(row, "pefa_pi11_pim_score")
    if pefa is not None:
        gap = 0.40 - (pefa - 1) / 3 * (0.40 - 0.13)
        out["s"] = (1 - gap, f"dashboard: PEFA PI-11 = {pefa:g}, mapped as in InvestmentNeeds v1.1")
    else:
        g, src = dflt("pim_gap")
        out["s"] = (1 - g, src)
    out["Rg"] = dflt("Rg")
    out["ig"] = dflt("ig")
    out["constrained"] = dflt("constrained")
    out["grants"] = dflt("grants")
    out["nfd"] = dflt("nfd")
    d = _num(row, "gross_debt_pct_gdp")
    out["debt"] = (min(max(d / 100, 0.1), 1.5), "dashboard: gross debt (IMF WEO)") if d is not None else dflt("debt")
    t = _num(row, "tax_revenue_pct_gdp")
    out["tax"] = (min(max(t / 100, 0.05), 0.40), "dashboard: tax revenue (% GDP)") if t is not None else dflt("tax")
    m = _num(row, "ids_multilateral_share_pct")
    out["conc_share"] = (m / 100, "dashboard: IDS multilateral share of PPG external debt (proxy)") if m is not None else dflt("conc_share")
    pr = _num(row, "ids_private_creditor_rate_pct")
    if pr is not None:
        out["rstar"] = (min(max(pr / 100 - 0.03, 0.03), 0.08), f"dashboard: IDS private creditor rate {pr:.1f}% less 3% inflation")
    else:
        out["rstar"] = (0.04 if k < 2 else 0.03, "default: real borrowing cost by income group")
    fr = _num(row, "fossil_rents_total_pct_gdp")
    og = (_num(row, "oil_rents_pct_gdp") or 0) + (_num(row, "gas_rents_pct_gdp") or 0)
    fx = _num(row, "fuel_exports_pct_merch")
    if fr is not None and fr >= 1 and og == 0 and fx is not None and fx < 5:
        # coal rents with no fuel exports (e.g. Cabo Verde, 17% of GDP in WDI): treated as a data error
        out["rents"] = (0.0, f"dashboard rents {fr:.1f}% ignored: coal only and fuel exports {fx:.1f}% (likely WDI error)")
    else:
        out["rents"] = (fr / 100, "dashboard: fossil rents (WDI oil + gas + coal)") if fr is not None else (0.0, "default: no rents data")
    for c, v in COMMON.items():
        out[c] = (v, "default: common to both models")
    # government consumption closes the starting budget so transfers start near TR_TARGET
    T, b, r = out["tax"][0], out["debt"][0], out["rstar"][0]
    g = T + out["grants"][0] - r * b - out["ig"][0] - TR_TARGET
    out["g"] = (max(g, 0.05), "residual: closes the starting budget (floor 5% of GDP)")
    return _feasible(out)


def _feasible(out):
    """GIMF-lite needs a discount factor below 1. Where wealth is large relative to the OLG
    households' consumption (low-income countries with many constrained households), raise the
    shared borrowing cost in steps of 0.5 points until it is, and flag the change."""
    from gimf_lite import calibrate
    r0 = out["rstar"][0]
    for step in range(11):
        try:
            calibrate(to_gimf(out))
            if step:
                out["rstar"] = (out["rstar"][0], out["rstar"][1] + f"; raised from {r0:.3f} for GIMF-lite feasibility")
            return out
        except ValueError:
            T, b = out["tax"][0], out["debt"][0]
            out["rstar"] = (out["rstar"][0] + 0.005, out["rstar"][1])
            g = T + out["grants"][0] - out["rstar"][0] * b - out["ig"][0] - TR_TARGET
            out["g"] = (max(g, 0.05), out["g"][1])
    raise ValueError("no feasible GIMF-lite calibration")


def values(sh):
    return {k: v[0] for k, v in sh.items()}


def to_gimf(sh, **over):
    from gimf_lite import Params
    v = values(sh)
    T = v["tax"]
    p = Params(rstar=v["rstar"], s=v["s"], Rg=v["Rg"], delta=v["delta"], delta_g=v["delta_g"], alpha=v["alpha"],
               ig=v["ig"], g=v["g"], b=v["debt"], nfd=v["nfd"], psi=v["constrained"],
               tau_c=TAX_SPLIT["cons"] * T / BASES["cons"], tau_l=TAX_SPLIT["lab"] * T / BASES["lab"],
               tau_k=TAX_SPLIT["cap"] * T / BASES["cap"])
    return replace(p, **over)


def to_dignad(sh, **over):
    from dignad_lite import Params
    v = values(sh)
    T, b = v["tax"], v["debt"]
    p = Params(rstar=v["rstar"], s=v["s"], Rg=v["Rg"], delta=v["delta"], delta_g=v["delta_g"], alpha=v["alpha"],
               ig1=0.85 * v["ig"], ig2=0.15 * v["ig"], gc=v["g"], grants=v["grants"], phi=v["constrained"],
               tau_c=TAX_SPLIT["cons"] * T / BASES["cons"], tau_y=(1 - TAX_SPLIT["cons"]) * T,
               dk=b * v["conc_share"], dc=b * (1 - v["conc_share"]), bp=0.5 * v["nfd"],
               eps_k=0, eps_g=0, eps_a=0, rec=0)
    return replace(p, **over)


RENT_DECLINE_2050 = 0.8   # assumption: fossil rents fall 80% by 2050 on a global net zero path (IEA NZE order of magnitude)
GOV_TAKE = 0.5            # assumption: government share of fossil rents
RESOURCE_RICH = 0.10      # rents of at least 10% of GDP route a non-high-income country to DIGNAR-lite


def to_dignar(sh, years=26, **over):
    """DIGNAD-lite with the resource block: rents as % of non-resource GDP, falling RENT_DECLINE_2050
    by the last year; a firmer fiscal rule (a permanent revenue loss needs one)."""
    r = values(sh)["rents"]
    r0 = r / (1 - r)
    path = tuple(r0 * (1 - RENT_DECLINE_2050 * np.linspace(0, 1, years)))
    kw = dict(lam_tau=0.2, lam_tr=0.2)
    kw.update(over)
    return to_dignad(sh, rents0=r0, gov_take=GOV_TAKE, rents_path=path, **kw)


def choose_model(iso, disaster_prone=False):
    """DIGNAR-lite for resource-rich non-high-income countries; DIGNAD-lite for low-income,
    concessional-reliant or disaster-prone countries; GIMF-lite otherwise."""
    v = values(shared(iso))
    if v["rents"] >= RESOURCE_RICH and v["income"] != "High income":
        return "DIGNAR-lite", f"fossil rents {100 * v['rents']:.0f}% of GDP"
    if v["income"] == "Low income":
        return "DIGNAD-lite", "low-income country"
    if disaster_prone:
        return "DIGNAD-lite", "disaster-prone"
    if v["income"] == "Lower middle income" and v["conc_share"] >= 0.5:
        return "DIGNAD-lite", f"lower-middle income relying on concessional finance ({100 * v['conc_share']:.0f}% multilateral)"
    return "GIMF-lite", f"{v['income'].lower()}, mainly market finance"


def cross_check(iso, prog=None):
    """Both models on the public part of the NMIM programme, with the same shared settings."""
    from nmim_bridge import programme, run_gimf, run_dignad
    prog = prog or programme(iso)
    sh = shared(iso)
    og, sg, _, _ = run_gimf(prog, to_gimf(sh), private=False, savings=False)
    od, sd, _ = run_dignad(prog, to_dignad(sh))
    res = {}
    for name, o, s in (("GIMF-lite", og, sg), ("DIGNAD-lite", od, sd)):
        y = 100 * (o["y"] / s["y"] - 1)
        res[name] = dict(gdp_2030=y[5], gdp_2050=y[25], debt_2050=100 * o["debt_gdp"][25])
    res["spend_share"] = float(np.mean(prog["spend_share"][:11]))
    return res, sh


def report(iso):
    res, sh = cross_check(iso)
    model, why = choose_model(iso)
    print(f"\n=== {iso}: {sh['income'][0]} -> {model} ({why})")
    for k in ["s", "Rg", "ig", "tax", "debt", "conc_share", "constrained", "rstar", "g", "grants"]:
        print(f"   {k:12s} {sh[k][0]:7.3f}   {sh[k][1]}")
    print("   cross-check, public part of the NMIM programme: GDP % dev. 2030 / 2050, debt % GDP 2050")
    for name in ("GIMF-lite", "DIGNAD-lite"):
        r = res[name]
        print(f"   {name:12s} {r['gdp_2030']:6.2f} {r['gdp_2050']:6.2f}   {r['debt_2050']:6.1f}")
    if res["spend_share"] > 0.25:
        print(f"   note: {100 * res['spend_share']:.0f}% of incremental investment 2025-2035 is non-CO2 / land-use / other"
              " spending, which builds no public capital")
    return res


if __name__ == "__main__":
    for iso in (sys.argv[1:] or ["KEN"]):
        report(iso)
