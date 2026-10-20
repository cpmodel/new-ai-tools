"""Bridge from PriceFeasibility (power-sector carbon tax vs feebate under political-feasibility limits
F1-F3) to the macro models.

For one country it takes the feasible mix that PriceFeasibility found (default: the hybrid) and turns
its tax part into a revenue path: tax rate (USD/t) x fossil power CO2 on NMIM's net zero path. A
feebate is revenue-neutral, so only the tax part raises money. The path, in % of GDP, is earmarked
for the public investment need (`rev` in DIGNAD-lite, `rev_extra` in GIMF-lite).

Reads PriceFeasibility/output/country_metrics.csv (run `python3 PriceFeasibility/price_feasibility.py`
to refresh it, with other thresholds if wanted).
"""
import os
import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
METRICS = os.path.join(HERE, "..", "PriceFeasibility", "output", "country_metrics.csv")


def _metrics():
    if not hasattr(_metrics, "df"):
        _metrics.df = pd.read_csv(METRICS).set_index("iso3")
    return _metrics.df


def feasible_mix(iso, mode="hybrid"):
    """Dict with the feasible tax and feebate rates, abatement, binding limit and F1-F3; None if absent."""
    df = _metrics()
    if iso not in df.index:
        return None
    r = df.loc[iso]
    return dict(mode=mode, tau_t=float(r[f"{mode}_tau_t"]), tau_f=float(r[f"{mode}_tau_f"]),
                E=float(r[f"{mode}_E"]), binding=str(r[f"{mode}_binding"]),
                F1=float(r[f"{mode}_F1"]), F2=float(r[f"{mode}_F2"]), F3=float(r[f"{mode}_F3"]),
                shadow_F1=float(r[f"{mode}_shadow_F1_pp"]) if f"{mode}_shadow_F1_pp" in r else np.nan)


def pricing_revenue(prog, mode="hybrid"):
    """Revenue path (% of GDP, 2025-2050) and the mix behind it, for an NMIM programme dict."""
    mix = feasible_mix(prog["iso"], mode)
    if mix is None:
        return np.zeros_like(prog["gdp"]), None
    rev_bn = mix["tau_t"] * prog["power_co2"] / 1000.0       # USD/t x Mt = USD m; / 1000 = USD bn
    return rev_bn / prog["gdp"], mix
