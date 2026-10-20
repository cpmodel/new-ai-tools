"""Bridge from the National Mitigation Investment Model (InvestmentNeeds/national_model, NMIM v0.2)
to the macro-fiscal models in this folder.

For one country it runs the NMIM engine and returns annual paths, 2025-2050, as shares of NMIM's
real GDP (2024 prices):
  pub  - incremental direct budget need (P1 + public gap of P4), NMIM row mit.dtyp.bud
  prv  - incremental private investment: total incremental (mit.dinv.tot) less pub
  sav  - net energy cost saving: fuel savings (mit.dfuel) less extra O&M (mit.dom)
These feed GIMF-lite (ig_extra, ix, sav), DIGNAD-lite (prog) and, through the macro paths,
LIC-DSF-lite (growth and primary deficit).
"""
import os
import sys
from dataclasses import replace
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
NMIM_DIR = os.path.join(HERE, "..", "InvestmentNeeds", "national_model")
_CACHE = {}


def _nmim():
    if "m" not in _CACHE:
        sys.path.insert(0, NMIM_DIR)
        import nmim_v0_2 as m
        D, _ = m.load_calibrated()
        _CACHE["m"], _CACHE["D"] = m, D
    return _CACHE["m"], _CACHE["D"]


def programme(iso, first=2025, last=2050):
    m, D = _nmim()
    if iso not in D:
        raise KeyError(f"{iso} not in NMIM data")
    S, V = m.run_country(D[iso])
    yi = [m.YEARS.index(y) for y in range(first, last + 1)]
    gdp = np.array([V["mit.gdp.lvl"][i] for i in yi])
    get = lambda c: np.array([V[c][i] for i in yi])
    pub = get("mit.dtyp.bud") / gdp
    tot = get("mit.dinv.tot") / gdp
    sav = (get("mit.dfuel") - get("mit.dom")) / gdp
    return dict(iso=iso, name=D[iso]["name"], years=np.arange(first, last + 1), gdp=gdp,
                pub=pub, prv=tot - pub, tot=tot, sav=sav, wacc=S["wacc"])


def run_gimf(prog, p=None, public=True, private=True, savings=True):
    from gimf_lite import Params, solve
    p = p or Params()
    z = np.zeros_like(prog["pub"])
    return solve(p, ig_extra=prog["pub"] if public else z, ix=prog["prv"] if private else z,
                 sav=prog["sav"] if savings else z)


def run_dignad(prog, p=None):
    """DIGNAD-lite sees only the public part (it has no private investment push); no disaster."""
    from dignad_lite import Params, solve
    p = p or Params(eps_k=0, eps_g=0, eps_a=0, rec=0)
    return solve(replace(p, prog=tuple(prog["pub"])))


def dsf_overlay(inp, y_path, prim_def_path, prim_def_ss, n=None):
    """Shift a LIC-DSF-lite baseline by a macro model's deviations: growth by the change in the
    model's GDP growth, primary deficit by the change in the model's primary deficit (% GDP)."""
    n = n or inp.N
    y = np.concatenate([[1.0], np.asarray(y_path[:n])])
    g_dev = y[1:] / y[:-1] - 1
    d_pd = (np.asarray(prim_def_path[:n]) - prim_def_ss) / np.asarray(y_path[:n])
    return replace(inp, g=inp.g + g_dev, pd_gdp=inp.pd_gdp + d_pd, gx=inp.gx + g_dev)
