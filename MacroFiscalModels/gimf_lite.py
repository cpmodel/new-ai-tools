"""GIMF-lite: a one-country, annual simplification of the IMF Global Integrated Monetary and
Fiscal Model (Kumhof, Laxton, Muir & Mursula 2010, IMF WP/10/34).

Kept from GIMF
  * non-Ricardian households: Blanchard-Yaari overlapping generations with planning horizon
    1/(1 - theta) and declining lifetime labour income (chi), so human wealth is discounted at
    roughly (1 + r) / (theta * chi), above the government's borrowing rate
  * liquidity-constrained households (share psi) who spend current income
  * firms with private capital (adjustment costs, Tobin's q) and productive public capital
  * monopolistic competition with sticky prices (New Keynesian Phillips curve) and sticky wages,
    so output is demand-determined in the short run
  * inflation-targeting central bank (Taylor rule) and uncovered interest parity with a
    debt-elastic country premium: the real exchange rate and the current account adjust
  * fiscal block: government consumption, public investment, transfers, labour, consumption and
    capital taxes, and a debt/deficit fiscal rule with a chosen instrument
Simplifications
  * one country facing an exogenous world (world real rate, world demand), not 2-6 regions
  * one good; trade enters through net exports that respond to the real exchange rate
  * no oil or commodity block, no financial accelerator, no zero lower bound
  * linear Phillips curve (no Rotemberg resource cost); annual frequency

Climate add-on (for the link with InvestmentNeeds/NMIM)
  * ig_extra: incremental public investment (direct budget need), % of GDP
  * ix: incremental private investment, building a 'green' capital stock KX that adds zeta
    units of effective capital per unit (zeta = 0: it only replaces fossil capacity)
  * sav: energy cost savings, % of GDP, entering as a gain in value added (TFP)
  * dfs: share of the extra public investment that the fiscal rule lets be debt financed

Units: steady-state GDP = 1, so ratios are shares of initial GDP.
"""
from dataclasses import dataclass, replace
import numpy as np
from pf_solver import solve_path, solve_static


@dataclass
class Params:
    rstar: float = 0.03       # world real interest rate
    pibar: float = 0.03       # inflation target
    theta: float = 0.95       # survival / planning horizon: 1/(1-theta) = 20 years
    chi: float = 0.97         # decline of lifetime labour income
    psi: float = 0.25         # liquidity-constrained share of households
    alpha: float = 0.33       # private capital share
    Rg: float = 0.15          # gross return on public capital (calibrates its output elasticity)
    delta: float = 0.08
    delta_g: float = 0.04
    s: float = 0.7            # public-investment efficiency (PIMA)
    kappa: float = 4.0        # investment adjustment cost
    mu: float = 1.10          # price markup
    kappa_p: float = 0.25     # Phillips-curve slope (annual)
    rho_w: float = 0.5        # wage stickiness
    wage_el: float = 1.0      # elasticity of the target real wage to employment
    phi_pi: float = 1.5       # Taylor rule
    rho_i: float = 0.5
    eta_f: float = 0.02       # country premium per unit of net foreign debt / GDP
    eps_x: float = 1.0        # export elasticity to the real exchange rate
    eps_m: float = 0.5        # import elasticity
    m_share: float = 0.30     # imports / absorption
    # fiscal (shares of GDP)
    g: float = 0.12
    ig: float = 0.04
    tau_l: float = 0.15
    tau_c: float = 0.10
    tau_k: float = 0.20
    b: float = 0.50           # government debt target
    nfd: float = 0.30         # net foreign debt
    lam_b: float = 0.05       # fiscal rule: speed of return to the debt target
    instrument: str = "tr"    # "tr" (transfers) or "tau_l" (labour tax) closes the rule
    # climate add-on
    zeta: float = 0.0
    dfs: float = 1.0
    T: int = 200


VARS = ["y", "l", "w", "mc", "rk", "pi", "i", "r", "q", "k", "inv", "tq", "v", "hw",
        "co", "cl", "b", "fis", "nfd", "nx"]
NV = len(VARS)
IX = {v: j for j, v in enumerate(VARS)}


def calibrate(p: Params):
    """Initial steady state (y = l = q = 1, no programme) and the implied structural constants."""
    r = p.rstar
    mc = 1 / p.mu
    rk = (r + p.delta) / (1 - p.tau_k)
    k = mc * p.alpha / rk
    inv = p.delta * k
    kg = p.s * p.ig / p.delta_g
    alpha_g = p.Rg * kg
    A = 1 / (kg ** alpha_g * k ** p.alpha)
    w = mc * (1 - p.alpha)
    nx = r * p.nfd
    c = 1 - inv - p.g - p.ig - nx
    tax = p.tau_l * w + p.tau_c * c + p.tau_k * (rk * k + 1 - mc)
    tr = tax - r * p.b - p.g - p.ig
    cl = p.psi * ((1 - p.tau_l) * w + tr) / (1 + p.tau_c)
    co = c - cl
    div = (1 - p.tau_k) * (rk * k + 1 - mc) - inv
    v = div / r
    hw = (1 - p.psi) * ((1 - p.tau_l) * w + tr) / (1 - p.theta * p.chi / (1 + r))
    fw = div + v + (1 + r) * p.b - (1 + r) * p.nfd
    beta = (1 - (1 + p.tau_c) * co / (fw + hw)) / p.theta
    if not (0 < beta < 1) or co <= 0 or cl < 0:
        raise ValueError(f"infeasible calibration: beta={beta:.3f}, co={co:.3f}, cl={cl:.3f}")
    m = p.m_share * (1 - nx)
    cal = dict(A=A, alpha_g=alpha_g, beta=beta, w_bar=w, x_bar=m + nx, m_bar=m, abs_bar=1 - nx,
               tr_bar=tr, mc_bar=mc, kg0=kg)
    fis = tr if p.instrument == "tr" else p.tau_l
    ss = dict(y=1.0, l=1.0, w=w, mc=mc, rk=rk, pi=p.pibar, i=(1 + r) * (1 + p.pibar) - 1, r=r, q=1.0,
              k=k, inv=inv, tq=1.0, v=v, hw=hw, co=co, cl=cl, b=p.b, fis=fis, nfd=p.nfd, nx=nx)
    return ss, cal


def exo_paths(p: Params, cal, ig_extra=(), ix=(), sav=()):
    """Programme paths for t = 1..T; each input is held at its last value after it ends."""
    def pad(a):
        a = np.asarray(a, float)
        out = np.zeros(p.T)
        if a.size:
            n = min(a.size, p.T)
            out[:n] = a[:n]
            out[n:] = a[n - 1]
        return out
    ige, ixp, svp = pad(ig_extra), pad(ix), pad(sav)
    ig = p.ig + ige
    kg = np.empty(p.T); kx = np.empty(p.T)
    kg_prev, kx_prev = cal["kg0"], 0.0
    for t in range(p.T):
        kg_prev = (1 - p.delta_g) * kg_prev + p.s * ig[t]
        kx_prev = (1 - p.delta) * kx_prev + ixp[t]
        kg[t], kx[t] = kg_prev, kx_prev
    # steady-state stocks after the programme (held constant beyond T)
    kg_inf = p.s * ig[-1] / p.delta_g
    kx_inf = ixp[-1] / p.delta
    return dict(ig=ig, ig_extra=ige, ix=ixp, sav=svp, kg=kg, kx=kx,
                kg_inf=kg_inf, kx_inf=kx_inf, ig_inf=ig[-1], ix_inf=ixp[-1], sav_inf=svp[-1],
                ige_inf=ige[-1])


def _resid(X, lag, lead, p, cal, ss0, exo):
    """Residuals for a path X (T, NV) given its lag (row t-1) and lead (row t+1) arrays.
    exo: dict of arrays (length T) ig, ige, ix, sav, kg_l (public capital at t-1), kx_l."""
    g = lambda M, v: M[:, IX[v]]
    y, l, w, mc, rk, pi, i, r, q, k, inv, tq, v, hw, co, cl, b, fis, nfd, nx = (X[:, j] for j in range(NV))
    k_l, i_l, r_l, b_l, nfd_l, w_l, y_l = (g(lag, s) for s in ["k", "i", "r", "b", "nfd", "w", "y"])
    pi_f, r_f, q_f, rk_f, inv_f, tq_f, v_f, hw_f = (g(lead, s) for s in ["pi", "r", "q", "rk", "inv", "tq", "v", "hw"])
    y_f, mc_f, k_f = g(lead, "y"), g(lead, "mc"), k
    tr = fis if p.instrument == "tr" else cal["tr_bar"]
    tau_l = p.tau_l if p.instrument == "tr" else fis
    keff_l = k_l + p.zeta * exo["kx_l"]
    A = cal["A"] * (1 + exo["sav"])
    prem = p.eta_f * (nfd - p.nfd)
    prem_l = p.eta_f * (nfd_l - p.nfd)
    x = inv / k_l
    adj = p.kappa / 2 * (x - p.delta) ** 2 * k_l
    pi_prof = (1 - mc) * y
    div = (1 - p.tau_k) * (rk * keff_l + pi_prof) - inv - adj - exo["ix"]
    # dividends next period (needed for the value of the firm): built from lead values
    keff = k + p.zeta * exo["kx"]
    x_f = inv_f / k
    adj_f = p.kappa / 2 * (x_f - p.delta) ** 2 * k
    div_f = (1 - p.tau_k) * (rk_f * keff + (1 - mc_f) * y_f) - inv_f - adj_f - exo["ix_f"]
    fw = div + v + (1 + r_l) * b_l - (1 + p.rstar + prem_l) * q * nfd_l
    c = co + cl
    absorb = c + inv + adj + exo["ix"] + p.g + exo["ig"]
    R = np.empty_like(X)
    R[:, 0] = y - A * exo["kg_l"] ** cal["alpha_g"] * keff_l ** p.alpha * l ** (1 - p.alpha)
    R[:, 1] = rk - mc * p.alpha * y / keff_l
    R[:, 2] = w - mc * (1 - p.alpha) * y / l
    R[:, 3] = np.log(w) - (p.rho_w * np.log(w_l) + (1 - p.rho_w) * (np.log(cal["w_bar"]) + p.wage_el * np.log(l)))
    R[:, 4] = (pi - p.pibar) - (1 / (1 + p.rstar) * (pi_f - p.pibar) + p.kappa_p * (mc - cal["mc_bar"]) / cal["mc_bar"])
    i_neutral = (1 + p.rstar + prem) * (1 + p.pibar) - 1                               # moves with the premium
    R[:, 5] = i - (p.rho_i * i_l + (1 - p.rho_i) * (i_neutral + p.phi_pi * (pi - p.pibar)))
    R[:, 6] = (1 + i) - (1 + r) * (1 + pi_f)
    R[:, 7] = (1 + r) - (1 + p.rstar + prem) * q_f / q                                 # UIP (real)
    R[:, 8] = tq - (1 + p.kappa * (x - p.delta))
    R[:, 9] = k - ((1 - p.delta) * k_l + inv)
    R[:, 10] = tq - ((1 - p.tau_k) * rk_f + (1 - p.delta) * tq_f
                     - p.kappa / 2 * (x_f - p.delta) ** 2 + p.kappa * (x_f - p.delta) * x_f) / (1 + r)
    R[:, 11] = v - (div_f + v_f) / (1 + r)                                            # value of firms
    lab_inc = (1 - tau_l) * w * l + tr
    R[:, 12] = hw - ((1 - p.psi) * lab_inc + p.theta * p.chi / (1 + r) * hw_f)        # human wealth
    R[:, 13] = (1 + p.tau_c) * co - (1 - cal["beta"] * p.theta) * (fw + hw)           # OLG consumption
    R[:, 14] = (1 + p.tau_c) * cl - p.psi * lab_inc                                   # liquidity-constrained
    tax = tau_l * w * l + p.tau_c * c + p.tau_k * (rk * keff_l + pi_prof)
    R[:, 15] = b - ((1 + r_l) * b_l + p.g + exo["ig"] + tr - tax)                     # government budget
    R[:, 16] = (b - b_l) - (p.dfs * exo["ige"] - p.lam_b * (b_l - p.b))               # fiscal rule
    R[:, 17] = nfd - ((1 + p.rstar + prem_l) * nfd_l - nx / q)                        # balance of payments
    R[:, 18] = nx - (cal["x_bar"] * q ** p.eps_x - cal["m_bar"] * (absorb / cal["abs_bar"]) * q ** (-p.eps_m))
    R[:, 19] = y - (absorb + nx)                                                      # goods market
    return R


def _exo_rows(exo, T):
    kg_l = np.concatenate([[exo["kg0"]], exo["kg"][:-1]])
    kx_l = np.concatenate([[0.0], exo["kx"][:-1]])
    ix_f = np.concatenate([exo["ix"][1:], [exo["ix_inf"]]])
    return dict(ig=exo["ig"], ige=exo["ig_extra"], ix=exo["ix"], ix_f=ix_f, sav=exo["sav"],
                kg_l=kg_l, kx_l=kx_l, kx=exo["kx"])


def terminal_ss(p, cal, ss0, exo):
    """Steady state with the programme's final values held forever (solved numerically)."""
    e = dict(ig=np.array([exo["ig_inf"]]), ige=np.array([exo["ige_inf"]]), ix=np.array([exo["ix_inf"]]),
             ix_f=np.array([exo["ix_inf"]]), sav=np.array([exo["sav_inf"]]),
             kg_l=np.array([exo["kg_inf"]]), kx_l=np.array([exo["kx_inf"]]), kx=np.array([exo["kx_inf"]]))
    pp = p
    def f(xv):
        X = xv[None, :]
        return _resid(X, X, X, pp, cal, ss0, e)[0]
    x0 = np.array([ss0[v] for v in VARS])
    return dict(zip(VARS, solve_static(f, x0)))


def solve(p: Params = None, ig_extra=(), ix=(), sav=(), verbose=False):
    p = p or Params()
    ss0, cal = calibrate(p)
    exo = exo_paths(p, cal, ig_extra, ix, sav)
    exo["kg0"] = cal["kg0"]
    ssT = terminal_ss(p, cal, ss0, exo)
    rows = _exo_rows(exo, p.T)
    x0, xT = np.array([ss0[v] for v in VARS]), np.array([ssT[v] for v in VARS])

    def F(X):
        lag = np.vstack([x0, X[:-1]])
        lead = np.vstack([X[1:], xT])
        return _resid(X, lag, lead, p, cal, ss0, rows)

    w8 = np.linspace(0, 1, p.T)[:, None] ** 0.5
    X = solve_path(F, (1 - w8) * x0 + w8 * xT, verbose=verbose)
    out = {v: X[:, IX[v]] for v in VARS}
    out.update(exo)
    tr = out["fis"] if p.instrument == "tr" else np.full(p.T, cal["tr_bar"])
    tau_l = np.full(p.T, p.tau_l) if p.instrument == "tr" else out["fis"]
    out["tr"], out["tau_l"] = tr, tau_l
    out["c"] = out["co"] + out["cl"]
    out["debt_gdp"] = out["b"] / out["y"]
    out["nfd_gdp"] = out["q"] * out["nfd"] / out["y"]
    keff_l = np.concatenate([[ss0["k"]], out["k"][:-1]]) + p.zeta * rows["kx_l"]
    out["revenue"] = (tau_l * out["w"] * out["l"] + p.tau_c * out["c"]
                      + p.tau_k * (out["rk"] * keff_l + (1 - out["mc"]) * out["y"]))
    out["prim_def"] = p.g + out["ig"] + tr - out["revenue"]
    # check: OLG households' budget, implied by Walras' law (assets end t = V + B - q*nfd)
    assets = out["v"] + out["b"] - out["q"] * out["nfd"]
    lag = lambda a, a0: np.concatenate([[a0], a[:-1]])
    r_l, b_l, nfd_l = lag(out["r"], ss0["r"]), lag(out["b"], ss0["b"]), lag(out["nfd"], ss0["nfd"])
    prem_l = p.eta_f * (nfd_l - p.nfd)
    inv, k_l = out["inv"], lag(out["k"], ss0["k"])
    adj = p.kappa / 2 * (inv / k_l - p.delta) ** 2 * k_l
    div = (1 - p.tau_k) * (out["rk"] * keff_l + (1 - out["mc"]) * out["y"]) - inv - adj - out["ix"]
    fw = div + out["v"] + (1 + r_l) * b_l - (1 + p.rstar + prem_l) * out["q"] * nfd_l
    lab = (1 - tau_l) * out["w"] * out["l"] + tr
    out["hh_budget_gap"] = assets - (fw + (1 - p.psi) * lab - (1 + p.tau_c) * out["co"])
    return out, ss0, ssT, cal


if __name__ == "__main__":
    p = Params()
    ss0, cal = calibrate(p)
    print("calibrated beta =", round(cal["beta"], 4), " alpha_g =", round(cal["alpha_g"], 3))
    # 1% of GDP permanent public investment increase, debt financed
    out, ss0, ssT, cal = solve(p, ig_extra=[0.01], verbose=True)
    for t in [0, 1, 4, 9, 19, 49]:
        print(f"t={t+1:2d} GDP {100*(out['y'][t]-1):6.2f}%  C {100*(out['c'][t]/(ss0['co']+ss0['cl'])-1):6.2f}%"
              f"  infl {100*out['pi'][t]:5.2f}  r {100*out['r'][t]:5.2f}  debt/GDP {100*out['debt_gdp'][t]:5.1f}"
              f"  NFD/GDP {100*out['nfd_gdp'][t]:5.1f}")
    print("max household budget gap (should be ~0):", np.abs(out["hh_budget_gap"]).max())
