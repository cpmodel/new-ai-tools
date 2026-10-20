"""DIGNAD-lite: a simplified Debt-Investment-Growth and Natural Disasters model.

Inspired by Marto, Papageorgiou & Klyuev (2018, J. Dev. Econ.) / IMF DIGNAD toolkit,
itself built on the DIG model (Buffie, Berg, Pattillo, Portillo & Zanna, 2012).

Simplifications vs. full DIGNAD
  * one traded good (no tradable / non-tradable split, no real exchange rate)
  * inelastic labour (L = 1); log utility
  * no domestic debt (government borrows concessionally + on commercial external markets)
  * disaster is an unanticipated one-off shock at t = 1 (capital destruction + TFP loss),
    then perfect foresight; solved as a stacked non-linear system (Newton), like Dynare's
    perfect-foresight solver.

Kept from DIGNAD
  * savers (Ricardian, borrow abroad at a debt-elastic premium) and hand-to-mouth households
  * private capital with adjustment costs (Tobin's q)
  * two public capital stocks: standard and adaptation (more costly, more resilient)
  * public-investment efficiency s (only s of each $ becomes capital)
  * reconstruction spending financed by grants, concessional loans and commercial debt
  * debt-elastic commercial interest rate; fiscal rule closing the gap with consumption
    tax and transfers.

All quantities are per capita, normalised so that steady-state GDP = 1 (ratios = share of GDP).
"""
from dataclasses import dataclass, field, asdict
import numpy as np
from pf_solver import solve_path


@dataclass
class Params:
    # preferences / technology
    rstar: float = 0.04       # world real interest rate
    alpha: float = 0.35       # private capital share
    delta: float = 0.06       # private depreciation
    delta_g: float = 0.05     # public depreciation
    kappa: float = 5.0        # private investment adjustment cost
    phi: float = 0.5          # share of hand-to-mouth households
    # public investment
    s: float = 0.6            # public-investment efficiency (PIMA / PIE-X)
    Rg: float = 0.20          # steady-state gross return on public capital (calibrates psi)
    ig1: float = 0.06         # standard public investment, % GDP
    ig2: float = 0.01         # adaptation public investment, % GDP
    mu_a: float = 0.25        # cost premium of adaptation capital
    # fiscal
    tau_y: float = 0.10       # income tax rate (fixed)
    tau_c: float = 0.10       # consumption tax, steady state
    gc: float = 0.08          # government consumption, % GDP
    grants: float = 0.02      # grants, % GDP
    dk: float = 0.30          # concessional debt, % GDP
    rk: float = 0.01          # concessional real rate
    dc: float = 0.20          # commercial external debt, % GDP
    eta_g: float = 0.05       # commercial-debt risk-premium elasticity
    bp: float = 0.10          # private external debt, % GDP
    eta_p: float = 0.05       # private risk-premium elasticity
    lam_tau: float = 0.05     # fiscal rule: consumption-tax response to debt gap
    lam_tr: float = 0.05      # fiscal rule: transfer response to debt gap
    # disaster (fractions of stock destroyed at t = 1)
    eps_k: float = 0.10       # private capital destroyed
    eps_g: float = 0.10       # standard public capital destroyed
    rho_res: float = 0.80     # adaptation capital is destroyed at (1 - rho_res) * eps_g
    eps_a: float = 0.03       # TFP loss
    rho_a: float = 0.6        # TFP shock persistence
    theta_a: float = 0.5      # how much adaptation share reduces private/TFP damage
    # reconstruction & financing
    rec: float = 0.03         # extra public investment in year 1, % GDP
    rho_rec: float = 0.7      # decay of reconstruction spending
    bbb: float = 0.0          # 'build back better': share of reconstruction as adaptation
    sh_grant: float = 0.3     # share of reconstruction financed by grants
    sh_conc: float = 0.3      # share financed by new concessional loans (rest: commercial)
    amort_k: float = 0.03     # repayment rate of concessional debt
    # exogenous public investment programme (e.g. NMIM direct budget need), % GDP by year from t = 1;
    # financed with the same grant / concessional / commercial shares; tapered to zero after its end
    prog: tuple = ()
    prog_bbb: float = 0.0     # share of the programme that is adaptation capital
    prog_taper: int = 10
    T: int = 150


VARS = ["y", "cs", "ch", "i", "k", "q", "bp", "rp", "kg1", "kg2", "dc", "rc", "tr", "tauc"]
NV = len(VARS)
IX = {v: j for j, v in enumerate(VARS)}


def steady_state(p: Params):
    beta = 1 / (1 + p.rstar)
    k = (1 - p.tau_y) * p.alpha / (p.rstar + p.delta)          # y = 1
    kg1 = p.s * p.ig1 / p.delta_g
    kg2 = p.s * p.ig2 / (1 + p.mu_a) / p.delta_g
    kg = kg1 + kg2
    psi = p.Rg * kg                                            # MPK_g = psi*y/kg = Rg
    A = 1 / (kg ** psi * k ** p.alpha)
    i = p.delta * k
    # transfers close the steady-state budget (see README for derivation)
    rhs = p.tau_y + p.grants - p.rstar * p.dc - p.rk * p.dk - p.ig1 - p.ig2 - p.gc
    tr = (1 + p.tau_c) * (rhs + p.tau_c / (1 + p.tau_c) * ((1 - p.tau_y) - i - p.rstar * p.bp))
    cs = ((1 - p.tau_y) * (1 - p.phi * (1 - p.alpha)) + (1 - p.phi) * tr - i - p.rstar * p.bp) / ((1 - p.phi) * (1 + p.tau_c))
    ch = ((1 - p.tau_y) * (1 - p.alpha) + tr) / (1 + p.tau_c)
    ss = dict(y=1.0, cs=cs, ch=ch, i=i, k=k, q=1.0, bp=p.bp, rp=p.rstar, kg1=kg1, kg2=kg2,
              dc=p.dc, rc=p.rstar, tr=tr, tauc=p.tau_c)
    if tr < -0.5 or cs <= 0 or ch <= 0:
        raise ValueError(f"Infeasible calibration: tr={tr:.3f}, cs={cs:.3f}, ch={ch:.3f}")
    return ss, dict(beta=beta, psi=psi, A=A)


def exogenous_paths(p: Params):
    T = p.T
    t = np.arange(1, T + 1)
    rec = p.rec * p.rho_rec ** (t - 1)
    prog = np.zeros(T)
    if len(p.prog):
        n = min(len(p.prog), T)
        prog[:n] = p.prog[:n]
        last = p.prog[n - 1]
        for j in range(p.prog_taper):
            if n + j < T:
                prog[n + j] = last * (1 - (j + 1) / (p.prog_taper + 1))
    ig1 = p.ig1 + (1 - p.bbb) * rec + (1 - p.prog_bbb) * prog
    ig2 = p.ig2 + p.bbb * rec + p.prog_bbb * prog
    rec = rec + prog
    grants = p.grants + p.sh_grant * rec
    dk = np.empty(T)
    prev = p.dk
    for j in range(T):
        prev = p.dk + (prev - p.dk) * (1 - p.amort_k) + p.sh_conc * rec[j]
        dk[j] = prev
    return dict(rec=rec, prog=prog, ig1=ig1, ig2=ig2, grants=grants, dk=dk)


def initial_conditions(p: Params, ss):
    """Disaster hits at t=1: destroys stocks inherited from t=0 and lowers TFP."""
    share_a = ss["kg2"] / (ss["kg1"] + ss["kg2"])
    mitig = 1 - p.theta_a * share_a
    init = dict(ss)
    init["k"] = ss["k"] * (1 - p.eps_k * mitig)
    init["kg1"] = ss["kg1"] * (1 - p.eps_g)
    init["kg2"] = ss["kg2"] * (1 - p.eps_g * (1 - p.rho_res))
    a = np.array([-p.eps_a * mitig * p.rho_a ** j for j in range(p.T)])
    return init, a


def residuals(X, p, ss, cal, ex, init, a):
    """X: (T, NV). Returns (T, NV) residuals. Lag at t=0 from init, lead at T+1 = steady state."""
    T = p.T
    lag = np.vstack([[init[v] for v in VARS], X[:-1]])
    lead = np.vstack([X[1:], [ss[v] for v in VARS]])
    g = lambda M, v: M[:, IX[v]]
    y, cs, ch, i, k, q, bp, rp, kg1, kg2, dc, rc, tr, tauc = (X[:, j] for j in range(NV))
    k_l, kg1_l, kg2_l, bp_l, rp_l, dc_l, rc_l = (g(lag, v) for v in ["k", "kg1", "kg2", "bp", "rp", "dc", "rc"])
    y_f, cs_f, i_f, q_f, tauc_f = (g(lead, v) for v in ["y", "cs", "i", "q", "tauc"])
    dk = ex["dk"]; dk_l = np.concatenate([[p.dk], dk[:-1]])
    beta, psi, A = cal["beta"], cal["psi"], cal["A"]
    lam, lam_f = 1 / (cs * (1 + tauc)), 1 / (cs_f * (1 + tauc_f))
    x, x_f = i / k_l, i_f / k
    C = (1 - p.phi) * cs + p.phi * ch
    R = np.empty((T, NV))
    R[:, 0] = y - A * np.exp(a) * (kg1_l + kg2_l) ** psi * k_l ** p.alpha
    R[:, 1] = lam - beta * (1 + rp) * lam_f                                          # savers' bond Euler
    R[:, 2] = (1 + tauc) * ch - ((1 - p.tau_y) * (1 - p.alpha) * y + tr)              # hand-to-mouth
    R[:, 3] = q - (1 + p.kappa * (x - p.delta))                                       # Tobin's q
    R[:, 4] = k - ((1 - p.delta) * k_l + i)
    R[:, 5] = q - beta * lam_f / lam * ((1 - p.tau_y) * p.alpha * y_f / k + (1 - p.delta) * q_f
                                        - p.kappa / 2 * (x_f - p.delta) ** 2 + p.kappa * (x_f - p.delta) * x_f)
    R[:, 6] = ((1 - p.phi) * cs * (1 + tauc) + i + p.kappa / 2 * (x - p.delta) ** 2 * k_l + (1 + rp_l) * bp_l
               - (1 - p.tau_y) * (1 - p.phi * (1 - p.alpha)) * y - (1 - p.phi) * tr - bp)  # savers' budget
    R[:, 7] = rp - (p.rstar + p.eta_p * (np.exp(bp - p.bp) - 1))
    R[:, 8] = kg1 - ((1 - p.delta_g) * kg1_l + p.s * ex["ig1"])
    R[:, 9] = kg2 - ((1 - p.delta_g) * kg2_l + p.s * ex["ig2"] / (1 + p.mu_a))
    R[:, 10] = dc - ((1 + rc_l) * dc_l + (1 + p.rk) * dk_l - dk + ex["ig1"] + ex["ig2"] + p.gc + tr
                     - tauc * C - p.tau_y * y - ex["grants"])                         # government budget
    R[:, 11] = rc - (p.rstar + p.eta_g * (np.exp(dc - p.dc) - 1))
    R[:, 12] = tr - (ss["tr"] - p.lam_tr * (dc_l - p.dc))                             # fiscal rule
    R[:, 13] = tauc - (p.tau_c + p.lam_tau * (dc_l - p.dc))
    return R


def solve(p: Params = None, tol=1e-10, maxit=30, verbose=False):
    p = p or Params()
    ss, cal = steady_state(p)
    ex = exogenous_paths(p)
    init, a = initial_conditions(p, ss)
    X = np.tile([ss[v] for v in VARS], (p.T, 1))
    F = lambda Z: residuals(Z, p, ss, cal, ex, init, a)
    X = solve_path(F, X, tol, maxit, verbose)
    out = {v: X[:, IX[v]] for v in VARS}
    out.update(ex)
    out["a"] = a
    out["C"] = (1 - p.phi) * out["cs"] + p.phi * out["ch"]
    out["debt"] = out["dc"] + out["dk"]                                   # total public debt (level, ss GDP = 1)
    out["debt_gdp"] = out["debt"] / out["y"]
    out["revenue"] = out["tauc"] * out["C"] + p.tau_y * out["y"]
    out["prim_bal"] = out["revenue"] + out["grants"] - out["ig1"] - out["ig2"] - p.gc - out["tr"]
    return out, ss, cal


if __name__ == "__main__":
    out, ss, cal = solve(verbose=True)
    print("psi =", round(cal["psi"], 3), " steady state:", {k: round(v, 3) for k, v in ss.items()})
    C_ss = (1 - Params().phi) * ss["cs"] + Params().phi * ss["ch"]
    for t in [0, 1, 2, 4, 9, 19]:
        print(f"t={t+1:2d}  GDP {100*(out['y'][t]-1):6.2f}%  C {100*(out['C'][t]/C_ss-1):6.2f}%"
              f"  debt/GDP {100*out['debt_gdp'][t]:5.1f}  rc {100*out['rc'][t]:.2f}")
