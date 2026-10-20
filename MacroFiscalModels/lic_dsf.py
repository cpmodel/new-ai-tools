"""LIC-DSF-lite: simplified Joint World Bank-IMF Debt Sustainability Framework for
Low-Income Countries (2018 framework; Guidance Note, Feb 2018).

Implements: debt-carrying capacity (Composite Indicator), external and public debt dynamics,
loan-cohort debt-service and PV (5% discount), the five debt-burden indicators, standardised
stress tests (+ natural-disaster tailored test) and the mechanical risk signal.

Coefficients, thresholds and shock sizes follow the 2018 Guidance Note as best recalled here;
check them against the current Guidance Note / LIC-DSF Excel template before official use.
Simplifications: everything in USD; PPG external debt only; domestic debt valued at face value
and rolled over with an effective interest rate; no realism tools; no judgement overlay.
"""
from dataclasses import dataclass, field, replace
import numpy as np

DISC = 0.05  # unified discount rate for PVs

# Composite Indicator: CI = sum(coef * variable); variables are 10-yr averages (5 hist + 5 proj)
CI_COEF = dict(cpia=0.385, growth=2.719, reserves=4.052, reserves_sq=-3.990,
               remittances=2.022, world_growth=13.520)
CI_CUTOFFS = (2.69, 3.05)  # weak < 2.69 <= medium <= 3.05 < strong

THRESHOLDS = {  # weak, medium, strong
    "PV ext debt / GDP": (30, 40, 55),
    "PV ext debt / exports": (140, 180, 240),
    "Ext debt service / exports": (10, 15, 21),
    "Ext debt service / revenue": (14, 18, 23),
    "PV public debt / GDP": (35, 55, 70),  # benchmark for total public debt
}
EXTERNAL = list(THRESHOLDS)[:4]


def composite_indicator(cpia, growth, reserves_cov, remittances, world_growth):
    """growth, remittances (% GDP), world_growth as decimals; reserves_cov = reserves / imports."""
    ci = (CI_COEF["cpia"] * cpia + CI_COEF["growth"] * growth + CI_COEF["reserves"] * reserves_cov
          + CI_COEF["reserves_sq"] * reserves_cov ** 2 + CI_COEF["remittances"] * remittances
          + CI_COEF["world_growth"] * world_growth)
    cls = "weak" if ci < CI_CUTOFFS[0] else ("medium" if ci <= CI_CUTOFFS[1] else "strong")
    return ci, cls


@dataclass
class Terms:
    share: float     # share of new external borrowing on these terms
    rate: float      # nominal USD interest rate
    maturity: int    # years
    grace: int       # years of grace on principal


@dataclass
class Inputs:
    N: int = 10                                     # projection years (signals use years 1..N)
    gdp0: float = 20.0                              # nominal GDP, USD bn, base year
    g: np.ndarray = None                            # real GDP growth, decimals, years 1..N
    defl: np.ndarray = None                         # USD GDP deflator growth
    exports0: float = 4.0                           # USD bn
    gx: np.ndarray = None                           # nominal USD export growth
    rev_gdp: np.ndarray = None                      # revenue excl. grants, % GDP (decimal)
    pd_gdp: np.ndarray = None                       # primary deficit incl. grants, % GDP
    ext_share: np.ndarray = None                    # share of financing needs met externally
    ext_stock0: float = 7.0                         # PPG external debt (face), USD bn
    ext_rate0: float = 0.015                        # avg interest on existing ext debt
    ext_mat0: int = 20                              # residual maturity of existing debt (straight-line)
    dom_stock0: float = 3.0                         # domestic debt, USD bn
    dom_rate: float = 0.08                          # nominal interest on domestic debt
    new_terms: list = field(default_factory=lambda: [Terms(0.6, 0.01, 40, 10), Terms(0.4, 0.06, 10, 3)])
    # historical standard deviations (10-yr) for stress tests
    sd_g: float = 0.02
    sd_pb: float = 0.02
    sd_gx: float = 0.10
    sd_of: float = 0.02                              # other flows (grants+FDI+remittances), % GDP

    def __post_init__(self):
        N = self.N
        d = lambda v: np.full(N, v, float)
        self.g = d(0.05) if self.g is None else np.asarray(self.g, float)
        self.defl = d(0.02) if self.defl is None else np.asarray(self.defl, float)
        self.gx = d(0.07) if self.gx is None else np.asarray(self.gx, float)
        self.rev_gdp = d(0.16) if self.rev_gdp is None else np.asarray(self.rev_gdp, float)
        self.pd_gdp = d(0.03) if self.pd_gdp is None else np.asarray(self.pd_gdp, float)
        self.ext_share = d(0.7) if self.ext_share is None else np.asarray(self.ext_share, float)


@dataclass
class Shock:
    name: str = "Baseline"
    g: float = 0.0            # change in real growth (decimal) in shock years
    pb: float = 0.0           # worsening of primary balance, % GDP
    gx: float = 0.0           # change in export growth
    of: float = 0.0           # lower other flows -> extra external borrowing, % GDP
    dep: float = 0.0          # one-off nominal depreciation (fall in USD value of LC aggregates)
    debt: float = 0.0         # one-off debt shock, % GDP (contingent liabilities / disaster)
    years: tuple = (2, 3)     # projection years in which flow shocks hit


def standard_shocks(inp: Inputs, disaster_prone=False):
    s = [
        Shock("Real GDP growth", g=-inp.sd_g),
        Shock("Primary balance", pb=inp.sd_pb),
        Shock("Exports", gx=-inp.sd_gx),
        Shock("Other flows", of=inp.sd_of),
        Shock("Depreciation (30%)", dep=0.30, years=(2,)),
        Shock("Combination", g=-0.5 * inp.sd_g, pb=0.5 * inp.sd_pb, gx=-0.5 * inp.sd_gx,
              of=0.5 * inp.sd_of, dep=0.15),
        Shock("Contingent liabilities", debt=0.07, years=(2,)),
    ]
    if disaster_prone:
        # tailored natural-disaster test: 10pp debt shock, growth -1.5pp, export growth -3.5pp
        s.append(Shock("Natural disaster", debt=0.10, g=-0.015, gx=-0.035, years=(2,)))
    return s


def _schedule(amount, terms: Terms, start, H):
    """Interest/principal of a loan contracted in year `start`, equal principal after grace."""
    interest, principal = np.zeros(H), np.zeros(H)
    n_rep = max(terms.maturity - terms.grace, 1)
    out = amount
    for y in range(start + 1, min(start + terms.maturity + 1, H)):
        interest[y] = out * terms.rate
        if y - start > terms.grace:
            p = amount / n_rep
            principal[y] = p
            out -= p
    return interest, principal


def simulate(inp: Inputs, shock: Shock = Shock(), H=80):
    N = inp.N
    hit = np.zeros(N + 1, bool)
    for y in shock.years:
        if y <= N:
            hit[y] = True
    yrs = np.arange(N + 1)
    gdp, X = np.zeros(N + 1), np.zeros(N + 1)
    gdp[0], X[0] = inp.gdp0, inp.exports0
    dep_year = shock.years[0]
    for t in range(1, N + 1):
        g = inp.g[t - 1] + (shock.g if hit[t] else 0)
        gdp[t] = gdp[t - 1] * (1 + g) * (1 + inp.defl[t - 1])
        if shock.dep and t == dep_year:
            gdp[t] *= 1 - shock.dep        # LC aggregates lose USD value; exports are USD-priced
        X[t] = X[t - 1] * (1 + inp.gx[t - 1] + (shock.gx if hit[t] else 0))
    # debt-service matrix by cohort; cohort 0 = existing debt (straight-line amortisation)
    INT, PRIN = [np.zeros(H)], [np.zeros(H)]
    rem = inp.ext_stock0
    for y in range(1, min(inp.ext_mat0, H - 1) + 1):
        PRIN[0][y] = inp.ext_stock0 / inp.ext_mat0
        INT[0][y] = rem * inp.ext_rate0
        rem -= PRIN[0][y]
    ext, dom = np.zeros(N + 1), np.zeros(N + 1)
    ext[0], dom[0] = inp.ext_stock0, inp.dom_stock0
    ds, rev, pv_ext = np.zeros(N + 1), np.zeros(N + 1), np.zeros(N + 1)
    disc = (1 + DISC) ** -np.arange(H)

    def pv_at(t):
        tot = sum(INT) + sum(PRIN)
        return float(tot[t + 1:] @ disc[1:H - t])

    pv_ext[0] = pv_at(0)
    dom_val = inp.dom_stock0
    for t in range(1, N + 1):
        if shock.dep and t == dep_year:
            dom_val *= 1 - shock.dep
        int_ext = sum(c[t] for c in INT)
        prin_ext = sum(c[t] for c in PRIN)
        ds[t] = int_ext + prin_ext
        rev[t] = inp.rev_gdp[t - 1] * gdp[t]
        pd = (inp.pd_gdp[t - 1] + (shock.pb if hit[t] else 0)) * gdp[t]
        oneoff = shock.debt * gdp[t] if (shock.debt and t == dep_year) else 0.0
        need = pd + ds[t] + inp.dom_rate * dom_val + oneoff     # domestic principal rolled over
        new_ext = inp.ext_share[t - 1] * need + (shock.of * gdp[t] if hit[t] else 0)
        dom_val = dom_val + need - inp.ext_share[t - 1] * need
        for tm in inp.new_terms:
            i_, p_ = _schedule(new_ext * tm.share, tm, t, H)
            INT.append(i_); PRIN.append(p_)
        ext[t] = ext[t - 1] + new_ext - prin_ext
        dom[t] = dom_val
        pv_ext[t] = pv_at(t)
    ind = {
        "PV ext debt / GDP": 100 * pv_ext / gdp,
        "PV ext debt / exports": 100 * pv_ext / X,
        "Ext debt service / exports": 100 * ds / X,
        "Ext debt service / revenue": 100 * np.where(rev > 0, ds / np.where(rev > 0, rev, 1), np.nan),
        "PV public debt / GDP": 100 * (pv_ext + dom) / gdp,
    }
    for k in ["Ext debt service / exports", "Ext debt service / revenue"]:
        ind[k][0] = np.nan  # flows not computed for base year
    return dict(years=yrs, gdp=gdp, exports=X, ext_face=ext, pv_ext=pv_ext, dom=dom, ds=ds, ind=ind)


def assess(inp: Inputs, ci_class="medium", disaster_prone=False):
    lvl = {"weak": 0, "medium": 1, "strong": 2}[ci_class]
    thr = {k: v[lvl] for k, v in THRESHOLDS.items()}
    base = simulate(inp)
    tests = {s.name: simulate(inp, s) for s in standard_shocks(inp, disaster_prone)}
    rows = {}
    for k in THRESHOLDS:
        b = np.nanmax(base["ind"][k][1:])
        worst_name, worst = max(((n, np.nanmax(r["ind"][k][1:])) for n, r in tests.items()), key=lambda z: z[1])
        rows[k] = dict(threshold=thr[k], baseline_max=b, stress_max=worst, worst_test=worst_name,
                       base_breach=b > thr[k], stress_breach=worst > thr[k])

    def rating(keys):
        if any(rows[k]["base_breach"] for k in keys):
            return "High"
        if any(rows[k]["stress_breach"] for k in keys):
            return "Moderate"
        return "Low"

    ext = rating(EXTERNAL)
    pub = rating(["PV public debt / GDP"])
    order = ["Low", "Moderate", "High"]
    overall = order[max(order.index(ext), order.index(pub))]
    return dict(thresholds=thr, rows=rows, external=ext, public=pub, overall=overall,
                base=base, tests=tests)


def print_assessment(a):
    print(f"{'Indicator':30s} {'Thr':>5s} {'Base max':>9s} {'Stress max':>10s}  Worst test")
    for k, r in a["rows"].items():
        flag = "BASE BREACH" if r["base_breach"] else ("stress breach" if r["stress_breach"] else "")
        print(f"{k:30s} {r['threshold']:5.0f} {r['baseline_max']:9.1f} {r['stress_max']:10.1f}  {r['worst_test']:24s} {flag}")
    print(f"Risk of external debt distress: {a['external']} | public: {a['public']} | overall: {a['overall']}")


if __name__ == "__main__":
    ci, cls = composite_indicator(cpia=3.3, growth=0.05, reserves_cov=0.30, remittances=0.04, world_growth=0.033)
    print(f"Composite indicator = {ci:.2f} -> {cls} debt-carrying capacity")
    print_assessment(assess(Inputs(), cls, disaster_prone=True))
