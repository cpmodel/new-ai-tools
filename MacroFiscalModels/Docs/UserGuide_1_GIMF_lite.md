---
title: "GIMF-lite user guide"
subtitle: "User guide 1 of 3 · MacroFiscalModels v0.1"
---

# What it is

GIMF-lite is a one-country, annual simplification of the IMF's Global Integrated Monetary and Fiscal Model (Kumhof, Laxton, Muir and Mursula 2010). It is built to show what a large investment programme does to the economy and the budget, and why the financing choice matters.

Its key feature is that **households are not Ricardian**: most plan over a finite horizon (about 20 years) and expect their labour income to decline, so they treat government debt partly as wealth; a quarter are liquidity-constrained and spend their current income. Prices and wages are sticky, the central bank follows a Taylor rule, and the real exchange rate and current account adjust.

**Use it for:** the growth, consumption, inflation, exchange-rate and debt effects of a public and private investment push; comparing debt and tax financing; testing how much depends on energy savings or on investment efficiency.

**Do not use it for:** short-term forecasting, commodity-price shocks, spillovers between countries, or carbon pricing (use CPAT).

# Running it

```python
from nmim_bridge import programme, run_gimf
from gimf_lite import Params, solve

# with the NMIM programme
out, ss0, ssT, cal = run_gimf(programme("KEN"), Params())

# with your own paths (% of GDP, one value per year from 2025; the last value is held)
out, ss0, ssT, cal = solve(Params(), ig_extra=[0.01] * 10, ix=[0.02] * 10, sav=[0.0])
```

`run_gimf` takes switches `public`, `private` and `savings` (all `True` by default) to switch each part of the programme off.

# Parameters

Change them with `dataclasses.replace(Params(), name=value)`. All shares are of GDP; rates are annual.

## Households and preferences

| Parameter | Default | Meaning | Where to get it |
|---|---|---|---|
| `theta` | 0.95 | Survival probability; planning horizon 1/(1 − θ) = 20 years | GIMF papers (Kumhof et al. 2010; Anderson et al. 2013) |
| `chi` | 0.97 | Annual decline of lifetime labour income | GIMF papers |
| `psi` | 0.25 | Share of liquidity-constrained households | Findex (no account, no borrowing), household surveys; often 0.4-0.5 in LICs |

## Production and investment

| Parameter | Default | Meaning | Where to get it |
|---|---|---|---|
| `alpha` | 0.33 | Private capital share | National accounts (1 − labour share) |
| `Rg` | 0.15 | Gross return on public capital (sets its output elasticity) | IMF public investment literature: 10-25% |
| `delta`, `delta_g` | 0.08, 0.04 | Depreciation, private and public | Penn World Table; IMF Investment and Capital Stock dataset |
| `s` | 0.7 | Public investment efficiency: share of spending that becomes capital | PIMA, PIE-X efficiency scores |
| `kappa` | 4.0 | Investment adjustment cost | Literature; 2-10 |
| `zeta` | 0.0 | Output from each unit of green private capital (0 = replacement only) | Judgement; 0 is conservative |

## Prices, wages and monetary policy

| Parameter | Default | Meaning |
|---|---|---|
| `pibar` | 0.03 | Inflation target |
| `mu` | 1.10 | Price markup |
| `kappa_p` | 0.25 | Slope of the Phillips curve (annual) |
| `rho_w` | 0.5 | Wage stickiness (0 = flexible) |
| `wage_el` | 1.0 | Response of the target real wage to employment |
| `phi_pi`, `rho_i` | 1.5, 0.5 | Taylor rule: response to inflation, smoothing |

## External sector

| Parameter | Default | Meaning | Where to get it |
|---|---|---|---|
| `rstar` | 0.03 | World real interest rate | |
| `nfd` | 0.30 | Net foreign debt, % GDP | IMF IIP data (net international investment position) |
| `eta_f` | 0.02 | Country premium per unit of net foreign debt / GDP | Spread regressions; 0.01-0.05 |
| `m_share` | 0.30 | Imports / absorption | National accounts |
| `eps_x`, `eps_m` | 1.0, 0.5 | Export and import elasticities to the real exchange rate | Literature |

## Fiscal

| Parameter | Default | Meaning |
|---|---|---|
| `g`, `ig` | 0.12, 0.04 | Government consumption and public investment, % GDP (before the programme) |
| `tau_l`, `tau_c`, `tau_k` | 0.15, 0.10, 0.20 | Effective labour, consumption and capital tax rates |
| `b` | 0.50 | Public debt target, % GDP |
| `lam_b` | 0.05 | Speed at which the fiscal rule returns debt to target |
| `instrument` | `"tr"` | What closes the fiscal rule: `"tr"` transfers, `"tau_l"` labour tax |
| `dfs` | 1.0 | Share of the extra public investment the rule lets be debt financed (0 = fully tax financed) |

# Outputs

`out` is a dictionary of annual paths, index 0 = 2025:

| Key | Meaning |
|---|---|
| `y`, `c`, `inv` | GDP, consumption, private investment (excluding green), levels with initial GDP = 1 |
| `co`, `cl` | Consumption of OLG and liquidity-constrained households |
| `l`, `w` | Employment, real wage |
| `pi`, `i`, `r` | Inflation, nominal policy rate, real rate |
| `q` | Real exchange rate (up = depreciation) |
| `debt_gdp`, `nfd_gdp` | Public debt and net foreign debt, % of GDP |
| `revenue`, `prim_def` | Revenue and primary deficit, % of initial GDP |
| `tr`, `tau_l` | Transfers and labour tax rate (one of them moves with the rule) |
| `kg`, `kx` | Public capital and green private capital |
| `hh_budget_gap` | Check: should be about 0 (1e-15) |

Deviations: `100 * (out["y"] / ss0["y"] - 1)` gives GDP in % deviation from the no-programme path.

# Recipes

| Question | How |
|---|---|
| Debt versus tax financing | Compare `Params(dfs=1.0)` with `Params(dfs=0.0, instrument="tau_l")` |
| How much depends on energy savings? | `run_gimf(prog, p, savings=False)` |
| Weak public investment management | `Params(s=0.4)` against `s=0.7` |
| More credit-constrained households | `Params(psi=0.45)` |
| Ricardian benchmark | `Params(theta=0.9999, chi=1.0, psi=0.0)`: debt and tax financing then give the same consumption |
| Fixed-rate-like monetary response | Raise `phi_pi` (e.g. 3.0) and `rho_i` (0.8) |

# Reading the results

- **Short run:** the programme raises demand; with sticky prices output rises, inflation rises a little and the policy rate responds. Consumption can fall on impact because the rule and higher real rates reduce wealth.
- **Long run:** public capital (through `s` and `Rg`) and energy savings raise output; debt-financed programmes leave higher debt and net foreign debt and so a higher country premium.
- **Financing:** with non-Ricardian households, debt financing crowds out more in the long run; tax financing keeps debt lower for almost the same growth.

# Limits

One country and one good; no oil or commodity block, financial accelerator or zero lower bound; linear Phillips curve; energy savings enter as a value-added gain (generous for imported fuel); the programme is held at its 2050 level forever. See the method note, section 6.
