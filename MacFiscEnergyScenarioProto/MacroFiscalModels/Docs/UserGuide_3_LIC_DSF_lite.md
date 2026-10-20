---
title: "LIC-DSF-lite user guide"
subtitle: "User guide 3 of 3 · MacroFiscalModels v0.1"
---

# What it is

LIC-DSF-lite reproduces the mechanics of the joint World Bank-IMF Debt Sustainability Framework for Low-Income Countries (2018): the country's debt-carrying capacity, the debt-burden indicators and their thresholds, the standard stress tests and the mechanical risk signal.

**It is not a DSA.** It has no realism tools, no judgement, no small-breach rule and no market-financing module; its coefficients, thresholds and shock sizes are as recalled from the 2018 Guidance Note and **must be checked against the current Guidance Note and template before any use outside exploration.**

**Use it for:** seeing how an investment programme or a macro shock moves the debt indicators against the thresholds; teaching the framework; quick sensitivity work around a country's published DSA.

# Running it

```python
from lic_dsf import Inputs, Terms, composite_indicator, assess, print_assessment

ci, cls = composite_indicator(cpia=3.3, growth=0.05, reserves_cov=0.30,
                              remittances=0.04, world_growth=0.033)
inp = Inputs(gdp0=20.0, exports0=4.0, ext_stock0=7.0, dom_stock0=3.0)   # USD bn
a = assess(inp, cls, disaster_prone=True)
print_assessment(a)
```

With a macro model: `dsf_overlay(inp, out["y"], out["prim_def"], prim_def_ss)` from `nmim_bridge` returns new `Inputs` shifted by the model's change in growth and primary deficit (see `run_demo.py`, section C).

# Step 1: debt-carrying capacity

`composite_indicator` returns the Composite Indicator (CI) and the capacity class:

CI = 0.385·CPIA + 2.719·growth + 4.052·(reserves/imports) − 3.990·(reserves/imports)² + 2.022·(remittances/GDP) + 13.520·(world growth)

Weak if CI < 2.69, medium from 2.69 to 3.05, strong above 3.05. Each input is a 10-year average (5 years of history and 5 of projection), in decimals (5% = 0.05); CPIA is on its 1-6 scale; reserves are in proportion of imports.

| Input | Source |
|---|---|
| `cpia` | World Bank CPIA |
| `growth` | WEO, country DSA |
| `reserves_cov` | IMF IFS reserves / WEO imports |
| `remittances` | World Bank KNOMAD |
| `world_growth` | WEO |

# Step 2: inputs

`Inputs` (all USD bn or decimals; arrays have one value per projection year, `N` = 10 years):

| Field | Default | Meaning |
|---|---|---|
| `gdp0`, `exports0` | 20, 4 | Nominal GDP and exports, base year |
| `g`, `defl`, `gx` | 5%, 2%, 7% | Real growth, USD GDP deflator growth, nominal USD export growth |
| `rev_gdp`, `pd_gdp` | 16%, 3% | Revenue excluding grants and primary deficit, % GDP |
| `ext_share` | 70% | Share of financing needs met by external borrowing |
| `ext_stock0`, `ext_rate0`, `ext_mat0` | 7, 1.5%, 20 | Existing PPG external debt, its interest rate and residual maturity (straight-line repayment) |
| `dom_stock0`, `dom_rate` | 3, 8% | Domestic debt and its interest rate (rolled over) |
| `new_terms` | 60% at 1%, 40 years, 10 grace; 40% at 6%, 10 years, 3 grace | Terms of new external borrowing (`Terms(share, rate, maturity, grace)`) |
| `sd_g`, `sd_pb`, `sd_gx`, `sd_of` | 2%, 2%, 10%, 2% | Ten-year historical standard deviations used in the stress tests |

Replace the defaults with the country's DSA baseline: the macro framework, the loan-by-loan debt service of existing debt (approximated here by `ext_rate0` and `ext_mat0`) and the financing assumptions.

# Step 3: indicators and thresholds

| Indicator | Weak | Medium | Strong |
|---|---|---|---|
| PV of external debt / GDP | 30 | 40 | 55 |
| PV of external debt / exports | 140 | 180 | 240 |
| External debt service / exports | 10 | 15 | 21 |
| External debt service / revenue | 14 | 18 | 23 |
| PV of total public debt / GDP (benchmark) | 35 | 55 | 70 |

Present values discount future debt service at 5%. New loans are tracked by cohort with their own repayment schedules.

# Step 4: stress tests

| Test | Shock (years 2-3 unless stated) |
|---|---|
| Real GDP growth | growth − 1 SD |
| Primary balance | primary deficit + 1 SD |
| Exports | export growth − 1 SD |
| Other flows | extra external borrowing of 1 SD of other flows |
| Depreciation | 30% one-off in year 2 |
| Combination | half of each of the four above + 15% depreciation |
| Contingent liabilities | one-off debt of 7% of GDP in year 2 |
| Natural disaster (`disaster_prone=True`) | 10% of GDP debt, growth − 1.5 pp, export growth − 3.5 pp in year 2 |

# Step 5: the rating

`assess` returns, for each indicator, the maximum over the 10 years in the baseline and under the most extreme test. Then:

- a baseline breach of any external threshold → **High** risk of external debt distress;
- a breach only under stress → **Moderate**;
- no breach → **Low**.

The same rule on the public debt benchmark gives the public signal; the overall rating is the worse of the two. `a["base"]` and `a["tests"]` hold the full paths for charts.

# Recipes

| Question | How |
|---|---|
| Effect of an investment programme | `dsf_overlay` with GIMF-lite or DIGNAD-lite paths, then `assess` |
| More concessional borrowing | Raise the concessional share in `new_terms` |
| Weaker capacity | `assess(inp, "weak")` |
| Bigger historical volatility | Raise `sd_g`, `sd_gx` |

# Limits

USD only; PPG external debt only; domestic debt at face value; existing debt service approximated; no realism tools, judgement or small-breach rule; coefficients and thresholds to be checked against the current Guidance Note.
