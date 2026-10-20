---
title: "Integrated country run: user guide"
subtitle: "User guide 4 · MacroFiscalModels v0.1"
---

# What it does

`integrated.py` runs one country, or all of them, through five tools in a fixed order, and writes results that the Tax and Macro-Fiscal Dashboard shows as Appendix D:

| Step | Tool | What it adds |
|---|---|---|
| 1 | Tax and Macro-Fiscal Dashboard (panel) | Data for the calibration: debt, taxes, revenue, primary balance, PEFA PI-11, IDS creditor terms, fossil rents, income group |
| 2 | National Mitigation Investment Model (NMIM) | Incremental budget need 2025-2050, split into **capital** and **spending** (non-CO2, land use, other energy CO2); private investment; energy savings; fossil power CO2 on the net zero path |
| 3 | PriceFeasibility | The most power-sector carbon pricing the feasibility limits F1-F3 allow (tax, feebate or hybrid), and which limit binds. The tax part times NMIM's power CO2 is revenue **earmarked for the budget need** |
| 4 | GIMF-lite, DIGNAD-lite or DIGNAR-lite | What the programme, financed this way, does to GDP and public debt |
| 5 | LIC-DSF-lite (low and lower-middle income only) | The mechanical risk of debt distress, before and with the programme |

CPAT is not part of the chain.

# Run it

```
cd MacroFiscalModels
python3 integrated.py KEN SLE UGA LBR      # print the chain for these countries
python3 integrated.py --all                # every NMIM country -> results/integrated_results.csv (~10 minutes)
```

Then rebuild the dashboard so Appendix D picks up the results:

```
python3 TaxMacroFiscalDashboard/build/scripts/build_global_indicators.py
python3 TaxMacroFiscalDashboard/build/scripts/build_dashboard.py
```

The order is stable: `integrated.py` reads only the panel's source indicators, never Appendix D. After a change to the raw data, run: build panel, `integrated.py --all`, build panel again, build workbook.

# Which model a country gets

| Rule (first that applies) | Model |
|---|---|
| Fossil rents at least 10% of GDP and not high income | DIGNAR-lite: DIGNAD-lite plus rents that fall 80% by 2050 on a global net zero path; government take 50% |
| Low income | DIGNAD-lite |
| Lower-middle income with a multilateral share of at least 50% | DIGNAD-lite |
| Otherwise | GIMF-lite |

GIMF-lite takes the whole programme (public capital, public spending, private investment, energy savings). DIGNAD-lite and DIGNAR-lite take the public part only. This is a known inconsistency: see "Caveats".

# Reading the output

`results/integrated_results.csv` has one row per country:

| Column | Meaning |
|---|---|
| `model`, `model_reason` | The model used and why |
| `assumptions` | How many shared settings are income-group defaults rather than dashboard data |
| `budget_need_2030_pct_gdp` | Incremental direct budget need (P1 + public gap of P4), % GDP; `budget_capital_...` and `budget_spending_...` split it |
| `private_2030_pct_gdp` | Incremental private investment |
| `pricing_tau_t`, `pricing_tau_f`, `pricing_binding` | Feasible tax and feebate rates (USD/t) and the binding limit (F1 tariff rise, F2 bottom-40% burden, F3 generators' payments) |
| `pricing_revenue_2030_pct_gdp`, `pricing_share_of_need_pct` | Revenue in 2030, and its share of the budget need over 2025-2050 |
| `financing_gap_2030_pct_gdp` | Budget need less pricing revenue |
| `status` | `ok`; `ok, fiscal rule tightened to ...` when the default rule could not stabilise debt; `not financeable under the fiscal rule` when even a tight rule could not |
| `gdp_dev_2030_pct`, `gdp_dev_2050_pct` | GDP against today's path (total GDP, including rents, for DIGNAR-lite) |
| `debt_2024_pct_gdp`, `debt_2030_pct_gdp`, `debt_2050_pct_gdp`, `debt_change_2050_pp` | Public debt |
| `dsf_rating_base`, `dsf_rating_programme`, `dsf_risk_code` | LIC-DSF-lite rating without and with the programme (code 1 low, 2 moderate, 3 high) |
| `dsf_assumptions` | Which DSF inputs are defaults |

In the dashboard, Appendix D shows six of these: budget need, pricing revenue, financing gap, GDP and debt in 2050, and the DSF risk code. The block is labelled as model output.

# Caveats

1. **Illustrative calibration.** The macro settings mix dashboard data and income-group defaults; `assumptions` counts the defaults.
2. **Energy savings only in GIMF-lite.** Countries routed to GIMF-lite get the savings channel, which can add several points of GDP by 2050; DIGNAD-lite and DIGNAR-lite countries do not. Compare GDP results across models with care (TODO item 1).
3. **Pricing revenue covers power only** and assumes the feasible rate is applied to NMIM's net zero emissions; the economic cost of the tariff rise is not in the macro models.
4. **DSF baseline** uses dashboard revenue and primary balance, NMIM growth and defaults for the external share of debt, exports (the WEO extract has no country trade data) and capacity (no CPIA). Ratings are mechanical.
5. **Fossil rents** fall 80% by 2050 in DIGNAR-lite whether or not the country runs the programme: it is the global transition, not a domestic choice.
