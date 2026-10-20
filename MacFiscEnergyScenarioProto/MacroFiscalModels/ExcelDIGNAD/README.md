# DIGNAD-R Excel (proof of concept, v0.3)

`DIGNAD_R_Excel_v0.3.xlsx` is DIGNAD-lite with the DIGNAR resource block written entirely in Excel
formulas, for 12 countries (Kenya, Sierra Leone, Uganda, Liberia, Ethiopia, Rwanda, Mozambique,
Ghana, Senegal, Angola, Congo Rep., Chad), 2024-2050. No macros, no Python to use it.

## Plan

| Step | Choice |
|---|---|
| Model form | **Recursive**: each year follows from the previous one, so plain formulas work. Python DIGNAD-lite solves the whole path at once (forward-looking savers); here private capital closes `p_lambda_k` of the gap to its desired level each year, and savers' consumption is not modelled. Supply side, budget, three kinds of financing, debt-elastic rate, fiscal rule, resource block and disaster follow DIGNAD-lite. |
| Inputs | **Dashboard norms.** `Data_wide`: one row per country, dashboard columns with the dashboard's codes, then an **augmented** block (shaded) with what the dashboard lacks, then `__year` columns. `Data_long`: annual programme paths (`iso3, indicator, year, value, source`). `Spec`: every input with block and source. |
| Assumptions | One sheet of named yellow cells `p_...` (as in the InvestmentNeeds workbook), incl. scenario switches. |
| Model | `Model`: a long panel, one row per country-year, baseline (B) and programme (P) side by side, all countries computed at once. `Calib`: per-country calibration in formulas. |
| Outputs | `Outputs`: results in the dashboard's long form, codes `dig.*`, so a master tool can read them like any `Data_long`. `Interface`: the contract (what the module needs, what it passes). |
| Check | Excel formulas equal a Python copy of the same equations (`--check`, to rounding); `Check` sheet compares with Python DIGNAD-lite / DIGNAR-lite. |

## Equations (per country, year t; non-resource GDP in 2024 = 1)

```
Disaster (optional)  dis_t = 1 in p_dis_year;  tfp_t = 1 - eps_a * rho_a^(t - dis_year) from then on
Production           Y_t = A * tfp_t * KGd^psi * Kd^alpha,  KGd = KG_{t-1}(1 - eps_g dis_t),  Kd = K_{t-1}(1 - eps_k dis_t)
Private capital      K*_t = alpha (1 - tax) Y_t / (r* + delta);   K_t = Kd + lambda_k (K*_t - Kd)
Public capital       KG_t = (1 - delta_g) KGd + s (ig + cap_t)            [cap_t: NMIM capital part, programme only]
Rents (DIGNAR)       R_t = R_0 (1 - rent_decline * min(1, (t - 2024)/26))  [only if rents >= p_res_threshold]
GDP                  GDP_t = Y_t + R_t
Programme            P_t = cap_t + spend_t                                 [spend_t: land use, non-CO2: no capital]
Revenue              tax Y_t + gov_take R_t + grants_0 + sh_grant P_t + pricing_t
Fiscal rule          tr_t = tr_bar - lambda_f (D_{t-1}/GDP_{t-1} - d_0)
Primary deficit      g + ig + P_t + tr_t - revenue_t
Debt                 D_t = D_{t-1} + rk Dk_{t-1} + rc_{t-1}(D_{t-1} - Dk_{t-1}) + primary deficit_t
Concessional debt    Dk_t = Dk_0 + (Dk_{t-1} - Dk_0)(1 - amort) + sh_conc P_t
Commercial rate      rc_t = r* + eta * max(0, D_t/GDP_t - d_0)
Hand-to-mouth        c_t = (1 - tax)(1 - alpha) Y_t + tr_t
Calibration          K_0 = alpha(1 - tax)/(r* + delta), KG_0 = s ig/delta_g, psi = Rg KG_0, A sets Y_0 = 1,
                     g = max(5%, revenue_0 - ig - interest_0 - 2%), tr_bar balances the 2024 budget
```

Efficiency `s` from PEFA PI-11 (as in InvestmentNeeds v1.1), the commercial rate from IDS, the
concessional share from the IDS multilateral share, rents from WDI: the same rules as
`calibration.py`.

## Results it passes (Outputs, long form)

| Code | Years |
|---|---|
| `dig.gdp_dev_programme_pct` | 2030, 2050 |
| `dig.gdp_vs_2024_programme_pct`, `dig.gdp_vs_2024_baseline_pct` | 2050 |
| `dig.debt_programme_pct_gdp` | 2030, 2050 |
| `dig.debt_baseline_pct_gdp`, `dig.debt_due_to_programme_pp` | 2050 |
| `dig.htm_consumption_dev_programme_pct`, `dig.transfers_change_pp_gdp`, `dig.rents_pct_gdp` | 2050 |

A master tool reads `Outputs` exactly like the dashboard's `Data_long`: `iso3, country, indicator,
year, value, source`.

## Check

| | GDP vs 2024, 2050 (Python / Excel) | Debt 2050 (Python / Excel) |
|---|---|---|
| Kenya | 0.26 / 0.25 | 74.6 / 75.0 |
| Uganda | 0.41 / 0.41 | 65.7 / 66.6 |
| Ethiopia | 0.74 / 0.80 | 48.2 / 48.4 |
| Liberia (rule tightened in Python) | 0.56 / 0.73 | 100.0 / 72.9 |
| Angola (DIGNAR) | -20.5 / -22.8 | 82.1 / 108.8 |

Close where the fiscal rule did not need tightening; further apart where it did, because Python
lets both the consumption tax and transfers respond and solves forward. Ghana and Senegal used
GIMF-lite in the integrated run, so they are not compared.

## Rebuild

```bash
python3 build_dignad_excel.py --check     # needs results/integrated_{results,paths}.csv and the dashboard panel
```

## Caveats

Illustrative calibration. Public investment now comes from the dashboard's Appendix C (IMF ICSD,
to 2019) where it has a value, else the income-group default; three augmented inputs are still
income-group defaults (return on public capital, hand-to-mouth share, grants: no data source yet);
`fiscal_rule_speed` is passed in from the integrated run. No energy savings, no private investment
push, no exchange rate; one fiscal instrument. Checked in LibreOffice, not in Excel itself.
