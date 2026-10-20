# PriceFeasibility: power-sector carbon pricing under political-feasibility limits

Standalone metrics (no CPAT run needed) for comparing a **carbon tax / ETS**, a
**feebate** (equivalently output-based allocation or a tradable performance
standard) and **tax + feebate hybrids** in the power sector. The question it
answers: *what is the most abatement each country can get from pricing without
exceeding given limits on political cost?*

Status: v0.1, power sector only. Response functions are deliberately simple and
transparent; every non-data parameter is a prior in `PRIORS` and should be
replaced by country evidence or CPAT runs before results are quoted.

See [`NOTES.md`](NOTES.md) for what this tool is for and how it can be integrated
with CPAT and the other tools in this repo.

## Run

```bash
cd PriceFeasibility
python3 price_feasibility.py                         # default thresholds
python3 price_feasibility.py --f1 5 --f2 0.3 --f3 10 --rebate 0.1
python3 -m unittest test_price_feasibility
```

Needs Python 3 with `numpy` and `pandas`. Reads
`TaxMacroFiscalDashboard/build/output/global_data_wide.csv` and
`TaxMacroFiscalDashboard/build/raw_data/release_generation_yearly_global.csv`
(Ember). Writes `output/calibration_inputs.csv` (data used, one row per
country) and `output/country_metrics.csv` (results).

## Excel screening workbook

`PriceFeasibility_Screen.xlsx` is the same model as live formulas, in three tabs:

- **Screen:** limits, priors and grid steps as yellow input cells; a country
  picker; the best tax, feebate and hybrid for that country; and a tax-by-feebate
  grid of abatement, green where all three limits hold, with the best cell
  outlined.
- **Countries:** one row per country with every parameter and result.
- **Data:** long format with the same columns as the dashboard's
  `global_data_long.csv`. It holds only the indicators used, plus Ember
  generation and emissions by fuel as extra indicator codes (`gen_*_twh`,
  `em_*_mt`).

Rebuild and check with `python3 build_workbook.py --check`. The check needs
LibreOffice, recalculates the workbook and compares it with the Python tool.
At $50/t the two agree exactly. Under the limits they differ by up to 0.3pp,
because the workbook solves the tax-only and feebate-only cases exactly and
searches 11 tax steps for the hybrid, while Python searches a rate grid. The
workbook doesn't compute the hybrid shadow price.

## Method

Abatement from a carbon price has two parts:

- **Substitution** (cleaner generation): driven by the *marginal* price
  `tau_m = tau_tax + tau_feebate`.
  `A_sub = s_max * (1 - exp(-tau_m / tau_switch))`.
- **Demand** (less electricity used): driven only by the *average* tariff rise,
  which only the tax causes: `dp = phi * e_avg * tau_tax / p_res`,
  `A_dem = |elasticity| * dp`.

Abatement is `E = 1 - (1 - A_sub)(1 - A_dem)`, as % of baseline fossil power CO2.
A feebate gives the same substitution as a tax at the same rate, with no
average price rise. The share of a tax's abatement a feebate keeps is roughly
`s = A_sub / (A_sub + A_dem)`.

### Feasibility metrics (constraints)

| | Metric | Formula |
|---|---|---|
| **F1** visible price | First-year rise in residential tariff, % | `100 * dp / phase_in_years` |
| **F2** household loss | Net burden on bottom 40%, % of their consumption | `100 * b40 * dp * (1 - min(rebate / (phi * res_share * q40), 1))` |
| **F3** concentrated losers | Net payment by coal and oil generators, % of their revenue | `100 * [tau_tax * e_hi * (1 - phi) + tau_feebate * (e_hi - e_avg)] / p_wholesale` |

Notes:
- **F2:** fully protecting the bottom 40% takes only `phi * res_share * q40` of
  revenue, about 7% with the priors.
- **F3:** the first term is tax cost that is not passed through to tariffs and
  so is absorbed by generators or a state utility. The second term is the
  feebate's net transfer away from high emitters.

The optimiser searches tax and feebate rates from $0 to $200/t and reports, for
each instrument (tax only, feebate only, hybrid):
- the highest-abatement rates that meet all three limits;
- the binding constraint;
- the **shadow price of F1**: the extra abatement (pp) bought by allowing one
  more point of tariff rise.

### Output columns (`country_metrics.csv`)

- **Country parameters:** `baseline_mt`, `e_avg_t_mwh`, `pass_through`,
  `tau_switch`, `s_max_pct`, `hi_gen_share`.
- **Reference results at $50/t:** `E_tax50`, `F1_tax50`, `E_feebate50`,
  `F3_feebate50`, `s_at50`, `feebate_tax_ratio50`, `E_per_pp_F1_tax50`
  (abatement per point of tariff rise).
- **Constrained optimum** for each of tax, feebate and hybrid
  (`{tax,feebate,hybrid}_`): `E`, `tau_t`, `tau_f`, `F1`, `F2`, `F3`,
  `binding`.
- **Shadow prices and totals:** `tax_shadow_F1_pp`, `hybrid_shadow_F1_pp`,
  `hybrid_mt_saved`.

## Calibration indicators

Each parameter and threshold is listed with the indicator that calibrates it by
country, and whether that indicator is already used here.

### Abatement effects

| Parameter | Indicator | Status |
|---|---|---|
| Baseline emissions, `e_avg`, `e_hi`, `s_max` | Generation and emissions by fuel (Ember) | Used |
| `tau_switch` (coal-to-gas switching price) | Coal and gas fuel prices for power (IMF/CPAT price data in panel) × heat rates | Used. Gas price falls back to industry gas price; floor $15/t |
| `s_max` scaling (how fast clean supply can respond) | RISE renewable energy score | Used |
| `s_max`, `tau_switch` (cost of clean alternative) | Solar WACC (panel, ~20% coverage), renewable capacity growth (panel) | In panel, not yet used |
| `s_max` (existing gas headroom, plant age) | Global Energy Monitor plant trackers; IEA/IRENA LCOE | External |
| `phi` (tariff pass-through) | Power market structure (PSRT: wholesale / single buyer / integrated) | Used. Missing for most high-income countries (prior 0.9) |
| `phi` (administered prices) | Explicit fossil subsidies, % GDP (IMF) | In `calibration_inputs.csv`, not yet used |
| `elasticity` | CPAT elasticities; meta-analyses (e.g. Labandeira et al. 2017) | Prior (-0.3) |

### Feasibility thresholds (where the limits should sit)

| Threshold | Indicators | Status |
|---|---|---|
| F1 visible price | Explicit subsidies and retail price gaps (IMF); past fuel and tariff protest episodes (ACLED; IMF energy subsidy reform case studies); recent CPI inflation (WEO) | Subsidies in panel; rest external |
| F2 household loss | Bottom-40% electricity budget shares (household surveys, WB Global Consumption Database, CPAT distribution module); social assistance coverage (ASPIRE), which sets how cheaply rebates reach the bottom 40%; stated willingness to pay (Andre et al. 2024, Nature Climate Change: willingness to contribute 1% of income, 125 countries; Dechezleprêtre et al.: support for a carbon tax with cash transfers) | External; `b40` is a placeholder by income group |
| F3 concentrated losers | Coal rents % GDP and fuel exports % merchandise trade (panel); coal and oil share of generation (Ember); state ownership and age of coal fleet (GEM); coal employment | Panel and Ember items in `calibration_inputs.csv`; rest external |

## Priors to replace (in `PRIORS`)

| Prior | Value | Replace with |
|---|---|---|
| Heat rates | coal 10.0, gas 7.0 GJ/MWh | Country fleet efficiency |
| `kappa` | 0.4: max share of coal and oil output displaceable by 2030 | CPAT power module or capacity-build constraints |
| Replacement mix | 30% gas, 70% clean | Country pipeline |
| Pass-through | 1.0 / 0.6 / 0.4 by market structure | Regulatory tariff rules |
| Wholesale share | 0.6 of residential tariff | Wholesale prices |
| Residential share | 0.35 of electricity sales | Energy balances |
| `q40` | 0.2 of residential use by bottom 40% | Household surveys |
| `b40` | 2-4% by income group | Household surveys |

Default thresholds (F1 10%, F2 0.5%, F3 15%) are placeholders, not
recommendations.

## Caveats

- Power sector only. Road fuels (where feebates act slowly, through new
  vehicles) and industry are not covered.
- One-period response, broadly 2030; no dynamics or investment lags.
- Indirect household effects (through prices of other goods) are ignored in F2.
- Only direct fossil emissions in the baseline; Ember's lifecycle emissions
  from clean sources are excluded.
- A feebate's revenue neutrality assumes the benchmark is the average
  intensity, updated each year.
