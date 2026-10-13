# Projecting supply costs and retail energy prices (legacy CPAT method, as implemented in v0.10)

Method note for section 2 of `CPAT_Mitigation_CopyPaste_v0.10.xlsx`. It describes how legacy CPAT projects energy prices, cut down to the basics, and how the workbook implements it. It leaves out subsidy phase-outs, price liberalisation, existing carbon prices and ETS, and electricity prices (the power module covers those). v0.1 (method only) and v0.2 (v0.9, pure legacy rules) are in `Old/`. v0.10 changes two data rules by user decision (A10, A11); the changed cells are bright yellow with red text in `Inputs_prices`.

Sources: CPAT documentation, Mitigation chapter 3.2 (github.com/cpmodel/cpat_public); legacy workbook `CPAT 1.0pre_456_NoPropData.xlsb`, sheet `Mitigation` (rows 452/455 US deflator and CPI, 456/467/725 inflation indices, 633-695 international prices by source) and sheet `Prices_int` (regional assumptions); `cpat_coded/cpat_model/components/prices/` (`prices.py`, `domestic_prices.py`, `international_prices.py`), which ports the legacy formulas line by line.

## 0. Holistic assumptions: what the numbers mean

| # | Assumption |
|---|---|
| A1 | **Every price, tax and policy value in the model is in real US dollars of the results year** (MTInputs `ResultsYear`, 2026, read from scenario 1; Settings C8), per GJ. A constant real price is a constant number. There is no inflation after the conversion, and fuel use responds only to real price changes. |
| A2 | **Two indices, both = 1 in the results year.** `infl` (row 8) = US CPI(ResultsYear) / US CPI(year) converts nominal *domestic* values. `defl` (row 9) = US GDP deflator(ResultsYear) / deflator(year) converts nominal *international* prices. This is the legacy split: CPI rows 456/725, deflator row 467 (row 630 is its reciprocal). Up to 2030 they differ by up to 1.2% (2022: 1.137 vs 1.124; 2030: 0.918 vs 0.929), so one index would not reproduce legacy. Both are kept. |
| A3 | **Domestic price data are nominal USD of their own year** (IMF price dataset, `Prices_dom`: supply cost, excise and other taxes, retail price). Each is converted with `infl` of that year and divided by GJ per price unit. Units: $/liter for gasoline, diesel, LPG and kerosene; $/bbl for other oil products; $/GJ for coal, gas and biomass. |
| A4 | **Forecasting coefficients are base-year (2022) nominal values**: the margin over international prices and the domestic production cost. They are converted with `infl(2022)` and then held constant in real terms. (`cpat_coded` applies no conversion here; to check against legacy in bucket 5.) |
| A5 | **International prices are nominal by source** (IMF, WB, IEA, EIA, averages, Manual). They are converted with `defl`. Only their year-on-year ratio moves the supply cost, so their units and level drop out. Beyond a source's horizon the data hold them constant in real terms. |
| A6 | **MTInputs policy inputs are real (ResultsYear USD)**: the carbon price, feebate rates and fuel price reform increases (in price units, e.g. $/liter real). Only the carbon price has a nominal option: `NomorReal` = Nominal turns the path into nominal values, which are then multiplied by `infl` (legacy row 2249). |
| A7 | **History vs projection.** Prices up to the last historical price year (Settings C10, 2024) come from data. Later years are projected. Fuel use is data in the base year (2022) and modelled from 2023, using the data prices for 2023-2024. |
| A8 | **Global vs per scenario.** Global (Settings, MTInputs scenario 1): results year, government price controls (pass-through), last historical year, the country's gas market. Per scenario (its MTInputs column): international price source and High/Low adjustment, carbon price and its nominal/real switch, fuel price reform, feebates, coverage. |
| A9 | **Horizon.** Values after 2030 are not compared with legacy. The legacy CPI is held flat after 2031 (WEO horizon), the deflator is not, and some legacy rules (A10) misbehave over long horizons. The 2031-2035 columns compute, but they are not a target. |
| A10 | **Legacy rules are copied as they are, with one exception.** Kept: with pass-through 0.5 or 0.8, `txo` jumps in the first projected year (`cs_t` multiplies `cs_L` by (1 - pcc) a second time). **Changed (v0.10, user decision): other oil products are treated like the other oil products**, taking pass-through and margin from the dataset (Egypt: 0 and 7.95 $/bbl) instead of the legacy hardcodes 1 and 0. Under the legacy rule their retail price fell with the oil price to the 0.01 floor (v0.9: 3.79 $/GJ in 2024, 0.23 in 2030), and their fuel use rose 10x by 2030. |
| A11 | **VAT rate assumption (v0.10).** Use the dataset's VAT rate (`mit.vatrate`) where it is filled. Otherwise use the country's general VAT rate (`VAT_WEO`, Egypt 14%) for residential coal and gas and the all-sector oil products (gasoline, diesel, LPG, kerosene, other oil products), and 0 for power, industry (VAT is credited to firms) and biomass (largely informal). For Egypt the dataset's own retail prices imply exactly this: rp = (sp + txo) x 1.14 for residential gas and the oil products, and rp = sp + txo for power and industry. Effects: historical oil-product retail prices now equal the dataset `rp` (v0.9 was 14% lower), and new policies pay 14% VAT in those sectors. |

## 1. Dimensions and units

- Prices are kept by **fuel x price sector**, as in the IMF dataset. This gives 12 price fuels: coal and natural gas for power, residential and industry; gasoline, diesel, LPG, kerosene, other oil products and biomass for all sectors. Electricity is left to the power module. Subsectors take the price of their price sector (road and residential use the residential coal and gas prices).
- All in real $/GJ of the results year (A1-A3).

## 2. International prices `gp` (4 rows)

`gp_t = Prices_int[source|commodity]_t x defl_t x factor_t`

- Source: MTInputs `IntEnerPricForeSource` (row 220), with the asterisk dropped (`IMF-WB*` = average of IMF and WB).
- Commodity: crude oil (gasoline, diesel, LPG, kerosene, other oil products); coal; gas of the country's market (`PriceAssump`; Egypt *Global* = average of LNG, North America and Europe); a flat row of 1s (biomass).
- `factor`: MTInputs `IntEnerPricForecastAdjustment` (row 258). *Base* = 1; *High* / *Low* = 1.5 / 0.5 for oil and 1.25 / 0.75 for gas and coal, after the last historical year.

## 3. Supply cost `sp`

`sp_t = fixsp + (sp_{t-1} - fixsp) x gp_t / gp_{t-1}` for t > L; data for t <= L.

`fixsp = (margin + bucketed pass-through x production cost [coal, gas]) x infl(2022) / GJ per unit`; the margin of biomass is 0 (legacy hardcode; other oil products from data since v0.10). This is legacy `sp = fixsp + fltsp` with `fltsp_t = fltsp_{t-1} x gp_t / gp_{t-1}`. Producer subsidies (`ps`) are left out (0 for Egypt).

## 4. Excise and other taxes `txo` (pass-through)

- Pass-through `pcc`: raw = `mit.ps` of the base year (blank = 0; biomass = 1; other oil products from data since v0.10). Bucketed: <= 0.25 -> 0, <= 0.5 -> 0.5, else 1. Chosen per Settings C9 (`GovPriceControls`): *Bucketed* (default), *Manual* = 0.8, *None* = 1. Egypt bucketed: coal 1, gas 0, oil products incl. other oil products 0, biomass 1.
- Historical `txo` (as legacy):
  - coal, gas and biomass: `rp / (1 + vr) - sp`;
  - oil products: the dataset's `txo` (legacy builds their retail price as `sp + txo`, times `(1 + vr)`).
- Projected, with L = last historical year:
  - `txo_t = txo_L x pcc + cs_t`;
  - `cs_t = (sp_L - sp_t + txo_L(1 - pcc)) x (1 - pcc)`;
  - if `txo_L(1 - pcc) >= 0`, then `cs_t = max(txo_L(1 - pcc), cs_t)`.
- pcc = 1: supply-cost changes are passed on in full. pcc = 0: the retail price before new policies stays at its value in L, and the floating subsidy absorbs the change.

## 5. Retail price before new policies `rpb` and the after-tax price `atp`

`rpb = (sp + txo) x (1 + vr)`, where `vr` is the VAT rate of L (A11).

By subsector: `atp = max(rpb[price fuel] + nce x (1 + vr), 0.01)`. New policies (`nce` = new carbon tax + new excise) are passed through in full and pay VAT (legacy `vat = (sp + txo) x vr` with `nce` inside `txo`). The 0.01 floor is legacy's.

## 6. Layout in the workbook (section 2, per scenario group)

| Rows | Content | Hidden params (D:G) |
|---|---|---|
| 2 | `IntEnerPricForeSource`, `IntEnerPricForecastAdjustment` (MTInputs) | MTInputs row |
| 4 | `gp` oil, coal, gas, flat | commodity key, High factor, Low factor |
| 12 + 1 | `sp` by price fuel | fixsp, gp position, price position |
| 12 + 1 | `txo` | txo_L, sp_L, pcc, price position |
| 12 + 1 | `rpb` (visible when rolled up) | VAT rate |

Per subsector the variables are now `ctxnew, ntx, nce, atp, shp, ener`; `sp` and `tax` per subsector are gone. One formula per variable covers history and projection, base year included (`=IF(year <= L, data, projection)`). The data step is the `Inputs_prices` sheet (one row per price fuel); the 2035 column uses the LAMBDAs `SUPPLYCOST` and `OTHERTAX`.

## 7. Left out (later steps)

- Producer and consumer subsidy phase-outs, and the price-control phase-out (phase-out factors = 1).
- Existing carbon taxes and ETS (none for Egypt).
- The Europe shift to LNG.
- Ad-valorem baseline taxes (Egypt: Fixed).
- The oil-product share in the VAT of transport fuels (share = 1 assumed).
- Electricity prices.
