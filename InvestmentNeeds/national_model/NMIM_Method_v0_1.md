# NMIM: a national model of net zero mitigation investment (method, v0.1)

NMIM estimates, for each of 186 countries and each year from 2024 to 2050, the investment a net zero path needs, the investment a current-trends baseline would make anyway, and the incremental cost of the difference under three cost-of-capital lenses. It is a simplified stock-and-flow model built to CPAT norms: CPAT's own data, codes and activity equation, a CPAT-layout workbook in live formulas, and a Python engine that evaluates the same equations for every country.

World results, 2030 (USD bn a year, 2024 prices):

| | Advanced economies | China | Other EMDCs | World |
|---|---|---|---|---|
| Investment, baseline | 1,028 | 705 | 351 | 2,084 |
| Investment, net zero | 2,273 | 1,526 | 1,303 | 5,103 |
| Incremental investment | 1,245 | 822 | 952 | 3,018 |
| Incremental cost (a) market WACC | 104 | 152 | 436 | 693 |
| Incremental cost (b) World Bank 6% | 145 | 207 | 339 | 690 |
| Incremental cost (c) 6% + local co-benefits | 72 | -105 | 199 | 166 |
| GHG, baseline / net zero (MtCO2e) | 13,655 / 10,345 | 16,650 / 13,658 | 27,851 / 20,265 | 58,156 / 44,268 |

Incremental cost in a year is the annual cost of all incremental capital then in service, so it rises as the net zero capital stock builds: USD 1.2 tn in 2040 and 1.8-2.0 tn in 2050 at (a) and (b), while lens (c) turns to a net saving of USD 0.7 tn in 2040 and 1.0 tn in 2050. Present value 2025-2050: USD 10.4 tn at (a), 11.3 tn at (b), -3.3 tn at (c). Cumulative abatement is 710 GtCO2e.

## 1. Structure

| Item | Choice |
|---|---|
| Countries | 186 with IMF CPAT 2024 data (fossil fuel subsidies panel); groups as IHLEG (41 IMF advanced economies, China, other EMDCs) |
| Years | 2024 base (data), 2025-2050 projected; workbook history columns 2021-2023 for data rows |
| Scenarios | 1 = baseline (current trends, real prices fixed at 2024); 2 = net zero |
| Units | PJ, TWh, GW, USD bn (2024 prices), MtCO2e |
| Codes | CPAT: `mit.<indicator>.<fuel>.<sector>.<scenario>`; results `<iso>.mit...` |

One equation specification (`nmim_v0_1.py`: `PARAMS`, `SCALARS`, `ROWS`) is evaluated in Python for every country and rendered as Excel formulas for the country chosen on the workbook's Settings sheet. Every same-year reference points to a row computed earlier, which the engine asserts, so the two implementations follow the same order. The check script confirms they agree to 1e-9 for eight countries.

## 2. Data

| Data | Source | Use |
|---|---|---|
| End-use energy by fuel and sector, 2021-2024 | IMF CPAT 2024 panel (`mit.con.*`; oil products split with CPAT's transport and residential portions) | Base year and history |
| Supply costs and air-pollution costs per unit | IMF CPAT 2024 panel (`mit.sup.cost.*`, `mit.airpol.cost.*`), converted to USD/GJ with CPAT fuel units | Fuel bills, co-benefits |
| Energy CO2, GHG with and without land use | IMF CPAT 2024 panel | Calibration, non-CO2 and land use |
| Income elasticities and autonomous efficiency improvement | CPAT elasticities by income group (CPAT-AI-Mitigation-MVP data, from legacy CPAT) | Baseline activity |
| Generation and capacity by source | Ember yearly electricity data (latest year) | Power base year, capacity factors, fossil trend |
| GDP, PPP GDP, growth 2025-2026 | IMF WEO | Activity, price level |
| CBAM goods output, growth, process emission factors | CBAM Policy Tool global v1.0, Country data (57 countries) | Industrial process CO2 |
| Market WACC | Mitigation Investment Needs v1.1 (observed IRENA/Steffen, Damodaran build, else IHLEG group rate) | Lens (a) |

## 3. Equations

**Macro.** Real growth is the WEO forecast in 2025-2026, then converges linearly to a long-run rate by CPAT income group by 2040 (LIC 4.5%, LMIC 4%, UMIC 3%, HIC 1.5%).

**Baseline activity (CPAT).** For each of 11 end-use rows (road gasoline and diesel; residential coal, gas, LPG, kerosene, electricity; industrial coal, gas, oil, electricity):

```
A_t = A_{t-1} x (1 + g_t)^inc / (1 + aei)
```

This is CPAT's energy-use equation with prices held at their 2024 real level, so the price terms drop out. Road activity is in ICE-equivalent fuel: base-year road fuel divided by (1 - EV stock share).

**Road transport.** Vehicles (car-equivalent) = road activity / 38.5 GJ a year. Sales = stock / 12 + growth. The EV share of sales rises from its 2024 level (AE 15%, China 40%, EMDC 3%) to a baseline 2035 level (50% / 80% / 15%) or, under net zero, to 100% by 2035 (AE, China) or 2040 (EMDC). EV stock share follows a 12-year turnover. Fossil road fuel = activity x (1 - EV stock share); EV electricity = activity x EV stock share x 0.25. Investment = EV sales x (premium + USD 1,000 charger) x price factor; the premium falls from USD 10,000 in 2024 to 0 in 2032. EVs save USD 0.012/km in maintenance.

**Buildings.** Under net zero, the share of fossil heat fuel replaced rises linearly to 100% by 2045 (AE) or 2050; heat pumps use 0.33 units of electricity per unit of fuel. Investment = USD 200 per GJ a year newly displaced x price factor.

**Industry.** Under net zero, industrial fossil use is abated linearly to 90% (AE, China) or 80% (EMDC) of baseline by 2050; 60% of the abated fuel is electrified at 0.8 units of electricity per unit. Investment = USD 150 per GJ a year newly abated. Process CO2 of the CBAM goods (CBAM tool output x process factors, growing at the CBAM tool's rate to 2030) is abated to 90% by 2050 at USD 800 per tonne a year of capacity.

**Power.** Generation = end-use electricity x the base-year ratio of Ember generation to CPAT consumption (0.9-1.4). Clean firm generation (hydro, nuclear, bioenergy, other renewables) is held at 2024. The fossil share falls on its recent Ember trend (capped at 2 points a year) in the baseline, and linearly to zero by 2035 / 2040 / 2045 under net zero. Variable renewables fill the rest, 60% solar. Capacity = generation / capacity factor (Ember-derived per country, else 0.17 / 0.35); additions = growth + replacement over 25 years. Fossil capacity is a stock with 30-year lives; new fossil build is gas at USD 1,304/kW (CPAT PowerTech). Solar and wind cost USD 760 and 1,160/kW in 2024 (IRENA), falling 3% and 1.5% a year. Grids add 0.6 USD per USD of generation investment (IEA); storage adds 2 kWh per kW of new VRE at full VRE share, phased in from a 20% share, at USD 250/kWh falling 3% a year. Power fuel uses country heat rates calibrated to CPAT power fuel use and Ember generation (coal 7-14, gas 5.5-12 GJ/MWh).

**Calibration to current investment.** Clean investment that the rows do not build bottom-up (hydro, nuclear, efficiency beyond the rows, other) is added to both scenarios. Its 2024 level by group is the gap between IEA World Energy Investment 2024 (China USD 675 bn, EMDEs other than China ~320 bn, advanced economies the residual of ~2,000 bn, ~1,005 bn) and the model's baseline in 2025 (gaps: AE 542, China 90, EMDC 105). It is allocated to countries by 2024 GDP and grows with GDP. It sets the levels of both scenarios but leaves incremental investment and incremental cost unchanged.

**Non-CO2, process and land use.** Non-CO2 and other process GHG (CPAT GHG excluding land use, less energy and CBAM process CO2) are cut 30% by 2030 (Global Methane Pledge) and 60% by 2050 at USD 50/t; net land-use emissions fall to zero by 2030 (Glasgow forests declaration) at USD 50/t. Energy CO2 in CPAT that the 11 rows and power do not cover (aviation, shipping, other uses) is held flat in the baseline and abated with the industry path at USD 100/t. These three are annual spending, not capital.

**Emissions, fuel bills, air pollution.** Fossil use x IPCC 2006 CO2 factors; x CPAT 2024 supply costs; x CPAT air-pollution costs per unit.

**Incremental cost.** For each capital sector, incremental capital in service K_t = K_{t-1} + dI_t - dI_{t-life}. Then:

```
IC(a)_t = sum_s K_s,t x CRF(WACC_c, life_s) + spending_t + dO&M_t - fuel savings_t
IC(b)_t = same with CRF(6%, life_s)
IC(c)_t = IC(b)_t - air pollution avoided_t
```

Lives: generation 25, grids 40, vehicles 12, buildings 25, industry 25 years. Present values discount 2025-2050 at the lens rate (WACC for a, 6% for b and c).

## 4. Validation

| Check | Result |
|---|---|
| Workbook = Python (key rows, 2024-2050) | Relative difference < 1e-9 for IND, USA, CHN, BRA, KEN, EGY, DEU, IDN |
| Identities | Generation balance, non-negative capacity and investment, sector sum, net zero below baseline: all pass |
| Net zero investment 2030 vs IHLEG/IEA | USD 5.10 tn vs 4.8 tn clean energy + ~0.4 tn nature |
| Baseline investment 2030 vs IEA 2024 | USD 2.08 tn vs ~2.0 tn (calibrated in 2025, then grows) |
| EMDCs other than China, net zero 2030 | USD 1.30 tn vs IHLEG 1.6 tn clean energy + 0.35 tn nature |
| Baseline GHG 2030 | 58.2 Gt vs ~53-57 Gt for current policies |
| Energy CO2 2024 | 11 rows + power = 33.6 Gt; CPAT 39.0 Gt; the gap is carried as other energy CO2, so totals tie to CPAT |

The full report is `check_report_v0_1.md`.

## 5. Comparison with the v1.1 allocation

NMIM's 2030 net zero investment matches v1.1's at world level (USD 5.10 tn vs 5.20 tn) but distributes it differently: China needs more (1.53 vs 0.86 tn) because its coal-heavy power must be replaced by 2040; advanced economies (2.27 vs 2.39 tn) and other EMDCs (1.30 vs 1.95 tn) need less. Incremental investment is a larger share of the total (59% vs 24-38%): it is measured against a current-trends baseline that invests USD 2.1 tn, not against the CCDR or McKinsey definitions behind v1.1's shares.

## 6. Open issues

- **Baseline composition.** Baseline levels are calibrated to IEA 2024 investment, but the bottom-up baseline still has no buildings or industry spending, so incremental investment in those sectors may be overstated; the calibration term holds that spending in both scenarios.
- **No price response.** CPAT's price and efficiency elasticities are off (real prices fixed); carbon pricing is not modelled.
- **Power is not optimised.** The mix follows target dates, not least cost; the cost of capital does not change it.
- **Car-equivalent vehicles.** Trucks and buses are scaled to cars per GJ of fuel.
- **Assumed transformation dates and unit costs** (marked † on the Inputs sheet), especially buildings, industry and the EV premium.
- **Residual emissions** of 10 Gt in 2050 (non-CO2, process, other energy) need removals or deeper abatement, which are not modelled.
- **Formula length.** Several workbook formulas exceed the 100-character norm; they should be split into helper rows.
- **Damodaran data** are still not loaded in the v1.1 WACC.

## 7. Files

| File | Role |
|---|---|
| `extract_data_v0_1.py` | Data extraction to `data/*.csv` |
| `nmim_v0_1.py` | Equation specification and Python engine; writes `results/` |
| `build_workbook_v0_1.py` | Builds `NMIM_v0.1.xlsx` (CPAT layout) |
| `check_v0_1.py` | LibreOffice checks; writes `check_report_v0_1.md` |
| `PLAN_v0_1.md` | Build plan |
