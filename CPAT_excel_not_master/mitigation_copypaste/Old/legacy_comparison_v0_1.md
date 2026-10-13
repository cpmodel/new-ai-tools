# Legacy CPAT cached baseline vs CPAT-AI-Mitigation-MVP (scenario 1)

Legacy: cached values of the last full run stored in `CPAT 1.0pre_456_NoPropData.xlsb` (Egypt, real data), extracted to `data/legacy_cached_baseline.csv`. MVP: v1.01 scenario 1 recalculated in LibreOffice, (a) as shipped, last historical price year 2024; (b) last historical price year 2022 (Settings C10, LAMBDAs copied over whole rows). Fuel use ktoe, CO2 MtCO2.

Legacy codes not found: none.

## Fuel use by sector (ktoe)

| Sector | Run | 2022 | 2023 | 2024 | 2025 | 2027 | 2030 |
|---|---|---|---|---|---|---|---|
| Transport sector | legacy | 18,617 | 14,910 | 16,904 | 17,410 | 18,196 | 19,642 |
| | MVP, prices to 2024 | 18,617 | 23,542 | 26,129 | 26,688 | 27,891 | 30,169 |
| | MVP, prices to 2022 | 18,617 | 18,940 | 19,105 | 19,512 | 20,387 | 22,044 |
| Buildings sector | legacy | 7,729 | 6,188 | 7,076 | 7,156 | 7,326 | 7,635 |
| | MVP, prices to 2024 | 7,729 | 9,899 | 11,108 | 11,246 | 11,542 | 12,095 |
| | MVP, prices to 2022 | 7,729 | 7,803 | 7,840 | 7,933 | 8,132 | 8,504 |
| Industrial sector | legacy | 23,299 | 19,175 | 22,281 | 24,570 | 26,156 | 29,289 |
| | MVP, prices to 2024 | 23,299 | 28,698 | 30,737 | 31,759 | 33,758 | 37,692 |
| | MVP, prices to 2022 | 23,299 | 24,238 | 24,717 | 25,582 | 27,249 | 30,556 |
| Other energy use | legacy | 101 | 75 | 97 | 147 | 146 | 145 |
| | MVP, prices to 2024 | 101 | 120 | 132 | 136 | 145 | 162 |
| | MVP, prices to 2022 | 101 | 103 | 105 | 108 | 116 | 129 |

## CO2 from fuel combustion by sector (MtCO2)

| Sector | Run | 2022 | 2023 | 2024 | 2025 | 2027 | 2030 |
|---|---|---|---|---|---|---|---|
| Transport sector | legacy | 51.6 | 41.5 | 47.0 | 48.4 | 50.7 | 54.8 |
| | MVP, prices to 2024 | 53.4 | 67.4 | 74.7 | 76.3 | 79.7 | 86.2 |
| | MVP, prices to 2022 | 53.4 | 54.3 | 54.8 | 55.9 | 58.4 | 63.2 |
| Buildings sector | legacy | 15.5 | 12.4 | 14.3 | 14.6 | 15.0 | 15.9 |
| | MVP, prices to 2024 | 16.8 | 22.7 | 25.9 | 26.3 | 27.1 | 28.6 |
| | MVP, prices to 2022 | 16.8 | 17.0 | 17.1 | 17.3 | 17.9 | 18.9 |
| Industrial sector | legacy | 59.8 | 48.9 | 57.2 | 63.7 | 67.6 | 75.4 |
| | MVP, prices to 2024 | 64.3 | 80.2 | 86.5 | 89.4 | 94.7 | 105.3 |
| | MVP, prices to 2022 | 64.3 | 67.3 | 68.7 | 71.1 | 75.5 | 84.3 |
| Other energy use | legacy | 0.3 | 0.2 | 0.3 | 0.4 | 0.4 | 0.4 |
| | MVP, prices to 2024 | 0.3 | 0.4 | 0.4 | 0.4 | 0.5 | 0.5 |
| | MVP, prices to 2022 | 0.3 | 0.3 | 0.3 | 0.3 | 0.4 | 0.4 |

## After-tax prices, scenario 1 (legacy units: $/liter for liquids, $/GJ otherwise; real 2026 USD)

| Row | Run | 2022 | 2023 | 2024 | 2025 | 2027 | 2030 |
|---|---|---|---|---|---|---|---|
| rod gso | legacy | 0.612 | 0.612 | 0.604 | 0.604 | 0.604 | 0.604 |
| | MVP, prices to 2024 | 0.600 | 0.411 | 0.335 | 0.335 | 0.335 | 0.335 |
| | MVP, prices to 2022 | 0.600 | 0.600 | 0.600 | 0.600 | 0.600 | 0.600 |
| rod die | legacy | 0.495 | 0.508 | 0.470 | 0.470 | 0.470 | 0.470 |
| | MVP, prices to 2024 | 0.414 | 0.281 | 0.249 | 0.249 | 0.249 | 0.249 |
| | MVP, prices to 2022 | 0.414 | 0.414 | 0.414 | 0.414 | 0.414 | 0.414 |
| res lpg | legacy | 0.157 | 0.164 | 0.151 | 0.151 | 0.151 | 0.151 |
| | MVP, prices to 2024 | 0.175 | 0.102 | 0.088 | 0.088 | 0.088 | 0.088 |
| | MVP, prices to 2022 | 0.175 | 0.175 | 0.175 | 0.175 | 0.175 | 0.175 |
| res nga | legacy | 6.241 | 5.814 | 5.543 | 5.543 | 5.543 | 5.543 |
| | MVP, prices to 2024 | 4.634 | 3.209 | 2.331 | 2.331 | 2.331 | 2.331 |
| | MVP, prices to 2022 | 4.634 | 4.634 | 4.634 | 4.634 | 4.634 | 4.634 |
| cem coa | legacy | 10.523 | 10.300 | 9.578 | 8.485 | 9.223 | 9.016 |
| | MVP, prices to 2024 | 11.466 | 8.830 | 7.746 | 7.294 | 7.599 | 7.513 |
| | MVP, prices to 2022 | 11.466 | 8.596 | 7.940 | 7.484 | 7.792 | 7.705 |
| omn oop | legacy | 34.415 | 37.369 | 28.983 | 15.000 | 15.000 | 15.000 |
| | MVP, prices to 2024 | 33.051 | 26.301 | 23.176 | 23.176 | 23.176 | 23.176 |
| | MVP, prices to 2022 | 33.051 | 33.051 | 33.051 | 33.051 | 33.051 | 33.051 |

## Fit: fuel use by subsector and fuel, 2023-2030

| Run | Mean abs. % difference vs legacy | Total fuel use 2030 vs legacy |
|---|---|---|
| MVP, prices to 2024 | 45.6% | +41.3% |
| MVP, prices to 2022 | 17.3% | +8.0% |


## Regulated prices: how legacy handles them (check of 2026-10-09)

Read from the cached legacy Mitigation sheet: forecasting coefficients (rows 745-760), historical prices (rows 763-799), and the gasoline price block (rows 2589-2604).

1. **Same pass-through rule and values.**
   - Chosen = bucketed (default). Egypt: coal 1; gas (power, residential, industry) 0; gasoline, diesel, LPG, kerosene 0; other oil products 1 (the MVP uses 0 since v0.10, user decision); biomass 1.
   - The floating subsidy is set at its **2024 level** (legacy labels "2024 level"), so legacy's last historical year is also 2024.
   - With pass-through 0, the retail price stays at its 2024 value and the floating subsidy absorbs supply-cost changes. Legacy gasoline: rp 0.6036 $/liter from 2024 to 2040, cs -0.154 -> -0.048 ... as sp moves. This is the MVP's `txo` rule.
2. **No other flattening mechanism for Egypt.** Price-control phase-out (MTInputs 191-195) is off; the "override subsidy" column is empty for fuels; domestic production cost enters the fixed supply cost of gas (flag "produced domestically" = 1), as in the MVP.
3. **The difference is the historical price data (2021-2024).** Legacy's real 2026 USD values:

   | Price | Source | 2021 | 2022 | 2023 | 2024 |
   |---|---|---|---|---|---|
   | Gasoline retail ($/liter) | legacy data | 0.671 | 0.612 | 0.612 | 0.604 |
   | Gasoline retail ($/liter) | corrected block (MVP) | 0.69 | 0.600 | 0.411 | 0.335 |
   | Residential gas retail ($/GJ) | legacy data | 6.46 | 6.24 | 5.81 | 5.54 |
   | Residential gas retail ($/GJ) | corrected block (MVP) | 5.83 | 4.63 | 3.21 | 2.33 |

   Legacy's price block has no devaluation fall: its 2024 gasoline price is about 0.57 $/liter nominal, against 0.316 in the corrected block. The fall in the corrected block matches the EGP devaluation of 2023-2024.
4. **Consequence.** From 2025 both models hold prices flat, so neither keeps reacting to prices. The MVP's extra fuel use comes from responding to the 2022-2024 price fall, which legacy's data do not have. Legacy also calibrates 2023 emissions to estimates (the dip of about -20%).
5. **Options:**
   - (a) Keep the corrected data and accept the 2023-2024 response.
   - (b) Rebase fuel use to observed 2023/2024 energy data when available, and let the price response start from 2024.
   - (c) Add legacy's 2023 emissions calibration.
   - (d) Hold prices at 2022 (Settings C10 = 2022). This is closest to legacy but ignores the observed devaluation.

## Reality check: Egypt pump prices 2022-2024 (2026-10-09)

92-octane gasoline, official prices (EGP/liter), time-weighted annual averages from the fuel pricing committee decisions:
- 2022: 8.50 (Jan), 8.75 (Apr), 9.25 (Jul; held in Oct) -> about 8.94.
- 2023: 9.25, then 10.25 (2 Mar), then 11.50 (3 Nov) -> about 10.29.
- 2024: 11.50, then 12.50 (22 Mar), 13.75 (Jul), 15.25 (18 Oct) -> about 13.1.

Average EGP per USD: 19.2 (2022), 30.7 (2023), 45.4 (2024) (secondary compilers, e.g. FocusEconomics and Penn World Table via FRED; the official IMF/World Bank series is still to be pulled).

| $/liter | 2022 | 2023 | 2024 | 2024 vs 2022 |
|---|---|---|---|---|
| 92-octane, nominal (computed) | 0.465 | 0.336 | 0.289 | -38% |
| Corrected block, nominal (all grades) | 0.528 | 0.376 | 0.316 | -40% |
| Legacy block, nominal (0.612/0.612/0.604 real / CPI index) | 0.538 | 0.560 | 0.570 | +6% |

The corrected block follows reality: the EGP price rose about 47%, but the pound lost about 58% against the dollar. Legacy's block does not have the fall; its USD prices even rise. The corrected block's level is a little higher than 92-octane alone, consistent with an average over grades (95-octane is about 10% dearer).

Sources: egyptianstreets.com (2022-07-13, 2024-10-18), cairo.gov.eg (2022, 2024), businesstodayegypt.com, enterpriseam.com (2022-10-23), focus-economics.com, fred.stlouisfed.org (XRNCUSEGA618NRUG).
