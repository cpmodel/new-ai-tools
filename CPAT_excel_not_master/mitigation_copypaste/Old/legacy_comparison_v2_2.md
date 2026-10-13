# Legacy CPAT cached baseline vs CPAT-AI-Mitigation-MVP (scenario 1)

Legacy: cached values of the last full run stored in `CPAT 1.0pre_456_NoPropData.xlsb` (Egypt, real data), extracted to `data/legacy_cached_baseline.csv`. MVP: v2.2 scenario 1 recalculated in LibreOffice, (a) as shipped, last historical price year 2024; (b) last historical price year 2022 (Settings C10, LAMBDAs copied over whole rows). Fuel use ktoe, CO2 MtCO2.

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

