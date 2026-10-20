# NMIM v0.1 check report

## 1. Workbook = Python (live formulas vs engine, key rows)

| Country | Max abs diff | Max abs value | Relative | Result |
| --- | --- | --- | --- | --- |
| IND | 4.82e-06 | 12,881.9 | 3.7e-10 | PASS |
| USA | 4.74e-06 | 46,127.1 | 1.0e-10 | PASS |
| CHN | 4.88e-06 | 46,123.4 | 1.1e-10 | PASS |
| BRA | 4.29e-07 | 4,323.1 | 9.9e-11 | PASS |
| KEN | 4.32e-08 | 345.2 | 1.3e-10 | PASS |
| EGY | 4.77e-07 | 1,084.7 | 4.4e-10 | PASS |
| DEU | 4.60e-07 | 6,464.3 | 7.1e-11 | PASS |
| IDN | 4.95e-07 | 4,189.0 | 1.2e-10 | PASS |

## 2. Identities (all countries, all years)

- PASS: fossil + VRE <= generation, scenario 1
- PASS: mit.inv.gen.1 >= 0
- PASS: mit.inv.grid.1 >= 0
- PASS: mit.cap.sol.1 >= 0
- PASS: mit.cap.fos.stk.1 >= 0
- PASS: mit.inv.trn.1 >= 0
- PASS: fossil + VRE <= generation, scenario 2
- PASS: mit.inv.gen.2 >= 0
- PASS: mit.inv.grid.2 >= 0
- PASS: mit.cap.sol.2 >= 0
- PASS: mit.cap.fos.stk.2 >= 0
- PASS: mit.inv.trn.2 >= 0
- PASS: total net zero investment = sum of sectors
- PASS: net zero GHG 2050 below baseline (every country)

## 3. Benchmarks (world)

| Item | Model | Benchmark | Source |
| --- | --- | --- | --- |
| Net zero investment 2030, USD bn | 5,103 | 4,800 clean energy + 400 nature | IHLEG (2024), IEA |
| of which EMDCs ex-China | 1,303 | 1,600 clean energy + 350 nature | IHLEG (2024) |
| Baseline investment 2030 (calibrated to 2024 by group) | 2,084 | ~2,000 clean energy in 2024 | IEA World Energy Investment 2024 |
| Baseline GHG 2030, MtCO2e | 58,156 | ~53,000-57,000 | UNEP Emissions Gap (current policies, approx.) |
| Energy CO2 2024, 11 end-use rows + power vs CPAT | 33,627 | 39,045 | CPAT 2024 panel; the gap is carried as other energy CO2 (mit.co2.res), so totals tie to CPAT |
| Net zero GHG 2050 (residual) | 10,029 | ~0 net with removals | Residual to be met by removals (not modelled) |

**Overall: PASS**
