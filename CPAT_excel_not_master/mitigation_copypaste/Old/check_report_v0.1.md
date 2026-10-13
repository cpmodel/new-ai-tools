# Check report - CPAT_Mitigation_CopyPaste_v0.1

## 1. Recalculation (LibreOffice)
- error values on Mitigation: 0 OK

## 2. Formula uniformity (relative R1C1)
- `pre`: 128 rows x 2 groups -> 1 base-year formula, 1 projection formula OK
- `tax`: 128 rows x 2 groups -> 1 base-year formula, 1 projection formula OK
  - projection R1C1: `=INDEX(R[0]C17:R[0]C36,1,(R4C[0]-1)*5+1)+R8C[0]*INDEX(R[0]C17:R[0]C36,1,(R4C[0]-1)*5+2)*INDEX(R[0]C17:R[0]C36,1,(R4C[0]-1)*5+3)...`
- `post`: 128 rows x 2 groups -> 1 base-year formula, 1 projection formula OK
- `use`: 128 rows x 2 groups -> 1 base-year formula, 1 projection formula OK
  - projection R1C1: `=R[0]C[-1]*(1/(1+INDEX(R[0]C17:R[0]C36,1,(R4C[0]-1)*5+4)))^(1+INDEX(R[0]C17:R[0]C36,1,(R4C[0]-1)*5+2))*(1+R7C[0])^INDEX(R[0]C17:R[0]C36,1,(R4C[0]-1)*5+1)*(R[-130]C[0]/R[-130]C[-1])^INDEX(R[0]C17:R[0]C...`
- `tot_sub`: 16 rows x 2 groups -> 1 base-year formula, 1 projection formula OK
- `tot_fuel`: 8 rows x 2 groups -> 1 base-year formula, 1 projection formula OK
- `macro_gdp`: 1 rows x 2 groups -> 1 base-year formula, 1 projection formula OK
- parameter columns Q:AJ: one formula (or input/blank) per block and column OK

## 3. Scenario-group copy test
- group 3 (copy of 2) = group 2: 7588 numeric cells, max abs diff 0 OK
- group 4 (copy of 2, carbon price 0) = group 1: 7588 numeric cells, max abs diff 0 OK

## 4. Independent recomputation (Python, from data/*.csv)
- all 4 blocks x 128 rows x 2 groups x 14 years: max relative diff 4.70e-15 (at BG498) OK

## 5. Sanity
- scenario 1 vs 2 before 2027 (fuel use): max abs diff 0 OK
- oen rows (no coverage): scenario 1 = scenario 2, max abs diff 0 OK
- bio rows (EF 0): scenario 1 = scenario 2, max abs diff 0 OK
- check row 558 (sum by fuel - sum by subsector): max abs 0 OK

Results (total fuel use, ktoe, and change vs scenario 1):

| Year | Scenario 1 | Scenario 2 | Change |
|---|---|---|---|
| 2022 | 49,746 | 49,746 | +0.0% |
| 2026 | 53,909 | 53,909 | +0.0% |
| 2027 | 55,427 | 49,521 | -10.7% |
| 2030 | 60,733 | 54,214 | -10.7% |
| 2035 | 69,675 | 62,110 | -10.9% |

**Overall: PASS**
