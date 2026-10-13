# Check report - CPAT_Mitigation_CopyPaste_v0.6

## 1. Shipped file recalculated in LibreOffice (no LAMBDA support)
- error values: 1080, all in the 2035 (LAMBDA) columns ['AN', 'Y']: OK

## 2. Formula uniformity (relative R1C1)
- `pre`: base 1, plain 2023-2034 1, LAMBDA 2035 1, parameter columns D:G [1, 1, 1, 1] formula(s) OK
  - plain: `=R[0]C[-1]*(1+R[0]C4)`
  - LAMBDA: `=PRETAX(R[0]C[-1],R[0]C4)`
- `tax`: base 1, plain 2023-2034 1, LAMBDA 2035 1, parameter columns D:G [1, 1, 1, 1] formula(s) OK
  - plain: `=R[0]C4+R14C[0]*R[0]C5*R[0]C6`
  - LAMBDA: `=TAX(R[0]C4,R14C[0],R[0]C5,R[0]C6)`
- `post`: base 1, plain 2023-2034 1, LAMBDA 2035 1, parameter columns D:G [1, 1, 1, 1] formula(s) OK
  - plain: `=R[-262]C[0]+R[-131]C[0]`
  - LAMBDA: `=POSTTAX(R[-262]C[0],R[-131]C[0])`
- `use`: base 1, plain 2023-2034 1, LAMBDA 2035 1, parameter columns D:G [1, 1, 1, 1] formula(s) OK
  - plain: `=R[0]C[-1]*(1/(1+R[0]C7))^(1+R[0]C5)*(1+R7C[0])^R[0]C4*(R[-131]C[0]/R[-131]C[-1])^R[0]C5*(R[-131]C[0]/R[-131]C[-1])^(R[0]C6*(1+R[0]C5))`
  - LAMBDA: `=FUELUSE(R[0]C[-1],R[-131]C[0],R[-131]C[-1],R7C[0],R[0]C4,R[0]C5,R[0]C6,R[0]C7)`
- parameter columns D:G grouped and hidden: OK

## 3. LAMBDA encoding
- PRETAX: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- TAX: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- POSTTAX: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- FUELUSE: 8 parameters, 20 _xlpm. tokens, bare names [] OK

## 4. Values vs independent Python recomputation
- A. LAMBDA bodies expanded in 2035: 4 blocks x 128 rows x 2 scenarios x 14 years, max relative diff 4.70e-15 (at AG508) OK
- B. plain formula dragged into 2035: 4 blocks x 128 rows x 2 scenarios x 14 years, max relative diff 4.70e-15 (at AG508) OK

## 5. Scenario copy test
- scenario 3 (copy of 2) = scenario 2: 7588 numeric cells, max abs diff 0 OK
- scenario 4 (copy of 2, carbon price 0 in MTInputs) = scenario 1: 7588 numeric cells, max abs diff 0 OK
- pasted group 3: scenario number 3, name "Test scenario 3", code `egy.mit.ener.rod.coa.e.3` OK
- pasted group 4: scenario number 4, name "Test scenario 4", code `egy.mit.ener.rod.coa.e.4` OK
- pasted group 5: scenario number 5, name "Test scenario 5", code `egy.mit.ener.rod.coa.e.5` OK
- scenario 5 with legacy Egypt settings (2026, 0 -> 50 by 2030, Linear*): carbon price [0, 12.5, 50, 112.5] vs legacy row 8582, max abs diff 0 OK

## 6. Regression vs v0.5
- 14364 block and carbon-price cells vs v0.5: max abs diff 0 OK

## 7. Labels and codes
- column H: 548 rows, 1 formula OK
- column I: 548 rows, 1 formula OK
- column J: 548 rows, 1 formula OK
- code column K (scenario 1): 548 rows, 1 formula OK
- code column Z (scenario 2): 548 rows, 1 formula OK
- K413: `egy.mit.ener.rod.gso.e.1` | Fuel use | Road | Gasoline OK
- Z413: `egy.mit.ener.rod.gso.e.2` | Fuel use | Road | Gasoline OK
- K20: `egy.mit.sp.rod.gso.a.1` | Pre-tax price (supply cost) | Road | Gasoline OK
- Z282: `egy.mit.atp.rod.gso.e.2` | After-tax price | Road | Gasoline OK
- K7: `egy.mit.gdp.pos.pct.1` | Real GDP growth OK
- Z14: `egy.mit.cptraj.2` | Carbon price trajectory used OK
- K541: `egy.mit.ener.rod.all.e.1` | Fuel use | Road | All fuels OK
- Z410: `egy.mit.ener.all.all.e.2` | Fuel use | All subsectors | All fuels OK
- K148: `egy.mit.cptraj.ref.1` | Policy carbon price (as in the scenario assumptions) OK
- scenario numbers at the top of the code columns: [1, 2] OK

## 8. Row outline
- block `pre`: rows 18-146 grouped under summary line 17 (band 16 and summary visible), collapsed OK
- block `tax`: rows 149-277 grouped under summary line 148 (band 147 and summary visible), collapsed OK
- block `post`: rows 280-408 grouped under summary line 279 (band 278 and summary visible), collapsed OK
- block `use`: rows 411-539 grouped under summary line 410 (band 409 and summary visible), collapsed OK
- totals: rows 541-570 grouped under band row 540, open OK

## 9. MTInputs
- columns A:H, rows 1-415: 0 cells differ from the template OK
- columns D:E hidden as in the template: OK

## 10. Format and summary lines
- band rows [3, 16, 147, 278, 409, 540]: white bold text OK
- scenario 1: tax summary = carbon price (diff 0); fuel-use summary = sum of 128 rows (diff 1.16e-10) OK
- scenario 2: tax summary = carbon price (diff 0); fuel-use summary = sum of 128 rows (diff 8.73e-11) OK

## 11. Sanity and results
- scenario 1 vs 2 before 2027 (fuel use): max abs diff 0 OK
- oen rows (no coverage): scenario 1 = scenario 2, max abs diff 0 OK
- bio rows (EF 0): scenario 1 = scenario 2, max abs diff 0 OK
- check row 567 (totals by subsector and by fuel vs total): max abs 0 OK

| Year | Scenario 1 (ktoe) | Scenario 2 (ktoe) | Change |
|---|---|---|---|
| 2022 | 49,746 | 49,746 | +0.0% |
| 2026 | 53,909 | 53,909 | +0.0% |
| 2027 | 55,427 | 49,521 | -10.7% |
| 2030 | 60,733 | 54,214 | -10.7% |
| 2035 | 69,675 | 62,110 | -10.9% |

**Overall: PASS**
