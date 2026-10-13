# Check report - CPAT_Mitigation_CopyPaste_v0.4

## 1. Shipped file recalculated in LibreOffice (no LAMBDA support)
- error values: 1080, all in the 2035 (LAMBDA) columns ['AN', 'Y']: OK

## 2. Formula uniformity (relative R1C1)
- `pre`: base 1, plain 2023-2034 1, LAMBDA 2035 1, parameter columns D:G [1, 1, 1, 1] formula(s) OK
  - plain: `=R[0]C[-1]*(1+R[0]C4)`
  - LAMBDA: `=PRETAX(R[0]C[-1],R[0]C4)`
- `tax`: base 1, plain 2023-2034 1, LAMBDA 2035 1, parameter columns D:G [1, 1, 1, 1] formula(s) OK
  - plain: `=R[0]C4+R7C[0]*R[0]C5*R[0]C6`
  - LAMBDA: `=TAX(R[0]C4,R7C[0],R[0]C5,R[0]C6)`
- `post`: base 1, plain 2023-2034 1, LAMBDA 2035 1, parameter columns D:G [1, 1, 1, 1] formula(s) OK
  - plain: `=R[-260]C[0]+R[-130]C[0]`
  - LAMBDA: `=POSTTAX(R[-260]C[0],R[-130]C[0])`
- `use`: base 1, plain 2023-2034 1, LAMBDA 2035 1, parameter columns D:G [1, 1, 1, 1] formula(s) OK
  - plain: `=R[0]C[-1]*(1/(1+R[0]C7))^(1+R[0]C5)*(1+R6C[0])^R[0]C4*(R[-130]C[0]/R[-130]C[-1])^R[0]C5*(R[-130]C[0]/R[-130]C[-1])^(R[0]C6*(1+R[0]C5))`
  - LAMBDA: `=FUELUSE(R[0]C[-1],R[-130]C[0],R[-130]C[-1],R6C[0],R[0]C4,R[0]C5,R[0]C6,R[0]C7)`
- parameter columns D:G grouped and hidden: OK

## 3. LAMBDA encoding
- PRETAX: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- TAX: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- POSTTAX: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- FUELUSE: 8 parameters, 20 _xlpm. tokens, bare names [] OK

## 4. Values vs independent Python recomputation
- A. LAMBDA bodies expanded in 2035: 4 blocks x 128 rows x 2 scenarios x 14 years, max relative diff 4.70e-15 (at AG497) OK
- B. plain formula dragged into 2035: 4 blocks x 128 rows x 2 scenarios x 14 years, max relative diff 4.70e-15 (at AG497) OK

## 5. Scenario copy test
- scenario 3 (copy of 2) = scenario 2: 7588 numeric cells, max abs diff 0 OK
- scenario 4 (copy of 2, carbon price 0) = scenario 1: 7588 numeric cells, max abs diff 0 OK
- pasted group 3: scenario number 3, code `egy.mit.ener.rod.coa.e.3` OK
- pasted group 4: scenario number 4, code `egy.mit.ener.rod.coa.e.4` OK

## 6. Regression vs v0.3
- 14336 block cells vs v0.3: max abs diff 0 OK

## 7. Labels and codes
- column H: 542 rows, 1 formula OK
- column I: 542 rows, 1 formula OK
- column J: 542 rows, 1 formula OK
- code column K (scenario 1): 542 rows, 1 formula OK
- code column Z (scenario 2): 542 rows, 1 formula OK
- K402: `egy.mit.ener.rod.gso.e.1` | Fuel use | Road | Gasoline OK
- Z402: `egy.mit.ener.rod.gso.e.2` | Fuel use | Road | Gasoline OK
- K12: `egy.mit.sp.rod.gso.a.1` | Pre-tax price (supply cost) | Road | Gasoline OK
- Z272: `egy.mit.atp.rod.gso.e.2` | After-tax price | Road | Gasoline OK
- K6: `egy.mit.gdp.pos.pct.1` | Real GDP growth OK
- Z7: `egy.mit.cptraj.2` | Carbon price OK
- K530: `egy.mit.ener.rod.all.e.1` | Fuel use | Road | All fuels OK
- Z556: `egy.mit.ener.all.all.e.2` | Fuel use | All subsectors | All fuels OK
- scenario numbers at the top of the code columns: [1, 2] OK

## 8. Row outline
- block `pre`: rows 10-138 grouped under band row 9, collapsed OK
- block `tax`: rows 140-268 grouped under band row 139, collapsed OK
- block `post`: rows 270-398 grouped under band row 269, collapsed OK
- block `use`: rows 400-528 grouped under band row 399, collapsed OK
- totals: rows 530-560 grouped under band row 529, open OK

## 9. Sanity and results
- scenario 1 vs 2 before 2027 (fuel use): max abs diff 0 OK
- oen rows (no coverage): scenario 1 = scenario 2, max abs diff 0 OK
- bio rows (EF 0): scenario 1 = scenario 2, max abs diff 0 OK
- check row 557 (sum by fuel - sum by subsector): max abs 0 OK

| Year | Scenario 1 (ktoe) | Scenario 2 (ktoe) | Change |
|---|---|---|---|
| 2022 | 49,746 | 49,746 | +0.0% |
| 2026 | 53,909 | 53,909 | +0.0% |
| 2027 | 55,427 | 49,521 | -10.7% |
| 2030 | 60,733 | 54,214 | -10.7% |
| 2035 | 69,675 | 62,110 | -10.9% |

**Overall: PASS**
