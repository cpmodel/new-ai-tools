# Check report - CPAT_Mitigation_CopyPaste_v0.8

## 1. Shipped file recalculated in LibreOffice (no LAMBDA support)
- error values: 1120, all in the 2035 (LAMBDA) columns: OK

## 2. Formula uniformity (relative R1C1)
- `sp` base year: 256 cells, 1 formula OK
- `sp` 2023-2034: 3072 cells, 1 formula OK
- `sp` 2035: 256 cells, 1 formula OK
- `sp` column D: 128 cells, 1 formula OK
- `sp` column E: 128 cells, 1 formula OK
- `sp` column F: 128 cells, 1 formula OK
- `sp` column G: 128 cells, 1 formula OK
  - `=R[0]C[-1]*(1+R[0]C4)`
  - 2035: `=PRETAX(R[0]C[-1],R[0]C4)`
- `ctxnew` base year: 256 cells, 1 formula OK
- `ctxnew` 2023-2034: 3072 cells, 1 formula OK
- `ctxnew` 2035: 256 cells, 1 formula OK
- `ctxnew` column D: 128 cells, 1 formula OK
- `ctxnew` column E: 128 cells, 1 formula OK
- `ctxnew` column F: 128 cells, 1 formula OK
- `ctxnew` column G: 128 cells, 1 formula OK
  - `=R10C[0]*R[0]C4*INDEX(R18C[0]:R25C[0],R[0]C5)*INDEX(R27C[0]:R43C[0],R[0]C6)`
- `ntx` base year: 256 cells, 1 formula OK
- `ntx` 2023-2034: 3072 cells, 1 formula OK
- `ntx` 2035: 256 cells, 1 formula OK
- `ntx` column D: 128 cells, 1 formula OK
- `ntx` column E: 128 cells, 1 formula OK
- `ntx` column F: 128 cells, 1 formula OK
- `ntx` column G: 128 cells, 1 formula OK
  - `=INDEX(R75C[0]:R88C[0],R[0]C4)/R[0]C5`
- `nce` base year: 256 cells, 1 formula OK
- `nce` 2023-2034: 3072 cells, 1 formula OK
- `nce` 2035: 256 cells, 1 formula OK
- `nce` column D: 128 cells, 1 formula OK
- `nce` column E: 128 cells, 1 formula OK
- `nce` column F: 128 cells, 1 formula OK
- `nce` column G: 128 cells, 1 formula OK
  - `=R[-18]C[0]+R[-9]C[0]`
- `tax` base year: 256 cells, 1 formula OK
- `tax` 2023-2034: 3072 cells, 1 formula OK
- `tax` 2035: 256 cells, 1 formula OK
- `tax` column D: 128 cells, 1 formula OK
- `tax` column E: 128 cells, 1 formula OK
- `tax` column F: 128 cells, 1 formula OK
- `tax` column G: 128 cells, 1 formula OK
  - `=R[0]C4+R[-9]C[0]`
  - 2035: `=TAX(R[0]C4,R[-9]C[0])`
- `atp` base year: 256 cells, 1 formula OK
- `atp` 2023-2034: 3072 cells, 1 formula OK
- `atp` 2035: 256 cells, 1 formula OK
- `atp` column D: 128 cells, 1 formula OK
- `atp` column E: 128 cells, 1 formula OK
- `atp` column F: 128 cells, 1 formula OK
- `atp` column G: 128 cells, 1 formula OK
  - `=R[-45]C[0]+R[-9]C[0]`
  - 2035: `=POSTTAX(R[-45]C[0],R[-9]C[0])`
- `shp` base year: 256 cells, 1 formula OK
- `shp` 2023-2034: 3072 cells, 1 formula OK
- `shp` 2035: 256 cells, 1 formula OK
- `shp` column D: 128 cells, 1 formula OK
- `shp` column E: 128 cells, 1 formula OK
- `shp` column F: 128 cells, 1 formula OK
- `shp` column G: 128 cells, 1 formula OK
  - `=INDEX(R124C[0]:R127C[0],R[0]C4)*R[0]C5*INDEX(R128C[0]:R144C[0],R[0]C6)`
- `ener` base year: 256 cells, 1 formula OK
- `ener` 2023-2034: 3072 cells, 1 formula OK
- `ener` 2035: 256 cells, 1 formula OK
- `ener` column D: 128 cells, 1 formula OK
- `ener` column E: 128 cells, 1 formula OK
- `ener` column F: 128 cells, 1 formula OK
- `ener` column G: 128 cells, 1 formula OK
  - `=R[0]C[-1]*(1/(1+R[0]C7))^(1+R[0]C5)*(1+R7C[0])^R[0]C4*(R[-18]C[0]/R[-18]C[-1])^R[0]C5*((R[-18]C[0]+R[-9]C[0])/(R[-18]C[-1]+R[-9]C[-1]))^(R[0]C6*(1+R[`
  - 2035: `=FUELUSE(R[0]C[-1],R[-18]C[0],R[-18]C[-1],R[-9]C[0],R[-9]C[-1],R7C[0],R[0]C4,R[0]C5,R[0]C6,R[0]C7)`
- Section 1 inputs: 2408 cells, 1 formula OK
- Section 1 paths `fpr.*`: 392 cells, 1 formula OK
- Section 1 paths `fb.*`: 112 cells, 1 formula OK
- Section 1 paths `shps.*`: 112 cells, 1 formula OK
- Section 1 paths `ssc.*`: 476 cells, 1 formula OK
- subsector heading totals: 448 cells, 1 formula OK
- parameter columns D:G grouped and hidden: OK

## 3. LAMBDA encoding
- PRETAX: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- TAX: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- POSTTAX: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- FUELUSE: 10 parameters, 24 _xlpm. tokens, bare names [] OK

## 4. Values vs independent Python recomputation
- A. LAMBDA expanded, scenario 1: 8 variables x 128 rows x 14 years, max relative diff 4.46e-15 (at N1325) OK
- B. plain formula dragged, scenario 1: 8 variables x 128 rows x 14 years, max relative diff 4.46e-15 (at N1325) OK
- A. LAMBDA expanded, scenario 2: 8 variables x 128 rows x 14 years, max relative diff 4.70e-15 (at AG1098) OK
- B. plain formula dragged, scenario 2: 8 variables x 128 rows x 14 years, max relative diff 4.70e-15 (at AG1098) OK

## 5. Scenario tests (copied MTInputs column + Mitigation group)
- scenario 3 (copy of 2) = scenario 2: 16814 numeric cells, max abs diff 0 OK
- scenario 4 (copy of 2, carbon price 0 in MTInputs) = scenario 1: 16786 numeric cells, max abs diff 0 OK
- pasted group 3: number 3, name "Test 3: copy of 2", code `egy.mit.ener.rod.coa.e.3` OK
- pasted group 4: number 4, name "Test 4: copy of 2, carbon price 0", code `egy.mit.ener.rod.coa.e.4` OK
- pasted group 5: number 5, name "Test 5: legacy Egypt carbon tax", code `egy.mit.ener.rod.coa.e.5` OK
- pasted group 6: number 6, name "Test 6: fuel price reform", code `egy.mit.ener.rod.coa.e.6` OK
- pasted group 7: number 7, name "Test 7: feebates", code `egy.mit.ener.rod.coa.e.7` OK
- scenario 5 (legacy Egypt carbon tax): carbon price [0, 12.5, 50, 112.5] vs legacy row 8582, max abs diff 0 OK
- scenario 5 vs Python: 8 variables x 128 rows x 14 years, max relative diff 4.81e-15 (at BY218) OK
- scenario 6 vs Python: 8 variables x 128 rows x 14 years, max relative diff 4.46e-15 (at CK1325) OK
- scenario 7 vs Python: 8 variables x 128 rows x 14 years, max relative diff 4.46e-15 (at CZ1325) OK
  - scenario 6, road gasoline new excise ($/GJ) 2026-2031: [0, 1.43, 2.86, 4.289, 5.719, 7.149]
  - scenario 7, road gasoline feebate shadow price ($/GJ) 2026-2031: [0, 0.662, 1.544, 2.427, 3.309, 4.191]
- scenario 7 vs scenario 1 (covered with a non-zero rate: cem, cst, ftr, irn, mac, mch, nfm, oen, omn, rod): other subsectors unchanged (max abs diff 0); covered subsectors in 2035: cem -12.20%, ftr -12.19%, irn -12.31%, oen -13.46%, omn -12.59%, rod -11.09% OK

## 6. Regression vs v0.7 (shared output codes)
- 2322 output codes in both versions, max abs diff 0 OK; intended change: egy.mit.ctcov.all.bio.1, egy.mit.ctcov.all.bio.2 (biomass carbon-tax coverage FALSE -> TRUE; EF 0)

## 7. Labels and codes
- column H: 1201 cells, 1 formula OK
- column I: 1201 cells, 1 formula OK
- column J: 1201 cells, 1 formula OK
- code column K (scenario 1): 1201 cells, 1 formula OK
- code column Z (scenario 2): 1201 cells, 1 formula OK
- K217: `egy.mit.ener.rod.gso.e.1` | Fuel use | Road | Gasoline OK
- Z199: `egy.mit.atp.rod.gso.e.2` | After-tax price | Road | Gasoline OK
- Z164: `egy.mit.ctxnew.rod.die.a.2` | New carbon tax | Road | Diesel OK
- K208: `egy.mit.shp.rod.gso.1` | Shadow price on the efficiency margin | Road | Gasoline OK
- K151: `egy.mit.ener.rod.all.e.1` | Fuel use | Road | All fuels OK
- K150: `egy.mit.ener.tra.all.e.1` | Fuel use | Transport | All fuels OK
- Z10: `egy.mit.cptraj.2` | Carbon price trajectory used OK
- K43: `egy.mit.ctcov.oen.all.1` | Carbon tax coverage (Apply tax?) | Other energy use | All fuels OK
- Z1330: `egy.mit.ener.all.all.e.2` | Fuel use | All subsectors | All fuels OK

## 8. Row outline
- carbon price (summary of section 1): 1 rows at level 0, visible OK
- section 1 inputs and paths: 127 rows at level 1, hidden OK
- sector totals: 4 rows at level 0, visible OK
- subsector headings: 16 rows at level 1, visible OK
- subsector variables: 1024 rows at level 2, hidden OK
- results: 30 rows at level 1, visible OK

## 9. MTInputs
- columns A:H, rows 1-415: 0 cells differ from the template; D:E hidden OK

## 10. Format
- section bands ['Scenario assumptions (', '1. Policies (scenario ', '3. Power sector (elast', '5. Transport sector', '6. Buildings sector', '7. Industrial sector', '8. Other energy use', '11. Results - energy c']: white bold text OK

## 11. Sanity and results
- scenario 1 vs 2 before 2027 (fuel use): max abs diff 0 OK
- other energy use not taxed (MCovOen FALSE): fuel use scenario 1 = 2, max abs diff 0 OK
- check row 1358: max abs 0; sector totals sum to total: max abs 8.73e-11 OK

| Year | Scenario 1 (ktoe) | Scenario 2 (ktoe) | Change |
|---|---|---|---|
| 2022 | 49,746 | 49,746 | +0.0% |
| 2026 | 53,909 | 53,909 | +0.0% |
| 2027 | 55,427 | 49,521 | -10.7% |
| 2030 | 60,733 | 54,214 | -10.7% |
| 2035 | 69,675 | 62,110 | -10.9% |

**Overall: PASS**
