# Check report - CPAT_Mitigation_CopyPaste_v0.10

## 1. Shipped file recalculated in LibreOffice (no LAMBDA support)
- error values: 680, all in the 2035 (LAMBDA) columns: OK

## 2. Formula uniformity (relative R1C1)
- `sp` base year: 24 cells, 1 formula OK
- `sp` 2023-2034: 288 cells, 1 formula OK
- `sp` 2035: 24 cells, 1 formula OK
- `sp` column D: 12 cells, 1 formula OK
- `sp` column E: 12 cells, 1 formula OK
- `sp` column F: 12 cells, 1 formula OK
- `sp` column G: 12 cells, 1 formula OK
- `sp` 2022-2034 (base year included): 312 cells, 1 formula OK
  - `=IF(R4C[0]<=Settings!R10C3,INDEX(Inputs_prices!R5C20:R16C23,R[0]C6,MATCH(R4C[0],Inputs_prices!R4C20:R4C23,0)),R[0]C4+(R[0]C[-1]-R[0]C4)*INDEX(R153C[0]`
  - 2035: `=IF(R4C[0]<=Settings!R10C3,INDEX(Inputs_prices!R5C20:R16C23,R[0]C6,MATCH(R4C[0],Inputs_prices!R4C20:R4C23,0)),SUPPLYCOST(R[0]C4,R[0]C[-1],INDEX(R153C[`
- `txo` base year: 24 cells, 1 formula OK
- `txo` 2023-2034: 288 cells, 1 formula OK
- `txo` 2035: 24 cells, 1 formula OK
- `txo` column D: 12 cells, 1 formula OK
- `txo` column E: 12 cells, 1 formula OK
- `txo` column F: 12 cells, 1 formula OK
- `txo` column G: 12 cells, 1 formula OK
- `txo` 2022-2034 (base year included): 312 cells, 1 formula OK
  - `=IF(R4C[0]<=Settings!R10C3,INDEX(Inputs_prices!R5C24:R16C27,R[0]C7,MATCH(R4C[0],Inputs_prices!R4C24:R4C27,0)),R[0]C4*R[0]C6+IF(R[0]C4*(1-R[0]C6)>=0,MA`
  - 2035: `=IF(R4C[0]<=Settings!R10C3,INDEX(Inputs_prices!R5C24:R16C27,R[0]C7,MATCH(R4C[0],Inputs_prices!R4C24:R4C27,0)),OTHERTAX(R[0]C4,R[0]C5,R[-13]C[0],R[0]C6`
- `rpb` base year: 24 cells, 1 formula OK
- `rpb` 2023-2034: 288 cells, 1 formula OK
- `rpb` 2035: 24 cells, 1 formula OK
- `rpb` column D: 12 cells, 1 formula OK
- `rpb` column E: 12 cells, 1 formula OK
- `rpb` column F: 12 cells, 1 formula OK
- `rpb` column G: 12 cells, 1 formula OK
- `rpb` 2022-2034 (base year included): 312 cells, 1 formula OK
  - `=(R[-26]C[0]+R[-13]C[0])*(1+R[0]C4)`
- `ctxnew` base year: 256 cells, 1 formula OK
- `ctxnew` 2023-2034: 3072 cells, 1 formula OK
- `ctxnew` 2035: 256 cells, 1 formula OK
- `ctxnew` column D: 128 cells, 1 formula OK
- `ctxnew` column E: 128 cells, 1 formula OK
- `ctxnew` column F: 128 cells, 1 formula OK
- `ctxnew` column G: 128 cells, 1 formula OK
  - `=R12C[0]*R[0]C4*INDEX(R21C[0]:R28C[0],R[0]C5)*INDEX(R30C[0]:R46C[0],R[0]C6)`
- `ntx` base year: 256 cells, 1 formula OK
- `ntx` 2023-2034: 3072 cells, 1 formula OK
- `ntx` 2035: 256 cells, 1 formula OK
- `ntx` column D: 128 cells, 1 formula OK
- `ntx` column E: 128 cells, 1 formula OK
- `ntx` column F: 128 cells, 1 formula OK
- `ntx` column G: 128 cells, 1 formula OK
  - `=INDEX(R78C[0]:R91C[0],R[0]C4)/R[0]C5`
- `nce` base year: 256 cells, 1 formula OK
- `nce` 2023-2034: 3072 cells, 1 formula OK
- `nce` 2035: 256 cells, 1 formula OK
- `nce` column D: 128 cells, 1 formula OK
- `nce` column E: 128 cells, 1 formula OK
- `nce` column F: 128 cells, 1 formula OK
- `nce` column G: 128 cells, 1 formula OK
  - `=R[-18]C[0]+R[-9]C[0]`
- `atp` base year: 256 cells, 1 formula OK
- `atp` 2023-2034: 3072 cells, 1 formula OK
- `atp` 2035: 256 cells, 1 formula OK
- `atp` column D: 128 cells, 1 formula OK
- `atp` column E: 128 cells, 1 formula OK
- `atp` column F: 128 cells, 1 formula OK
- `atp` column G: 128 cells, 1 formula OK
  - `=MAX(INDEX(R184C[0]:R195C[0],R[0]C4)+R[-9]C[0]*(1+R[0]C5),0.01)`
  - 2035: `=POSTTAX(INDEX(R184C[0]:R195C[0],R[0]C4),R[-9]C[0],R[0]C5)`
- `shp` base year: 256 cells, 1 formula OK
- `shp` 2023-2034: 3072 cells, 1 formula OK
- `shp` 2035: 256 cells, 1 formula OK
- `shp` column D: 128 cells, 1 formula OK
- `shp` column E: 128 cells, 1 formula OK
- `shp` column F: 128 cells, 1 formula OK
- `shp` column G: 128 cells, 1 formula OK
  - `=INDEX(R127C[0]:R130C[0],R[0]C4)*R[0]C5*INDEX(R131C[0]:R147C[0],R[0]C6)`
- `ener` base year: 256 cells, 1 formula OK
- `ener` 2023-2034: 3072 cells, 1 formula OK
- `ener` 2035: 256 cells, 1 formula OK
- `ener` column D: 128 cells, 1 formula OK
- `ener` column E: 128 cells, 1 formula OK
- `ener` column F: 128 cells, 1 formula OK
- `ener` column G: 128 cells, 1 formula OK
  - `=R[0]C[-1]*(1/(1+R[0]C7))^(1+R[0]C5)*(1+R7C[0])^R[0]C4*(R[-18]C[0]/R[-18]C[-1])^R[0]C5*((R[-18]C[0]+R[-9]C[0])/(R[-18]C[-1]+R[-9]C[-1]))^(R[0]C6*(1+R[`
  - 2035: `=FUELUSE(R[0]C[-1],R[-18]C[0],R[-18]C[-1],R[-9]C[0],R[-9]C[-1],R7C[0],R[0]C4,R[0]C5,R[0]C6,R[0]C7)`
- Section 1 and 2 MTInputs rows: 2492 cells, 1 formula OK
- section 2 international prices `gp`: 112 cells, 1 formula OK
- index row `infl`: 28 cells, 1 formula OK
- index row `defl`: 28 cells, 1 formula OK
- Section 1 paths `fpr.*`: 392 cells, 1 formula OK
- Section 1 paths `fb.*`: 112 cells, 1 formula OK
- Section 1 paths `shps.*`: 112 cells, 1 formula OK
- Section 1 paths `ssc.*`: 476 cells, 1 formula OK
- subsector heading totals: 448 cells, 1 formula OK
- parameter columns D:G grouped and hidden: OK

## 3. LAMBDA encoding
- SUPPLYCOST: 4 parameters, 9 _xlpm. tokens, bare names [] OK
- OTHERTAX: 4 parameters, 20 _xlpm. tokens, bare names [] OK
- POSTTAX: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- FUELUSE: 10 parameters, 24 _xlpm. tokens, bare names [] OK

## 4. Values vs independent Python recomputation
- A. LAMBDA expanded, scenario 1: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 4.53e-15 (at S1026) OK
- B. plain formula dragged, scenario 1: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 4.53e-15 (at S1026) OK
- A. LAMBDA expanded, scenario 2: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 4.36e-15 (at AB163) OK
- B. plain formula dragged, scenario 2: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 4.36e-15 (at AB163) OK

## 5. Scenario tests (copied MTInputs column + Mitigation group)
- scenario 3 (copy of 2) = scenario 2: 13818 numeric cells, max abs diff 0 OK
- scenario 4 (copy of 2, carbon price 0 in MTInputs) = scenario 1: 13790 numeric cells, max abs diff 0 OK
- pasted group 3: number 3, name "Test 3: copy of 2", code `egy.mit.ener.rod.coa.e.3` OK
- pasted group 4: number 4, name "Test 4: copy of 2, carbon price 0", code `egy.mit.ener.rod.coa.e.4` OK
- pasted group 5: number 5, name "Test 5: legacy Egypt carbon tax", code `egy.mit.ener.rod.coa.e.5` OK
- pasted group 6: number 6, name "Test 6: fuel price reform", code `egy.mit.ener.rod.coa.e.6` OK
- pasted group 7: number 7, name "Test 7: feebates", code `egy.mit.ener.rod.coa.e.7` OK
- pasted group 8: number 8, name "Test 8: IMF-WB*, High, nominal carbon price", code `egy.mit.ener.rod.coa.e.8` OK
- scenario 5 (legacy Egypt carbon tax): carbon price [0, 12.5, 50, 112.5] vs legacy row 8582, max abs diff 0 OK
- scenario 5 vs Python: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 4.78e-15 (at CE235) OK
- scenario 6 vs Python: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 4.87e-15 (at CO251) OK
- scenario 7 vs Python: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 4.62e-15 (at DJ1026) OK
- scenario 8 vs Python: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 4.45e-15 (at DT916) OK
  - scenario 6, road gasoline new excise ($/GJ) 2026-2031: [0, 1.43, 2.86, 4.289, 5.719, 7.149]
  - scenario 7, road gasoline feebate shadow price ($/GJ) 2026-2031: [0, 0.662, 1.544, 2.427, 3.309, 4.191]
- scenario 7 vs scenario 1 (covered with a non-zero rate: cem, cst, ftr, irn, mac, mch, nfm, oen, omn, rod): other subsectors unchanged (max abs diff 0); covered subsectors in 2034: cem -12.00%, ftr -12.74%, irn -10.98%, oen -14.47%, omn -14.19%, rod -13.62% OK
  - scenario 8: crude oil gp 2024-2026 (real) [84.95, 109.03, 114.14], carbon price 2027-2030 (nominal 20 x infl) [19.581, 19.162, 18.752, 18.35]
- price controls None (global), scenario 1 vs Python: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 2.60e-14 (at X473) OK
- price controls None (global), scenario 2 vs Python: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 1.23e-14 (at AM919) OK
- price controls Manual (global), scenario 1 vs Python: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 5.41e-15 (at P251) OK
- price controls Manual (global), scenario 2 vs Python: gp, sp, txo, rpb and 6 subsector variables, 11312 cells, max relative diff 5.41e-15 (at AE251) OK

## 6. Regression vs v0.9 (shared output codes)
- 1300 output codes in both versions and not affected (policies, gp, ctxnew, ntx, nce, shp), max abs diff 0 OK
- intended changes: 646 codes of sp, txo, rpb, atp and ener (oop treatment, VAT assumption); variables removed: ; added: 

| Output code | v0.9 2022 | v0.10 2022 | v0.9 2027 | v0.10 2027 | v0.9 2030 | v0.10 2030 |
|---|---|---|---|---|---|---|
| `egy.mit.rpb.all.oop.a.1` | 5.400 | 6.157 | 1.199 | 4.317 | 0.226 | 4.317 |
| `egy.mit.atp.rod.gso.e.1` | 15.051 | 17.159 | 8.400 | 9.576 | 8.400 | 9.576 |
| `egy.mit.atp.res.nga.e.1` | 4.634 | 4.634 | 2.331 | 2.331 | 2.331 | 2.331 |
| `egy.mit.atp.cem.coa.e.1` | 11.467 | 11.467 | 7.576 | 7.576 | 7.492 | 7.492 |
| `egy.mit.atp.rod.gso.e.2` | 15.051 | 17.159 | 9.724 | 11.085 | 9.724 | 11.085 |
| `egy.mit.ener.all.oop.e.1` | 3,032.015 | 3,032.015 | 9,037.208 | 4,355.190 | 29,178.098 | 4,869.442 |
| `egy.mit.ener.all.all.e.1` | 49,745.912 | 49,745.912 | 78,821.637 | 74,139.619 | 105,302.894 | 80,994.239 |
| `egy.mit.ener.all.all.e.2` | 49,745.912 | 49,745.912 | 65,016.706 | 62,965.243 | 73,605.106 | 68,791.324 |

## 7. Labels and codes
- column H: 990 cells, 1 formula OK
- column I: 990 cells, 1 formula OK
- column J: 990 cells, 1 formula OK
- code column K (scenario 1): 990 cells, 1 formula OK
- code column Z (scenario 2): 990 cells, 1 formula OK
- K251: `egy.mit.ener.rod.gso.e.1` | Fuel use | Road | Gasoline OK
- K158: `egy.mit.sp.pow.coa.a.1` | Supply cost (pre-tax price) | Power | Coal OK
- Z190: `egy.mit.rpb.all.gso.a.2` | Retail price before new policies | All subsectors | Gasoline OK
- K153: `egy.mit.gp.int.oil.1` | International energy price (real; source and adjustment from MTInputs) | International | Crude oil OK
- K8: `egy.mit.infl.1` | Inflation index: US CPI, ResultsYear = 1 (nominal -> real for domestic prices and nominal policy inputs) OK
- Z233: `egy.mit.atp.rod.gso.e.2` | After-tax price | Road | Gasoline OK
- Z207: `egy.mit.ctxnew.rod.die.a.2` | New carbon tax | Road | Diesel OK
- K242: `egy.mit.shp.rod.gso.1` | Shadow price on the efficiency margin | Road | Gasoline OK
- K203: `egy.mit.ener.rod.all.e.1` | Fuel use | Road | All fuels OK
- K202: `egy.mit.ener.tra.all.e.1` | Fuel use | Transport | All fuels OK
- Z12: `egy.mit.cptraj.2` | Carbon price trajectory used OK
- K46: `egy.mit.ctcov.oen.all.1` | Carbon tax coverage (Apply tax?) | Other energy use | All fuels OK
- Z1094: `egy.mit.ener.all.all.e.2` | Fuel use | All subsectors | All fuels OK

## 8. Row outline
- carbon price (summary of section 1): 1 rows at level 0, visible OK
- section 1 inputs and paths: 128 rows at level 1, hidden OK
- section 2 retail prices before new policies: 12 rows at level 0, visible OK
- section 2 selectors, gp, sp, txo: 30 rows at level 1, hidden OK
- sector totals: 4 rows at level 0, visible OK
- subsector headings: 16 rows at level 1, visible OK
- subsector variables: 768 rows at level 2, hidden OK
- results: 30 rows at level 1, visible OK

## 9. MTInputs
- columns A:H, rows 1-415: 0 cells differ from the template; D:E hidden OK

## 10. Format
- section bands ['Scenario assumptions (', '1. Policies (scenario ', '2. Retail energy price', '3. Power sector (elast', '5. Transport sector', '6. Buildings sector', '7. Industrial sector', '8. Other energy use', '11. Results - energy c']: white bold text OK
- Inputs_prices: data changed by assumption marked bright yellow (columns ['Margin (mit.mar, base year, nominal pric', 'Raw pass-through coefficient (mit.ps, ba', 'VAT applies to final consumers? (assumpt', 'VAT rate assumption = general VAT rate (']; other oil products pass-through and margin) OK

## 11. Sanity and results
- scenario 1 vs 2 before 2027 (fuel use): max abs diff 0 OK
- other energy use not taxed (MCovOen FALSE): fuel use scenario 1 = 2, max abs diff 0 OK
- check row 1122: max abs 0; sector totals sum to total: max abs 1.16e-10 OK
- pass-through 0 (nga.pow, nga.res, nga.ind, gso.all, die.all, lpg.all, ker.all, oop.all): retail price before new policies stays at its 2024 value, max abs diff 0 OK

Retail price before new policies, scenario 1 (real $/GJ of ResultsYear), and chosen pass-through:

| Price fuel | pass-through | 2022 | 2024 | 2027 | 2030 | 2034 |
|---|---|---|---|---|---|---|
| coa.pow | 1 | 13.324 | 5.912 | 5.677 | 5.540 | 5.492 |
| coa.res | 1 | 23.189 | 14.740 | 14.472 | 14.316 | 14.262 |
| coa.ind | 1 | 11.467 | 7.721 | 7.576 | 7.492 | 7.462 |
| nga.pow | 0 | 2.784 | 2.654 | 2.654 | 2.654 | 2.654 |
| nga.res | 0 | 4.634 | 2.331 | 2.331 | 2.331 | 2.331 |
| nga.ind | 0 | 5.336 | 5.087 | 5.087 | 5.087 | 5.087 |
| gso.all | 0 | 17.159 | 9.576 | 9.576 | 9.576 | 9.576 |
| die.all | 0 | 11.120 | 6.677 | 6.677 | 6.677 | 6.677 |
| lpg.all | 0 | 6.837 | 3.189 | 3.189 | 3.189 | 3.189 |
| ker.all | 0 | 10.986 | 6.469 | 6.469 | 6.469 | 6.469 |
| oop.all | 0 | 6.157 | 4.317 | 4.317 | 4.317 | 4.317 |
| bio.all | 1 | 14.471 | 14.471 | 14.471 | 14.471 | 14.471 |

| Year | Scenario 1 (ktoe) | Scenario 2 (ktoe) | Change |
|---|---|---|---|
| 2022 | 49,746 | 49,746 | +0.0% |
| 2023 | 62,504 | 62,504 | +0.0% |
| 2024 | 68,850 | 68,850 | +0.0% |
| 2026 | 72,187 | 72,187 | +0.0% |
| 2027 | 74,140 | 62,965 | -15.1% |
| 2030 | 80,994 | 68,791 | -15.1% |
| 2034 | 90,093 | 76,526 | -15.1% |

**Overall: PASS**
