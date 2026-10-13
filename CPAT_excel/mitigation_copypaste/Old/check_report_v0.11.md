# Check report - CPAT_Mitigation_CopyPaste_v0.11

## 1. Shipped file recalculated in LibreOffice (no LAMBDA support)
- error values: 2174, all in the 2040 (LAMBDA) columns: OK

## 2. Formula uniformity (relative R1C1)
- `sp` history 2022-2024: 72 cells, 1 formula OK
- `sp` projection 2025-2039: 360 cells, 1 formula OK
- `sp` 2040 (LAMBDA): 24 cells, 1 formula OK
- `sp` column D: 12 cells, 1 formula OK
- `sp` column E: 12 cells, 1 formula OK
- `sp` column F: 12 cells, 1 formula OK
- `sp` column G: 12 cells, 1 formula OK
  - `=INDEX(Inputs_prices!R5C20:R16C23,R[0]C6,MATCH(R4C[0],Inputs_prices!R4C20:R4C23,0))`
  - `=R[0]C4+(R[0]C[-1]-R[0]C4)*INDEX(R201C[0]:R204C[0],R[0]C5)/INDEX(R201C[-1]:R204C[-1],R[0]C5)`
  - 2040: `=SUPPLYCOST(R4C[0],Settings!R10C3,IFERROR(INDEX(Inputs_prices!R5C20:R16C23,R[0]C6,MATCH(R4C[0],Inputs_prices!R4C20:R4C23,0)),0),R[0]C4,R[0]C[-1],INDEX` OK
  - history and projection blocks without the history/projection IF (no reference to the last historical year): OK
- `txo` history 2022-2024: 72 cells, 1 formula OK
- `txo` projection 2025-2039: 360 cells, 1 formula OK
- `txo` 2040 (LAMBDA): 24 cells, 1 formula OK
- `txo` column D: 12 cells, 1 formula OK
- `txo` column E: 12 cells, 1 formula OK
- `txo` column F: 12 cells, 1 formula OK
- `txo` column G: 12 cells, 1 formula OK
  - `=INDEX(Inputs_prices!R5C24:R16C27,R[0]C7,MATCH(R4C[0],Inputs_prices!R4C24:R4C27,0))`
  - `=R[0]C4*R[0]C6+IF(R[0]C4*(1-R[0]C6)>=0,MAX(R[0]C4*(1-R[0]C6),(R[0]C5-R[-13]C[0]+R[0]C4*(1-R[0]C6))*(1-R[0]C6)),(R[0]C5-R[-13]C[0]+R[0]C4*(1-R[0]C6))*(`
  - 2040: `=OTHERTAX(R4C[0],Settings!R10C3,IFERROR(INDEX(Inputs_prices!R5C24:R16C27,R[0]C7,MATCH(R4C[0],Inputs_prices!R4C24:R4C27,0)),0),R[0]C4,R[0]C5,R[-13]C[0]` OK
  - history and projection blocks without the history/projection IF (no reference to the last historical year): OK
- `rpb` 2022-2039: 432 cells, 1 formula OK
- `rpb` 2040 (LAMBDA): 24 cells, 1 formula OK
- `rpb` column D: 12 cells, 1 formula OK
- `rpb` column E: 12 cells, 1 formula OK
- `rpb` column F: 12 cells, 1 formula OK
- `rpb` column G: 12 cells, 1 formula OK
  - `=(R[-26]C[0]+R[-13]C[0])*(1+R[0]C4)`
  - 2040: `=RETAILPRICE(R[-26]C[0],R[-13]C[0],R[0]C4)` OK
- `vat` 2022-2039: 432 cells, 1 formula OK
- `vat` 2040 (LAMBDA): 24 cells, 1 formula OK
- `vat` column D: 12 cells, 1 formula OK
- `vat` column E: 12 cells, 1 formula OK
- `vat` column F: 12 cells, 1 formula OK
- `vat` column G: 12 cells, 1 formula OK
  - `=(R[-39]C[0]+R[-26]C[0])*R[0]C4`
  - 2040: `=VATPAY(R[-39]C[0],R[-26]C[0],R[0]C4)` OK
- `etx` 2022-2039: 432 cells, 1 formula OK
- `etx` 2040 (LAMBDA): 24 cells, 1 formula OK
- `etx` column D: 12 cells, 1 formula OK
- `etx` column E: 12 cells, 1 formula OK
- `etx` column F: 12 cells, 1 formula OK
- `etx` column G: 12 cells, 1 formula OK
  - `=MAX(R[-39]C[0],0)`
  - 2040: `=TAXPART(R[-39]C[0])` OK
- `esub` 2022-2039: 432 cells, 1 formula OK
- `esub` 2040 (LAMBDA): 24 cells, 1 formula OK
- `esub` column D: 12 cells, 1 formula OK
- `esub` column E: 12 cells, 1 formula OK
- `esub` column F: 12 cells, 1 formula OK
- `esub` column G: 12 cells, 1 formula OK
  - `=MAX(-R[-52]C[0],0)`
  - 2040: `=SUBSIDYPART(R[-52]C[0])` OK
- `esubpu` 2022-2039: 432 cells, 1 formula OK
- `esubpu` 2040 (LAMBDA): 24 cells, 1 formula OK
- `esubpu` column D: 12 cells, 1 formula OK
- `esubpu` column E: 12 cells, 1 formula OK
- `esubpu` column F: 12 cells, 1 formula OK
- `esubpu` column G: 12 cells, 1 formula OK
  - `=R[-13]C[0]*R[0]C4`
  - 2040: `=PERUNIT(R[-13]C[0],R[0]C4)` OK
- `ctxnew` base year: 256 cells, 1 formula OK
- `ctxnew` 2023-2039: 4352 cells, 1 formula OK
- `ctxnew` 2040 (LAMBDA): 256 cells, 1 formula OK
- `ctxnew` column D: 128 cells, 1 formula OK
- `ctxnew` column E: 128 cells, 1 formula OK
- `ctxnew` column F: 128 cells, 1 formula OK
- `ctxnew` column G: 128 cells, 1 formula OK
  - `=R12C[0]*R[0]C4*INDEX(R21C[0]:R28C[0],R[0]C5)*INDEX(R30C[0]:R46C[0],R[0]C6)*(1-INDEX(R117C[0]:R133C[0],R[0]C6))`
  - `=R12C[0]*R[0]C4*INDEX(R21C[0]:R28C[0],R[0]C5)*INDEX(R30C[0]:R46C[0],R[0]C6)*(1-INDEX(R117C[0]:R133C[0],R[0]C6))`
  - 2040: `=CARBONTAX(R12C[0],R[0]C4,INDEX(R21C[0]:R28C[0],R[0]C5),INDEX(R30C[0]:R46C[0],R[0]C6),INDEX(R117C[0]:R133C[0],R[0]C6))` OK
- `ets` base year: 256 cells, 1 formula OK
- `ets` 2023-2039: 4352 cells, 1 formula OK
- `ets` 2040 (LAMBDA): 256 cells, 1 formula OK
- `ets` column D: 128 cells, 1 formula OK
- `ets` column E: 128 cells, 1 formula OK
- `ets` column F: 128 cells, 1 formula OK
- `ets` column G: 128 cells, 1 formula OK
  - `=R134C[0]*R[0]C4*INDEX(R117C[0]:R133C[0],R[0]C5)`
  - `=R134C[0]*R[0]C4*INDEX(R117C[0]:R133C[0],R[0]C5)`
  - 2040: `=ETSCOST(R134C[0],R[0]C4,INDEX(R117C[0]:R133C[0],R[0]C5))` OK
- `ntx` base year: 256 cells, 1 formula OK
- `ntx` 2023-2039: 4352 cells, 1 formula OK
- `ntx` 2040 (LAMBDA): 256 cells, 1 formula OK
- `ntx` column D: 128 cells, 1 formula OK
- `ntx` column E: 128 cells, 1 formula OK
- `ntx` column F: 128 cells, 1 formula OK
- `ntx` column G: 128 cells, 1 formula OK
  - `=INDEX(R78C[0]:R91C[0],R[0]C4)/R[0]C5`
  - `=INDEX(R78C[0]:R91C[0],R[0]C4)/R[0]C5`
  - 2040: `=NEWEXCISE(INDEX(R78C[0]:R91C[0],R[0]C4),R[0]C5)` OK
- `nce` base year: 256 cells, 1 formula OK
- `nce` 2023-2039: 4352 cells, 1 formula OK
- `nce` 2040 (LAMBDA): 256 cells, 1 formula OK
- `nce` column D: 128 cells, 1 formula OK
- `nce` column E: 128 cells, 1 formula OK
- `nce` column F: 128 cells, 1 formula OK
- `nce` column G: 128 cells, 1 formula OK
  - `=R[-27]C[0]+R[-18]C[0]+R[-9]C[0]`
  - `=R[-27]C[0]+R[-18]C[0]+R[-9]C[0]`
  - 2040: `=NEWPOLICY(R[-27]C[0],R[-18]C[0],R[-9]C[0])` OK
- `atp` base year: 256 cells, 1 formula OK
- `atp` 2023-2039: 4352 cells, 1 formula OK
- `atp` 2040 (LAMBDA): 256 cells, 1 formula OK
- `atp` column D: 128 cells, 1 formula OK
- `atp` column E: 128 cells, 1 formula OK
- `atp` column F: 128 cells, 1 formula OK
- `atp` column G: 128 cells, 1 formula OK
  - `=MAX(INDEX(R232C[0]:R243C[0],R[0]C4)+R[-9]C[0]*(1+R[0]C5),0.01)`
  - `=MAX(INDEX(R232C[0]:R243C[0],R[0]C4)+R[-9]C[0]*(1+R[0]C5),0.01)`
  - 2040: `=POSTTAX(INDEX(R232C[0]:R243C[0],R[0]C4),R[-9]C[0],R[0]C5)` OK
- `shp` base year: 256 cells, 1 formula OK
- `shp` 2023-2039: 4352 cells, 1 formula OK
- `shp` 2040 (LAMBDA): 256 cells, 1 formula OK
- `shp` column D: 128 cells, 1 formula OK
- `shp` column E: 128 cells, 1 formula OK
- `shp` column F: 128 cells, 1 formula OK
- `shp` column G: 128 cells, 1 formula OK
  - `=INDEX(R175C[0]:R178C[0],R[0]C4)*R[0]C5*INDEX(R179C[0]:R195C[0],R[0]C6)`
  - `=INDEX(R175C[0]:R178C[0],R[0]C4)*R[0]C5*INDEX(R179C[0]:R195C[0],R[0]C6)`
  - 2040: `=SHADOWEFF(INDEX(R175C[0]:R178C[0],R[0]C4),R[0]C5,INDEX(R179C[0]:R195C[0],R[0]C6))` OK
- `ener` base year: 256 cells, 1 formula OK
- `ener` 2023-2039: 4352 cells, 1 formula OK
- `ener` 2040 (LAMBDA): 256 cells, 1 formula OK
- `ener` column D: 128 cells, 1 formula OK
- `ener` column E: 128 cells, 1 formula OK
- `ener` column F: 128 cells, 1 formula OK
- `ener` column G: 128 cells, 1 formula OK
  - `=INDEX(Inputs!R5C21:R132C21,MATCH(R[0]C2&"|"&R[0]C3,Inputs!R5C1:R132C1,0))`
  - `=R[0]C[-1]*(1/(1+R[0]C7))^(1+R[0]C5)*(1+R7C[0])^R[0]C4*(R[-18]C[0]/R[-18]C[-1])^R[0]C5*((R[-18]C[0]+R[-9]C[0])/(R[-18]C[-1]+R[-9]C[-1]))^(R[0]C6*(1+R[`
  - 2040: `=FUELUSE(R[0]C[-1],R[-18]C[0],R[-18]C[-1],R[-9]C[0],R[-9]C[-1],R7C[0],R[0]C4,R[0]C5,R[0]C6,R[0]C7)` OK
- Section 1 and 2 MTInputs rows: 4256 cells, 1 formula OK
- section 2 international prices `gp` (data rows, plain): 152 cells, 1 formula OK
- index row `infl`: 38 cells, 1 formula OK
- index row `defl`: 38 cells, 1 formula OK
- Section 1 `cptraj*` plain: 36 cells, 1 formula OK
- Section 1 `cptraj*` 2040 (LAMBDA): 2 cells, 1 formula OK
- Section 1 `ets.p*` plain: 36 cells, 1 formula OK
- Section 1 `ets.p*` 2040 (LAMBDA): 2 cells, 1 formula OK
- Section 1 `ets.a*` plain: 36 cells, 1 formula OK
- Section 1 `ets.a*` 2040 (LAMBDA): 2 cells, 1 formula OK
- Section 1 `fpr.*` plain: 504 cells, 1 formula OK
- Section 1 `fpr.*` 2040 (LAMBDA): 28 cells, 1 formula OK
- Section 1 `fb.*` plain: 144 cells, 1 formula OK
- Section 1 `fb.*` 2040 (LAMBDA): 8 cells, 1 formula OK
- Section 1 `etsc.*` plain: 612 cells, 1 formula OK
- Section 1 `etsc.*` 2040 (LAMBDA): 34 cells, 1 formula OK
- Section 1 `shps.*` plain: 144 cells, 1 formula OK
- Section 1 `shps.*` 2040 (LAMBDA): 8 cells, 1 formula OK
- Section 1 `ssc.*` plain: 612 cells, 1 formula OK
- Section 1 `ssc.*` 2040 (LAMBDA): 34 cells, 1 formula OK
- subsector heading totals: 608 cells, 1 formula OK
- right column (2040): 1039 rows call a LAMBDA = all calculated rows (1039) OK
- parameter columns D:G grouped and hidden: OK

## 3. LAMBDA encoding
- PATH: 6 parameters, 20 _xlpm. tokens, bare names [] OK
- ETSPRICE: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- ETSCOVER: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- SHADOWSECTOR: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- SHADOWSHARE: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- SUPPLYCOST: 7 parameters, 15 _xlpm. tokens, bare names [] OK
- OTHERTAX: 7 parameters, 26 _xlpm. tokens, bare names [] OK
- RETAILPRICE: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- VATPAY: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- TAXPART: 1 parameters, 2 _xlpm. tokens, bare names [] OK
- SUBSIDYPART: 1 parameters, 2 _xlpm. tokens, bare names [] OK
- PERUNIT: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- CARBONTAX: 5 parameters, 10 _xlpm. tokens, bare names [] OK
- ETSCOST: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- NEWEXCISE: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- NEWPOLICY: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- POSTTAX: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- SHADOWEFF: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- FUELUSE: 10 parameters, 24 _xlpm. tokens, bare names [] OK

## 4. Values vs independent Python recomputation
- A. LAMBDA expanded, scenario 1: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 4.53e-15 (at S1261) OK
- B. plain formula dragged, scenario 1: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 4.53e-15 (at S1261) OK
- C. LAMBDA copied back over the whole row, scenario 1: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 4.53e-15 (at S1261) OK
- A. LAMBDA expanded, scenario 2: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 4.46e-15 (at AW216) OK
- B. plain formula dragged, scenario 2: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 4.46e-15 (at AW216) OK
- C. LAMBDA copied back over the whole row, scenario 2: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 4.46e-15 (at AW216) OK

## 5. Scenario tests (copied MTInputs column + Mitigation group)
- scenario 3 (copy of 2) = scenario 2: 22933 numeric cells, max abs diff 0 OK
- scenario 4 (copy of 2, carbon price 0 in MTInputs) = scenario 1: 22895 numeric cells, max abs diff 0 OK
- pasted group 3: number 3, name "Test 3: copy of 2", code `egy.mit.ener.rod.coa.e.3` OK
- pasted group 4: number 4, name "Test 4: copy of 2, carbon price 0", code `egy.mit.ener.rod.coa.e.4` OK
- pasted group 5: number 5, name "Test 5: legacy Egypt carbon tax", code `egy.mit.ener.rod.coa.e.5` OK
- pasted group 6: number 6, name "Test 6: fuel price reform", code `egy.mit.ener.rod.coa.e.6` OK
- pasted group 7: number 7, name "Test 7: feebates", code `egy.mit.ener.rod.coa.e.7` OK
- pasted group 8: number 8, name "Test 8: IMF-WB*, High, nominal carbon price", code `egy.mit.ener.rod.coa.e.8` OK
- pasted group 9: number 9, name "Test 9: new ETS (industry, power) + carbon tax elsewhere", code `egy.mit.ener.rod.coa.e.9` OK
- scenario 5 (legacy Egypt carbon tax): carbon price [0, 12.5, 50, 112.5] vs legacy row 8582, max abs diff 0 OK
- scenario 5 vs Python: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 5.18e-15 (at DF618) OK
- scenario 6 vs Python: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 4.87e-15 (at DN360) OK
- scenario 7 vs Python: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 4.62e-15 (at EN1261) OK
- scenario 8 vs Python: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 4.45e-15 (at FC1133) OK
- scenario 9 vs Python: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 4.46e-15 (at GG216) OK
  - scenario 6, road gasoline new excise ($/GJ) 2026-2031: [0, 1.43, 2.86, 4.289, 5.719, 7.149]
  - scenario 7, road gasoline feebate shadow price ($/GJ) 2026-2031: [0, 0.662, 1.544, 2.427, 3.309, 4.191]
- scenario 7 vs scenario 1 (covered with a non-zero rate: cem, cst, ftr, irn, mac, mch, nfm, oen, omn, rod): other subsectors unchanged (max abs diff 0); covered subsectors in 2039: cem -16.87%, ftr -17.72%, irn -15.62%, oen -19.82%, omn -19.86%, rod -18.60% OK
  - scenario 8: crude oil gp 2024-2026 (real) [84.95, 109.03, 114.14], carbon price 2027-2030 (nominal 20 x infl) [19.581, 19.162, 18.752, 18.35]
- scenario 9 (new ETS on power and industry from 2027, price = carbon price): nce, atp and fuel use equal scenario 2 (max abs diff 0); ETS cost > 0 (40.68, sum over rows 2039); carbon tax in cement 0 (0) OK
  - scenario 9 auction share 2026-2036: [0, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1, 1]
- price controls None (global), scenario 1 vs Python: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 2.86e-14 (at AD359) OK
- price controls None (global), scenario 2 vs Python: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 1.23e-14 (at AR1136) OK
- price controls Manual (global), scenario 1 vs Python: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 5.99e-15 (at T251) OK
- price controls Manual (global), scenario 2 vs Python: carbon and ETS price, ETS auction share, gp, 7 price-fuel and 7 subsector variables, 18753 cells, max relative diff 5.99e-15 (at AN251) OK

## 6. Regression vs v0.10 (shared output codes)
- 1902 output codes in both versions and not affected by the food & forestry elasticities, 2022-2035: max abs diff 0 OK
- intended changes: 44 fuel-use codes of food & forestry and its aggregates; variables added: D_ETSAuctCont, D_ETSAuctStart, D_ETSAuctTarget, D_ETSIntro, D_ETSOutro, D_NewETS, esub, esubpu, ets, etsc, etscov, etx, reg, vat; years 2036-2040 added

| Output code | v0.10 2027 | v0.11 2027 | v0.10 2035 | v0.11 2035 | v0.11 2040 |
|---|---|---|---|---|---|
| `egy.mit.ener.foo.all.e.1` | 1,739.8 | 1,441.0 | 2,079.9 | 1,709.1 | 1,880.7 |
| `egy.mit.ener.all.all.e.1` | 74,139.6 | 73,840.8 | 92,462.0 | 92,091.2 | 104,439.9 |
| `egy.mit.ener.all.all.e.2` | 62,965.2 | 62,824.8 | 78,540.3 | 78,360.8 | 88,873.8 |

## 7. Labels and codes
- column H: 1212 cells, 1 formula OK
- column I: 1212 cells, 1 formula OK
- column J: 1212 cells, 1 formula OK
- code column K (scenario 1): 1212 cells, 1 formula OK
- code column AE (scenario 2): 1212 cells, 1 formula OK
- K360: `egy.mit.ener.rod.gso.e.1` | Fuel use | Road | Gasoline OK
- K206: `egy.mit.sp.pow.coa.a.1` | Supply cost (pre-tax price) | Power | Coal OK
- AE238: `egy.mit.rpb.all.gso.a.2` | Retail price before new policies | All subsectors | Gasoline OK
- K201: `egy.mit.gp.int.oil.1` | International energy price (real; source and adjustment from MTInputs) | International | Crude oil OK
- K8: `egy.mit.infl.1` | Inflation index: US CPI, ResultsYear = 1 (nominal -> real for domestic prices and nominal policy inputs) OK
- AE314: `egy.mit.ets.rod.nga.a.2` | New ETS permit cost | Road | Natural gas OK
- K278: `egy.mit.esub.all.die.a.1` | Existing consumer subsidy (negative part of txo, positive number) | All subsectors | Diesel OK
- AE134: `egy.mit.ets.p.2` | ETS permit price (= carbon price inputs once the new ETS applies) OK
- AE342: `egy.mit.atp.rod.gso.e.2` | After-tax price | Road | Gasoline OK
- AE307: `egy.mit.ctxnew.rod.die.a.2` | New carbon tax | Road | Diesel OK
- K351: `egy.mit.shp.rod.gso.1` | Shadow price on the efficiency margin | Road | Gasoline OK
- K303: `egy.mit.ener.rod.all.e.1` | Fuel use | Road | All fuels OK
- K302: `egy.mit.ener.tra.all.e.1` | Fuel use | Transport | All fuels OK
- AE12: `egy.mit.cptraj.2` | Carbon price trajectory used OK
- K46: `egy.mit.ctcov.oen.all.1` | Carbon tax coverage (Apply tax?) | Other energy use | All fuels OK
- AE1338: `egy.mit.ener.all.all.e.2` | Fuel use | All subsectors | All fuels OK

## 8. Row outline
- carbon price (summary of section 1): 1 rows at level 0, visible OK
- section 1 inputs and paths: 174 rows at level 1, hidden OK
- section 2 retail prices before new policies: 12 rows at level 0, visible OK
- section 2 selectors, gp, sp, txo, vat, etx, esub, esubpu: 78 rows at level 1, hidden OK
- sector totals: 4 rows at level 0, visible OK
- subsector headings: 16 rows at level 1, visible OK
- subsector variables: 896 rows at level 2, hidden OK
- results: 30 rows at level 1, visible OK

## 9. MTInputs
- columns A:H, rows 1-415: 0 cells differ from the template; D:E hidden OK

## 10. Format
- section bands ['Scenario assumptions (', '1. Policies (scenario ', '2. Retail energy price', '3. Power sector (elast', '5. Transport sector', '6. Buildings sector', '7. Industrial sector', '8. Other energy use', '11. Results - energy c', '12. Revenues (next ste']: white bold text OK
- sheet LegacyDiff (second tab): 21 differences listed OK
- Inputs_prices: data changed by assumption marked bright yellow (columns ['Margin (mit.mar, base year, nominal pric', 'Raw pass-through coefficient (mit.ps, ba', 'VAT applies to final consumers? (assumpt', 'VAT rate assumption = general VAT rate (']; other oil products pass-through and margin) OK

## 11. Sanity and results
- scenario 1 vs 2 before 2027 (fuel use): max abs diff 0 OK
- other energy use not taxed (MCovOen FALSE): fuel use scenario 1 = 2, max abs diff 0 OK
- check row 1366: max abs 0; sector totals sum to total: max abs 1.6e-10 OK
- pass-through 0 (nga.pow, nga.res, nga.ind, gso.all, die.all, lpg.all, ker.all, oop.all): retail price before new policies stays at its 2024 value, max abs diff 0 OK

Retail price before new policies, scenario 1 (real $/GJ of ResultsYear), and chosen pass-through:

| Price fuel | pass-through | 2022 | 2024 | 2027 | 2030 | 2040 | VAT 2024 | existing tax 2024 | existing subsidy 2024 | subsidy per price unit 2024 |
|---|---|---|---|---|---|---|---|---|---|---|
| coa.pow | 1 | 13.324 | 5.912 | 5.677 | 5.540 | 5.492 | 0.000 | 0.000 | 0.000 | 0.000 |
| coa.res | 1 | 23.189 | 14.740 | 14.472 | 14.316 | 14.262 | 1.810 | 0.000 | 0.000 | 0.000 |
| coa.ind | 1 | 11.467 | 7.721 | 7.576 | 7.492 | 7.462 | 0.000 | 0.166 | 0.000 | 0.000 |
| nga.pow | 0 | 2.784 | 2.654 | 2.654 | 2.654 | 2.654 | 0.000 | 0.000 | 6.206 | 6.206 |
| nga.res | 0 | 4.634 | 2.331 | 2.331 | 2.331 | 2.331 | 0.286 | 0.000 | 12.434 | 12.434 |
| nga.ind | 0 | 5.336 | 5.087 | 5.087 | 5.087 | 5.087 | 0.000 | 0.000 | 4.576 | 4.576 |
| gso.all | 0 | 17.159 | 9.576 | 9.576 | 9.576 | 9.576 | 1.176 | 0.000 | 13.434 | 0.470 |
| die.all | 0 | 11.120 | 6.677 | 6.677 | 6.677 | 6.677 | 0.820 | 0.000 | 15.013 | 0.560 |
| lpg.all | 0 | 6.837 | 3.189 | 3.189 | 3.189 | 3.189 | 0.392 | 0.000 | 18.182 | 0.469 |
| ker.all | 0 | 10.986 | 6.469 | 6.469 | 6.469 | 6.469 | 0.794 | 0.000 | 12.642 | 0.477 |
| oop.all | 0 | 6.157 | 4.317 | 4.317 | 4.317 | 4.317 | 0.530 | 0.000 | 11.182 | 68.431 |
| bio.all | 1 | 14.471 | 14.471 | 14.471 | 14.471 | 14.471 | 0.000 | 0.000 | 0.000 | 0.000 |

| Year | Scenario 1 (ktoe) | Scenario 2 (ktoe) | Change |
|---|---|---|---|
| 2022 | 49,746 | 49,746 | +0.0% |
| 2023 | 62,317 | 62,317 | +0.0% |
| 2024 | 68,573 | 68,573 | +0.0% |
| 2026 | 71,896 | 71,896 | +0.0% |
| 2027 | 73,841 | 62,825 | -14.9% |
| 2030 | 80,669 | 68,636 | -14.9% |
| 2035 | 92,091 | 78,361 | -14.9% |
| 2040 | 104,440 | 88,874 | -14.9% |

**Overall: PASS**
