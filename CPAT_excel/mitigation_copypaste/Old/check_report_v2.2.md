# Check report - CPAT-AI-Mitigation-MVP-v2.2

## 0. Formula length (at most 100 characters per cell formula)
- 133331 cell formulas in 30 sheets (0 call a named LAMBDA): longest 99 characters (Mitigation!AA875); over 100: 0 OK
- 87 named LAMBDAs (longest: OTHERTAX, 299 characters); 68 defined names for data tables and fixed ranges: WEOVal, WEOKey, WEOYear, ElaVal, ElaKey, ElaGroup, EFVal, EFKey, ConsVal, ConsKey, PIntVal, PIntKey, PIntYear, PDomVal, PDomKey, PDomHdr, PPathVal, PPathKey, PPathYear, PTechVal, PTechKey, PTechHdr, PParVal, PParKey, VarKey, VarLab, VarUnit, VarSfx, VarSrc, VarEnd, SubKey, FuelKey, SecKey, SecLab, FuelKeyL, FuelLab, MemKey, MemLab, SecVar, SecFuel, SecSub, SemiEl, MitAll, MitScen, MitDesc, MitUnit, SRScen, SRName, SRCode, SRDesc, SRUnit, SRVal, SRYear, CodePrefix, CPIRes, DeflRes, PDomBase, PDomLast, OBREff, ECPKey, ECPCov, ECPVal, ECPYear, ECPLast, ECPAllKey, ECPAllVal, ECPAllYear, CtryCode

## 1. Shipped file recalculated in LibreOffice (no LAMBDA support)
- no cell calls a LAMBDA: OK
- shipped file recalculated by LibreOffice (no LAMBDA support): error values in all 30 sheets: none OK

## 2. Formula uniformity (relative R1C1)
- `sp` history 2022-2024: 72 cells, 1 formula OK
- `sp` projection 2025-2040: 384 cells, 1 formula OK
- `sp` column D: 12 cells, 1 formula OK
- `sp` column E: 12 cells, 1 formula OK
- `sp` column F: 12 cells, 1 formula OK
- `sp` column G: 12 cells, 1 formula OK
  - `=INDEX(Inputs_prices!R7C21:R18C24,R[0]C6,MATCH(R4C[0],Inputs_prices!R6C21:R6C24,0))`
  - `=R[0]C4+(R[0]C[-1]-R[0]C4)*INDEX(R337C[0]:R340C[0],R[0]C5)/INDEX(R337C[-1]:R340C[-1],R[0]C5)`
  - history and projection blocks without the history/projection IF (no reference to the last historical year): OK
- `txf` 2022-2040: 456 cells, 1 formula OK
- `txf` column D: 12 cells, 1 formula OK
- `txf` column E: 12 cells, 1 formula OK
- `txf` column F: 12 cells, 1 formula OK
- `txf` column G: 12 cells, 1 formula OK
  - `=(R[0]C5-R[-13]C[0]+R[0]C4*(1-R[0]C6))*(1-R[0]C6)`
- `txo` history 2022-2024: 72 cells, 1 formula OK
- `txo` projection 2025-2040: 384 cells, 1 formula OK
- `txo` column D: 12 cells, 1 formula OK
- `txo` column E: 12 cells, 1 formula OK
- `txo` column F: 12 cells, 1 formula OK
- `txo` column G: 12 cells, 1 formula OK
  - `=INDEX(Inputs_prices!R7C25:R18C28,R[0]C7,MATCH(R4C[0],Inputs_prices!R6C25:R6C28,0))`
  - `=R[0]C4*R[0]C6+IF(R[0]C4*(1-R[0]C6)>=0,MAX(R[0]C4*(1-R[0]C6),R[-13]C[0]),R[-13]C[0])`
  - history and projection blocks without the history/projection IF (no reference to the last historical year): OK
- `rpb` 2022-2040: 456 cells, 1 formula OK
- `rpb` column D: 12 cells, 1 formula OK
- `rpb` column E: 12 cells, 1 formula OK
- `rpb` column F: 12 cells, 1 formula OK
- `rpb` column G: 12 cells, 1 formula OK
  - `=(R[-39]C[0]+R[-13]C[0])*(1+R[0]C4)`
- `vat` 2022-2040: 456 cells, 1 formula OK
- `vat` column D: 12 cells, 1 formula OK
- `vat` column E: 12 cells, 1 formula OK
- `vat` column F: 12 cells, 1 formula OK
- `vat` column G: 12 cells, 1 formula OK
  - `=(R[-52]C[0]+R[-26]C[0])*R[0]C4`
- `etx` 2022-2040: 456 cells, 1 formula OK
- `etx` column D: 12 cells, 1 formula OK
- `etx` column E: 12 cells, 1 formula OK
- `etx` column F: 12 cells, 1 formula OK
- `etx` column G: 12 cells, 1 formula OK
  - `=MAX(R[-39]C[0],0)`
- `esub` 2022-2040: 456 cells, 1 formula OK
- `esub` column D: 12 cells, 1 formula OK
- `esub` column E: 12 cells, 1 formula OK
- `esub` column F: 12 cells, 1 formula OK
- `esub` column G: 12 cells, 1 formula OK
  - `=MAX(-R[-52]C[0],0)`
- `esubpu` 2022-2040: 456 cells, 1 formula OK
- `esubpu` column D: 12 cells, 1 formula OK
- `esubpu` column E: 12 cells, 1 formula OK
- `esubpu` column F: 12 cells, 1 formula OK
- `esubpu` column G: 12 cells, 1 formula OK
  - `=R[-13]C[0]*R[0]C4`
- `ctxnew` base year: 256 cells, 1 formula OK
- `ctxnew` 2023-2040: 4608 cells, 1 formula OK
- `ctxnew` column D: 128 cells, 1 formula OK
- `ctxnew` column E: 128 cells, 1 formula OK
- `ctxnew` column F: 128 cells, 1 formula OK
- `ctxnew` column G: 128 cells, 1 formula OK
  - `=INDEX(R263C[0]:R279C[0],R[0]C6)*R[0]C4*INDEX(R23C[0]:R30C[0],R[0]C5)`
  - `=INDEX(R263C[0]:R279C[0],R[0]C6)*R[0]C4*INDEX(R23C[0]:R30C[0],R[0]C5)`
- `ets` base year: 256 cells, 1 formula OK
- `ets` 2023-2040: 4608 cells, 1 formula OK
- `ets` column D: 128 cells, 1 formula OK
- `ets` column E: 128 cells, 1 formula OK
- `ets` column F: 128 cells, 1 formula OK
- `ets` column G: 128 cells, 1 formula OK
  - `=INDEX(R280C[0]:R296C[0],R[0]C5)*R[0]C4`
  - `=INDEX(R280C[0]:R296C[0],R[0]C5)*R[0]C4`
- `ntx` base year: 256 cells, 1 formula OK
- `ntx` 2023-2040: 4608 cells, 1 formula OK
- `ntx` column D: 128 cells, 1 formula OK
- `ntx` column E: 128 cells, 1 formula OK
- `ntx` column F: 128 cells, 1 formula OK
- `ntx` column G: 128 cells, 1 formula OK
  - `=INDEX(R80C[0]:R93C[0],R[0]C4)/R[0]C5`
  - `=INDEX(R80C[0]:R93C[0],R[0]C4)/R[0]C5`
- `nce` base year: 256 cells, 1 formula OK
- `nce` 2023-2040: 4608 cells, 1 formula OK
- `nce` column D: 128 cells, 1 formula OK
- `nce` column E: 128 cells, 1 formula OK
- `nce` column F: 128 cells, 1 formula OK
- `nce` column G: 128 cells, 1 formula OK
  - `=R[-27]C[0]+R[-18]C[0]+R[-9]C[0]`
  - `=R[-27]C[0]+R[-18]C[0]+R[-9]C[0]`
- `xcp` base year: 256 cells, 1 formula OK
- `xcp` 2023-2040: 4608 cells, 1 formula OK
- `xcp` column D: 128 cells, 1 formula OK
- `xcp` column E: 128 cells, 1 formula OK
- `xcp` column F: 128 cells, 1 formula OK
- `xcp` column G: 128 cells, 1 formula OK
  - `=R[0]C4*(R[0]C5*R321C[0]+R[0]C6*R324C[0])`
  - `=R[0]C4*(R[0]C5*R321C[0]+R[0]C6*R324C[0])`
- `xcpd` base year: 256 cells, 1 formula OK
- `xcpd` 2023-2040: 4608 cells, 1 formula OK
- `xcpd` column D: 128 cells, 1 formula OK
- `xcpd` column E: 128 cells, 1 formula OK
- `xcpd` column F: 128 cells, 1 formula OK
- `xcpd` column G: 128 cells, 1 formula OK
  - `=R[0]C4*(R[0]C5*(R321C[0]-R322C[0])+R[0]C6*(R324C[0]-R325C[0]))`
  - `=R[0]C4*(R[0]C5*(R321C[0]-R322C[0])+R[0]C6*(R324C[0]-R325C[0]))`
- `atp` base year: 256 cells, 1 formula OK
- `atp` 2023-2040: 4608 cells, 1 formula OK
- `atp` column D: 128 cells, 1 formula OK
- `atp` column E: 128 cells, 1 formula OK
- `atp` column F: 128 cells, 1 formula OK
- `atp` column G: 128 cells, 1 formula OK
  - `=MAX(INDEX(R381C[0]:R392C[0],R[0]C4)+(R[-27]C[0]+R[-9]C[0])*(1+R[0]C5),0.01)`
  - `=MAX(INDEX(R381C[0]:R392C[0],R[0]C4)+(R[-27]C[0]+R[-9]C[0])*(1+R[0]C5),0.01)`
- `shp` base year: 256 cells, 1 formula OK
- `shp` 2023-2040: 4608 cells, 1 formula OK
- `shp` column D: 128 cells, 1 formula OK
- `shp` column E: 128 cells, 1 formula OK
- `shp` column F: 128 cells, 1 formula OK
- `shp` column G: 128 cells, 1 formula OK
  - `=INDEX(R297C[0]:R313C[0],R[0]C5)*R[0]C4`
  - `=INDEX(R297C[0]:R313C[0],R[0]C5)*R[0]C4`
- `ener` base year: 256 cells, 1 formula OK
- `ener` 2023-2040: 4608 cells, 1 formula OK
- `ener` column D: 128 cells, 1 formula OK
- `ener` column E: 128 cells, 1 formula OK
- `ener` column F: 128 cells, 1 formula OK
- `ener` column G: 128 cells, 1 formula OK
  - `=INDEX(Inputs!R5C24:R132C24,MATCH(R[0]C2&"|"&R[0]C3,Inputs!R5C1:R132C1,0))`
  - `=R[0]C[-1]*R[0]C7*(1+R7C[0])^R[0]C4*(R[-18]C[0]/R[-18]C[-1])^R[0]C5*((R[-18]C[0]+R[-9]C[0])/(R[-18]C[-1]+R[-9]C[-1]))^R[0]C6`
- `co2` base year: 256 cells, 1 formula OK
- `co2` 2023-2040: 4608 cells, 1 formula OK
- `co2` column D: 128 cells, 1 formula OK
- `co2` column E: 128 cells, 1 formula OK
- `co2` column F: 128 cells, 1 formula OK
- `co2` column G: 128 cells, 1 formula OK
  - `=R[-9]C[0]*Settings!R13C3*R[0]C4`
  - `=R[-9]C[0]*Settings!R13C3*R[0]C4`
- `rtx` base year: 256 cells, 1 formula OK
- `rtx` 2023-2040: 4608 cells, 1 formula OK
- `rtx` column D: 128 cells, 1 formula OK
- `rtx` column E: 128 cells, 1 formula OK
- `rtx` column F: 128 cells, 1 formula OK
- `rtx` column G: 128 cells, 1 formula OK
  - `=R[-18]C[0]*Settings!R13C3*(INDEX(R407C[0]:R418C[0],R[0]C4)+INDEX(R394C[0]:R405C[0],R[0]C4)+R[-45]C[0]*(1+R[0]C5))`
  - `=R[-18]C[0]*Settings!R13C3*(INDEX(R407C[0]:R418C[0],R[0]C4)+INDEX(R394C[0]:R405C[0],R[0]C4)+R[-45]C[0]*(1+R[0]C5))`
- `rsub` base year: 256 cells, 1 formula OK
- `rsub` 2023-2040: 4608 cells, 1 formula OK
- `rsub` column D: 128 cells, 1 formula OK
- `rsub` column E: 128 cells, 1 formula OK
- `rsub` column F: 128 cells, 1 formula OK
- `rsub` column G: 128 cells, 1 formula OK
  - `=R[-27]C[0]*Settings!R13C3*(INDEX(R420C[0]:R431C[0],R[0]C4))`
  - `=R[-27]C[0]*Settings!R13C3*(INDEX(R420C[0]:R431C[0],R[0]C4))`
- `rnew` base year: 256 cells, 1 formula OK
- `rnew` 2023-2040: 4608 cells, 1 formula OK
- `rnew` column D: 128 cells, 1 formula OK
- `rnew` column E: 128 cells, 1 formula OK
- `rnew` column F: 128 cells, 1 formula OK
- `rnew` column G: 128 cells, 1 formula OK
  - `=R[-36]C[0]*Settings!R13C3*(R[-108]C[0]+R[-90]C[0]+R[-99]C[0]*R201C[0]+R[-81]C[0]*R[0]C4)`
  - `=R[-36]C[0]*Settings!R13C3*(R[-108]C[0]+R[-90]C[0]+R[-99]C[0]*R201C[0]+R[-81]C[0]*R[0]C4)`
- `rxcp` base year: 256 cells, 1 formula OK
- `rxcp` 2023-2040: 4608 cells, 1 formula OK
- `rxcp` column D: 128 cells, 1 formula OK
- `rxcp` column E: 128 cells, 1 formula OK
- `rxcp` column F: 128 cells, 1 formula OK
- `rxcp` column G: 128 cells, 1 formula OK
  - `=R[-45]C[0]*Settings!R13C3*(R[-81]C[0])`
  - `=R[-45]C[0]*Settings!R13C3*(R[-81]C[0])`
- Section 1 and 2 MTInputs rows: 4598 cells, 1 formula OK
- section 2 international prices `gp` (data rows, plain): 152 cells, 1 formula OK
- section 2 international prices before the adjustment `gpn` (data rows, plain): 152 cells, 1 formula OK
- section 2 `gp.src` (plain): 38 cells, 1 formula OK
- section 2 `gp.adj` (plain): 38 cells, 1 formula OK
- electricity demand memo lines in the sector blocks: 608 cells, 1 formula OK
- index row `infl`: 38 cells, 1 formula OK
- index row `defl`: 38 cells, 1 formula OK
- Section 1 `cptraj*` plain: 38 cells, 1 formula OK
- Section 1 `cp.t*` plain: 38 cells, 1 formula OK
- Section 1 `cp.nom*` plain: 38 cells, 1 formula OK
- Section 1 `ets.on*` plain: 38 cells, 1 formula OK
- Section 1 `ets.vol*` plain: 38 cells, 1 formula OK
- Section 1 `ets.imp*` plain: 38 cells, 1 formula OK
- Section 1 `ets.vadj*` plain: 38 cells, 1 formula OK
- Section 1 `ets.bce*` plain: 38 cells, 1 formula OK
- Section 1 `ets.chg*` plain: 38 cells, 1 formula OK
- Section 1 `ets.cap*` plain: 38 cells, 1 formula OK
- Section 1 `ets.se*` plain: 38 cells, 1 formula OK
- Section 1 `ets.est*` plain: 38 cells, 1 formula OK
- Section 1 `ets.p*` plain: 38 cells, 1 formula OK
- Section 1 `ets.pe*` plain: 38 cells, 1 formula OK
- Section 1 `ets.rf*` plain: 38 cells, 1 formula OK
- Section 1 `fpr.*` plain: 532 cells, 1 formula OK
- Section 1 `fb.*` plain: 152 cells, 1 formula OK
- Section 1 `etsc.*` plain: 646 cells, 1 formula OK
- Section 1 `etsb.*` plain: 152 cells, 1 formula OK
- Section 1 `obr.*` plain: 646 cells, 1 formula OK
- Section 1 `shps.*` plain: 152 cells, 1 formula OK
- Section 1 `ssc.*` plain: 646 cells, 1 formula OK
- Section 1 `ctr.*` plain: 646 cells, 1 formula OK
- Section 1 `etr.*` plain: 646 cells, 1 formula OK
- Section 1 `shpx.*` plain: 646 cells, 1 formula OK
- subsector heading totals: 608 cells, 1 formula OK
- power `tcf` 2022-2040 (data, sum or look-ahead: one formula): 342 cells, 1 formula OK
  - `=INDEX(PPathVal,MATCH(R[0]C4,PPathKey,0),MATCH(R4C[0],PPathYear,0))`
- power `tcf` column D: 9 cells, 1 formula OK
- power `tcf` column E: 9 cells, 1 formula OK
- power `tcf` column F: 9 cells, 1 formula OK
- power `tcf` column G: 9 cells, 1 formula OK
- power `cbat` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=INDEX(PPathVal,MATCH(R[0]C4,PPathKey,0),MATCH(R4C[0],PPathYear,0))`
- power `cbat` column D: 1 cells, 1 formula OK
- power `cbat` column E: 1 cells, 1 formula OK
- power `cbat` column F: 1 cells, 1 formula OK
- power `cbat` column G: 1 cells, 1 formula OK
- power `cint` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=INDEX(PPathVal,MATCH(R[0]C4,PPathKey,0),MATCH(R4C[0],PPathYear,0))`
- power `cint` column D: 1 cells, 1 formula OK
- power `cint` column E: 1 cells, 1 formula OK
- power `cint` column F: 1 cells, 1 formula OK
- power `cint` column G: 1 cells, 1 formula OK
- power `obat` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=INDEX(PPathVal,MATCH(R[0]C4,PPathKey,0),MATCH(R4C[0],PPathYear,0))`
- power `obat` column D: 1 cells, 1 formula OK
- power `obat` column E: 1 cells, 1 formula OK
- power `obat` column F: 1 cells, 1 formula OK
- power `obat` column G: 1 cells, 1 formula OK
- power `cel` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=INDEX(PPathVal,MATCH(R[0]C4,PPathKey,0),MATCH(R4C[0],PPathYear,0))`
- power `cel` column D: 1 cells, 1 formula OK
- power `cel` column E: 1 cells, 1 formula OK
- power `cel` column F: 1 cells, 1 formula OK
- power `cel` column G: 1 cells, 1 formula OK
- power `incf` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=INDEX(PPathVal,MATCH(R[0]C4,PPathKey,0),MATCH(R4C[0],PPathYear,0))`
- power `incf` column D: 1 cells, 1 formula OK
- power `incf` column E: 1 cells, 1 formula OK
- power `incf` column F: 1 cells, 1 formula OK
- power `incf` column G: 1 cells, 1 formula OK
- power `shock` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=INDEX(PPathVal,MATCH(R[0]C4,PPathKey,0),MATCH(R4C[0],PPathYear,0))`
- power `shock` column D: 1 cells, 1 formula OK
- power `shock` column E: 1 cells, 1 formula OK
- power `shock` column F: 1 cells, 1 formula OK
- power `shock` column G: 1 cells, 1 formula OK
- power `cax` 2022-2040: 342 cells, 1 formula OK
  - `=R[0]C4*R[-18]C[0]`
- power `cax` column D: 9 cells, 1 formula OK
- power `cax` column E: 9 cells, 1 formula OK
- power `cax` column F: 9 cells, 1 formula OK
- power `cax` column G: 9 cells, 1 formula OK
- power `caxav` base year: 18 cells, 1 formula OK
  - `=R[-10]C[0]`
- power `caxav` 2023-2040: 324 cells, 1 formula OK
  - `=R[0]C[-1]*(1-R[442]C[-1])+R[-10]C[-1]*R[442]C[-1]`
- power `caxav` column D: 9 cells, 1 formula OK
- power `caxav` column E: 9 cells, 1 formula OK
- power `caxav` column F: 9 cells, 1 formula OK
- power `caxav` column G: 9 cells, 1 formula OK
- power `stoav` base year: 18 cells, 1 formula OK
  - `=R[0]C4*R749C[0]`
- power `stoav` 2023-2040: 324 cells, 1 formula OK
  - `=R[0]C[-1]*(1-R[432]C[-1])+R[0]C4*R749C[0]*R[432]C[-1]`
- power `stoav` column D: 9 cells, 1 formula OK
- power `stoav` column E: 9 cells, 1 formula OK
- power `stoav` column F: 9 cells, 1 formula OK
- power `stoav` column G: 9 cells, 1 formula OK
- power `fix` 2022-2040: 342 cells, 1 formula OK
  - `=R[-20]C[0]*R[0]C4+R[0]C5+R[-10]C[0]`
- power `fix` column D: 9 cells, 1 formula OK
- power `fix` column E: 9 cells, 1 formula OK
- power `fix` column F: 9 cells, 1 formula OK
- power `fix` column G: 9 cells, 1 formula OK
- power `lfx` 2022-2040: 342 cells, 1 formula OK
  - `=R[-41]C[0]*R[0]C4+R[0]C5+R[0]C6*R749C[0]`
- power `lfx` column D: 9 cells, 1 formula OK
- power `lfx` column E: 9 cells, 1 formula OK
- power `lfx` column F: 9 cells, 1 formula OK
- power `lfx` column G: 9 cells, 1 formula OK
- power `xck` 2022-2040: 342 cells, 1 formula OK
  - `=R[0]C4*(R[0]C5*(R321C[0]-R322C[0])+R[0]C6*(R324C[0]-R325C[0]))`
- power `xck` column D: 9 cells, 1 formula OK
- power `xck` column E: 9 cells, 1 formula OK
- power `xck` column F: 9 cells, 1 formula OK
- power `xck` column G: 9 cells, 1 formula OK
- power `vbc` 2022-2040: 342 cells, 1 formula OK
  - `=R[0]C4+INDEX(R381C[0]:R392C[0],R[0]C5)*R[0]C6+R[-10]C[0]`
- power `vbc` column D: 9 cells, 1 formula OK
- power `vbc` column E: 9 cells, 1 formula OK
- power `vbc` column F: 9 cells, 1 formula OK
- power `vbc` column G: 9 cells, 1 formula OK
- power `ccp` 2022-2040: 342 cells, 1 formula OK
  - `=R[0]C4*(R263C[0]*INDEX(R23C[0]:R30C[0],R[0]C5)+R280C[0])+INDEX(R80C[0]:R93C[0],R[0]C6)*R[0]C7`
- power `ccp` column D: 9 cells, 1 formula OK
- power `ccp` column E: 9 cells, 1 formula OK
- power `ccp` column F: 9 cells, 1 formula OK
- power `ccp` column G: 9 cells, 1 formula OK
- power `shk` 2022-2040: 342 cells, 1 formula OK
  - `=R297C[0]*R[0]C4`
- power `shk` column D: 9 cells, 1 formula OK
- power `shk` column E: 9 cells, 1 formula OK
- power `shk` column F: 9 cells, 1 formula OK
- power `shk` column G: 9 cells, 1 formula OK
- power `reb` base year: 2 cells, 1 formula OK
  - `=0`
- power `reb` 2023-2040: 36 cells, 1 formula OK
  - `=-SUMPRODUCT(R549C[0]:R557C[0],R549C5:R557C5)`
- power `reb` column D: 1 cells, 1 formula OK
- power `reb` column E: 1 cells, 1 formula OK
- power `reb` column F: 1 cells, 1 formula OK
- power `reb` column G: 1 cells, 1 formula OK
- power `tvc` 2022-2040: 342 cells, 1 formula OK
  - `=R[-32]C[0]+R[-22]C[0]+R[-12]C[0]+R559C[0]`
- power `tvc` column D: 9 cells, 1 formula OK
- power `tvc` column E: 9 cells, 1 formula OK
- power `tvc` column F: 9 cells, 1 formula OK
- power `tvc` column G: 9 cells, 1 formula OK
- power `dsum` 2022-2040 (data, sum or look-ahead: one formula): 342 cells, 1 formula OK
  - `=(1+R[0]C4)^(Settings!R7C3-R4C[0])+R[0]C[1]`
- power `dsum` column D: 9 cells, 1 formula OK
- power `dsum` column E: 9 cells, 1 formula OK
- power `dsum` column F: 9 cells, 1 formula OK
- power `dsum` column G: 9 cells, 1 formula OK
- power `lvc` 2022-2040 (data, sum or look-ahead: one formula): 342 cells, 1 formula OK
  - `=((1+R[0]C4)^(Settings!R7C3-R4C[0])*R[-20]C[0]+(R[-10]C[0]-(1+R[0]C4)^(Settings!R7C3-R4C[0]))*R[0]C[1])/R[-10]C[0]`
- power `lvc` column D: 9 cells, 1 formula OK
- power `lvc` column E: 9 cells, 1 formula OK
- power `lvc` column F: 9 cells, 1 formula OK
- power `lvc` column G: 9 cells, 1 formula OK
- power `gnc` 2022-2040: 342 cells, 1 formula OK
  - `=R[-95]C[0]+R[-63]C[0]`
- power `gnc` column D: 9 cells, 1 formula OK
- power `gnc` column E: 9 cells, 1 formula OK
- power `gnc` column F: 9 cells, 1 formula OK
- power `gnc` column G: 9 cells, 1 formula OK
- power `gncav` 2022-2040: 38 cells, 1 formula OK
  - `=SUMPRODUCT(R716C[0]:R724C[0],R592C[0]:R600C[0])*R[0]C4`
- power `gncav` column D: 1 cells, 1 formula OK
- power `gncav` column E: 1 cells, 1 formula OK
- power `gncav` column F: 1 cells, 1 formula OK
- power `gncav` column G: 1 cells, 1 formula OK
- power `ccpav` 2022-2040: 38 cells, 1 formula OK
  - `=SUMPRODUCT(R716C[0]:R724C[0],R539C[0]:R547C[0])*R[0]C4`
- power `ccpav` column D: 1 cells, 1 formula OK
- power `ccpav` column E: 1 cells, 1 formula OK
- power `ccpav` column F: 1 cells, 1 formula OK
- power `ccpav` column G: 1 cells, 1 formula OK
- power `tic` 2022-2040: 342 cells, 1 formula OK
  - `=MAX(R[-97]C[0]+R[-24]C[0],R[0]C4)`
- power `tic` column D: 9 cells, 1 formula OK
- power `tic` column E: 9 cells, 1 formula OK
- power `tic` column F: 9 cells, 1 formula OK
- power `tic` column G: 9 cells, 1 formula OK
- power `cons` 2022-2040: 76 cells, 1 formula OK
  - `=(R[0]C4*R675C[0]+R[0]C5*R658C[0])*R[0]C6`
- power `cons` column D: 2 cells, 1 formula OK
- power `cons` column E: 2 cells, 1 formula OK
- power `cons` column F: 2 cells, 1 formula OK
- power `cons` column G: 2 cells, 1 formula OK
- power `sc` base year: 4 cells, 1 formula OK
  - `=R602C[0]+R[0]C4`
- power `sc` 2023-2040: 72 cells, 1 formula OK
  - `=R602C[-1]+R[0]C4`
- power `sc` column D: 2 cells, 1 formula OK
- power `sc` column E: 2 cells, 1 formula OK
- power `sc` column F: 2 cells, 1 formula OK
- power `sc` column G: 2 cells, 1 formula OK
- power `rppre` history 2022-2024: 12 cells, 1 formula OK
  - `=INDEX(Inputs_power!R19C10:R20C13,R[0]C5,MATCH(R4C[0],Inputs_power!R18C10:R18C13,0))`
- power `rppre` projection 2025-2040: 64 cells, 1 formula OK
  - `=R[0]C[-1]+R[0]C4*(R[-3]C[0]-R[-3]C[-1])*(1+R[0]C6)`
- power `rppre` column D: 2 cells, 1 formula OK
- power `rppre` column E: 2 cells, 1 formula OK
- power `rppre` column F: 2 cells, 1 formula OK
- power `rppre` column G: 2 cells, 1 formula OK
- power `rp` 2022-2040: 76 cells, 1 formula OK
  - `=R[-3]C[0]+R603C[0]*R[0]C4*(1+R[0]C5)`
- power `rp` column D: 2 cells, 1 formula OK
- power `rp` column E: 2 cells, 1 formula OK
- power `rp` column F: 2 cells, 1 formula OK
- power `rp` column G: 2 cells, 1 formula OK
- power `sgap` 2022-2040: 76 cells, 1 formula OK
  - `=R[-9]C[0]-R[-6]C[0]/(1+R[0]C4)`
- power `sgap` column D: 2 cells, 1 formula OK
- power `sgap` column E: 2 cells, 1 formula OK
- power `sgap` column F: 2 cells, 1 formula OK
- power `sgap` column G: 2 cells, 1 formula OK
- power `rvat` 2022-2040: 76 cells, 1 formula OK
  - `=R[-9]C[0]*R[0]C4*R[-15]C[0]`
- power `rvat` column D: 2 cells, 1 formula OK
- power `rvat` column E: 2 cells, 1 formula OK
- power `rvat` column F: 2 cells, 1 formula OK
- power `rvat` column G: 2 cells, 1 formula OK
- power `rcarb` 2022-2040: 76 cells, 1 formula OK
  - `=R603C[0]*R[0]C4*R[-18]C[0]`
- power `rcarb` column D: 2 cells, 1 formula OK
- power `rcarb` column E: 2 cells, 1 formula OK
- power `rcarb` column F: 2 cells, 1 formula OK
- power `rcarb` column G: 2 cells, 1 formula OK
- power `rnvat` 2022-2040: 76 cells, 1 formula OK
  - `=R603C[0]*R[0]C4*R[0]C5*R[-21]C[0]`
- power `rnvat` column D: 2 cells, 1 formula OK
- power `rnvat` column E: 2 cells, 1 formula OK
- power `rnvat` column F: 2 cells, 1 formula OK
- power `rnvat` column G: 2 cells, 1 formula OK
- power `rgap` 2022-2040: 76 cells, 1 formula OK
  - `=R[-12]C[0]*R[0]C4*R[-24]C[0]`
- power `rgap` column D: 2 cells, 1 formula OK
- power `rgap` column E: 2 cells, 1 formula OK
- power `rgap` column F: 2 cells, 1 formula OK
- power `rgap` column G: 2 cells, 1 formula OK
- power `pdem` base year: 8 cells, 1 formula OK
  - `=R[0]C5`
- power `pdem` 2023-2040: 144 cells, 1 formula OK
  - `=R[0]C4*R625C[-1]+(1-R[0]C4)*R626C[-1]`
- power `pdem` column D: 4 cells, 1 formula OK
- power `pdem` column E: 4 cells, 1 formula OK
- power `pdem` column F: 4 cells, 1 formula OK
- power `pdem` column G: 4 cells, 1 formula OK
- power `mult` base year: 8 cells, 1 formula OK
  - `=1`
- power `mult` 2023-2040: 144 cells, 1 formula OK
  - `=R[0]C4*(1+R7C[0])^(R[0]C5*R463C[0])*(R[-5]C[0]/R[-5]C[-1])^R[0]C6*(1+R464C[0]*Settings!R19C3)`
- power `mult` column D: 4 cells, 1 formula OK
- power `mult` column E: 4 cells, 1 formula OK
- power `mult` column F: 4 cells, 1 formula OK
- power `mult` column G: 4 cells, 1 formula OK
- power `dem` base year: 30 cells, 1 formula OK
  - `=R[0]C5`
- power `dem` 2023-2040: 540 cells, 1 formula OK
  - `=R[0]C[-1]*INDEX(R649C[0]:R652C[0],R[0]C4)`
- power `dem` column D: 15 cells, 1 formula OK
- power `dem` column E: 15 cells, 1 formula OK
- power `dem` column F: 15 cells, 1 formula OK
- power `dem` column G: 15 cells, 1 formula OK
- power `demg` 2022-2040 (data, sum or look-ahead: one formula): 152 cells, 1 formula OK
  - `=SUMIF(R654C4:R668C4,R[0]C4,R654C[0]:R668C[0])`
- power `demg` column D: 4 cells, 1 formula OK
- power `demg` column E: 4 cells, 1 formula OK
- power `demg` column F: 4 cells, 1 formula OK
- power `demg` column G: 4 cells, 1 formula OK
- power `demtot` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=SUM(R654C[0]:R668C[0])`
- power `demtot` column D: 1 cells, 1 formula OK
- power `demtot` column E: 1 cells, 1 formula OK
- power `demtot` column F: 1 cells, 1 formula OK
- power `demtot` column G: 1 cells, 1 formula OK
- power `req` base year: 2 cells, 1 formula OK
  - `=SUM(R703C[0]:R711C[0])`
- power `req` 2023-2040: 36 cells, 1 formula OK
  - `=R675C[0]*R[0]C4`
- power `req` column D: 1 cells, 1 formula OK
- power `req` column E: 1 cells, 1 formula OK
- power `req` column F: 1 cells, 1 formula OK
- power `req` column G: 1 cells, 1 formula OK
- power `reqcg` 2022-2040: 38 cells, 1 formula OK
  - `=MAX(R676C[0]-SUMPRODUCT(R764C[0]:R772C[0],1-R703C4:R711C4),0)`
- power `reqcg` column D: 1 cells, 1 formula OK
- power `reqcg` column E: 1 cells, 1 formula OK
- power `reqcg` column F: 1 cells, 1 formula OK
- power `reqcg` column G: 1 cells, 1 formula OK
- power `vcmin` 2022-2040: 38 cells, 1 formula OK
  - `=MAX(_xlfn.MINIFS(R561C[0]:R569C[0],R703C4:R711C4,1),R[0]C4)`
- power `vcmin` column D: 1 cells, 1 formula OK
- power `vcmin` column E: 1 cells, 1 formula OK
- power `vcmin` column F: 1 cells, 1 formula OK
- power `vcmin` column G: 1 cells, 1 formula OK
- power `rem` 2022-2040: 38 cells, 1 formula OK
  - `=MAX(R679C[0]-SUM(R683C[0]:R691C[0]),0)`
- power `rem` column D: 1 cells, 1 formula OK
- power `rem` column E: 1 cells, 1 formula OK
- power `rem` column F: 1 cells, 1 formula OK
- power `rem` column G: 1 cells, 1 formula OK
- power `gmin` 2022-2040: 342 cells, 1 formula OK
  - `=R[0]C4*MIN(MAX(R679C[0]-(SUMPRODUCT(R752C[0]:R760C[0],R703C4:R711C4)*R[0]C5-R[69]C[0]*R[0]C5),0),R[69]C[0]*R[0]C5)`
- power `gmin` column D: 9 cells, 1 formula OK
- power `gmin` column E: 9 cells, 1 formula OK
- power `gmin` column F: 9 cells, 1 formula OK
- power `gmin` column G: 9 cells, 1 formula OK
- power `lgd` 2022-2040: 342 cells, 1 formula OK
  - `=R[0]C4*EXP(-R[0]C5*MAX(R[-132]C[0],R[0]C6)/R680C[0])`
- power `lgd` column D: 9 cells, 1 formula OK
- power `lgd` column E: 9 cells, 1 formula OK
- power `lgd` column F: 9 cells, 1 formula OK
- power `lgd` column G: 9 cells, 1 formula OK
- power `gen` base year: 18 cells, 1 formula OK
  - `=R[0]C6`
- power `gen` 2023-2040: 324 cells, 1 formula OK
  - `=IF(R[0]C4=1,MIN(R[49]C[0]*R[0]C5,R[-20]C[0]+R681C[0]*R[-10]C[0]/SUM(R693C[0]:R701C[0])),R[61]C[0])`
- power `gen` column D: 9 cells, 1 formula OK
- power `gen` column E: 9 cells, 1 formula OK
- power `gen` column F: 9 cells, 1 formula OK
- power `gen` column G: 9 cells, 1 formula OK
- power `gtot` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=SUM(R703C[0]:R711C[0])`
- power `gtot` column D: 1 cells, 1 formula OK
- power `gtot` column E: 1 cells, 1 formula OK
- power `gtot` column F: 1 cells, 1 formula OK
- power `gtot` column G: 1 cells, 1 formula OK
- power `dgen` base year: 2 cells, 1 formula OK
  - `=0`
- power `dgen` 2023-2040: 36 cells, 1 formula OK
  - `=MAX(0,R713C[0]-R713C[-1])`
- power `dgen` column D: 1 cells, 1 formula OK
- power `dgen` column E: 1 cells, 1 formula OK
- power `dgen` column F: 1 cells, 1 formula OK
- power `dgen` column G: 1 cells, 1 formula OK
- power `shr` 2022-2040: 342 cells, 1 formula OK
  - `=IF(R713C[0]=0,0,R[-13]C[0]/R713C[0])`
- power `shr` column D: 9 cells, 1 formula OK
- power `shr` column E: 9 cells, 1 formula OK
- power `shr` column F: 9 cells, 1 formula OK
- power `shr` column G: 9 cells, 1 formula OK
- power `pe` 2022-2040: 342 cells, 1 formula OK
  - `=R[-23]C[0]*R[0]C4`
- power `pe` column D: 9 cells, 1 formula OK
- power `pe` column E: 9 cells, 1 formula OK
- power `pe` column F: 9 cells, 1 formula OK
- power `pe` column G: 9 cells, 1 formula OK
- power `co2` 2022-2040: 342 cells, 1 formula OK
  - `=R[-10]C[0]*R[0]C4`
- power `co2` column D: 9 cells, 1 formula OK
- power `co2` column E: 9 cells, 1 formula OK
- power `co2` column F: 9 cells, 1 formula OK
- power `co2` column G: 9 cells, 1 formula OK
- power `vre` 2022-2040: 38 cells, 1 formula OK
  - `=SUMPRODUCT(R716C[0]:R724C[0],R716C4:R724C4)*R[0]C4`
- power `vre` column D: 1 cells, 1 formula OK
- power `vre` column E: 1 cells, 1 formula OK
- power `vre` column F: 1 cells, 1 formula OK
- power `vre` column G: 1 cells, 1 formula OK
- power `cph` 2022-2040: 38 cells, 1 formula OK
  - `=(R459C[0]+R460C[0]/R[0]C4)/(R[0]C5*8760)+R461C[0]/8760`
- power `cph` column D: 1 cells, 1 formula OK
- power `cph` column E: 1 cells, 1 formula OK
- power `cph` column F: 1 cells, 1 formula OK
- power `cph` column G: 1 cells, 1 formula OK
- power `msc` 2022-2040: 38 cells, 1 formula OK
  - `=R[0]C4*R[0]C5*R747C[0]*R748C[0]+2*MAX(0,R747C[0]-R[0]C6)/(1-R[0]C6)^2*R462C[0]*R[0]C7`
- power `msc` column D: 1 cells, 1 formula OK
- power `msc` column E: 1 cells, 1 formula OK
- power `msc` column F: 1 cells, 1 formula OK
- power `msc` column G: 1 cells, 1 formula OK
- power `cap` base year: 18 cells, 1 formula OK
  - `=R[0]C4`
- power `cap` 2023-2040: 324 cells, 1 formula OK
  - `=R[0]C[-1]-R[32]C[-1]+R[143]C[-1]`
- power `cap` column D: 9 cells, 1 formula OK
- power `cap` column E: 9 cells, 1 formula OK
- power `cap` column F: 9 cells, 1 formula OK
- power `cap` column G: 9 cells, 1 formula OK
- power `captot` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=SUM(R752C[0]:R760C[0])`
- power `captot` column D: 1 cells, 1 formula OK
- power `captot` column E: 1 cells, 1 formula OK
- power `captot` column F: 1 cells, 1 formula OK
- power `captot` column G: 1 cells, 1 formula OK
- power `effcap` 2022-2040: 342 cells, 1 formula OK
  - `=R[-12]C[0]*R[0]C4`
- power `effcap` column D: 9 cells, 1 formula OK
- power `effcap` column E: 9 cells, 1 formula OK
- power `effcap` column F: 9 cells, 1 formula OK
- power `effcap` column G: 9 cells, 1 formula OK
- power `pret` 2022-2040 (data, sum or look-ahead: one formula): 342 cells, 1 formula OK
  - `=INDEX(PPathVal,MATCH(R[0]C4,PPathKey,0),MATCH(R4C[0],PPathYear,0))`
- power `pret` column D: 9 cells, 1 formula OK
- power `pret` column E: 9 cells, 1 formula OK
- power `pret` column F: 9 cells, 1 formula OK
- power `pret` column G: 9 cells, 1 formula OK
- power `ret` 2022-2040: 342 cells, 1 formula OK
  - `=MIN(IF(R[0]C6=1,R[-10]C[0],R[-32]C[0]*R[0]C4)+R[0]C5*R807C[0],R[-32]C[0])`
- power `ret` column D: 9 cells, 1 formula OK
- power `ret` column E: 9 cells, 1 formula OK
- power `ret` column F: 9 cells, 1 formula OK
- power `ret` column G: 9 cells, 1 formula OK
- power `reteff` 2022-2040: 342 cells, 1 formula OK
  - `=R[-10]C[0]*R[0]C4`
- power `reteff` column D: 9 cells, 1 formula OK
- power `reteff` column E: 9 cells, 1 formula OK
- power `reteff` column F: 9 cells, 1 formula OK
- power `reteff` column G: 9 cells, 1 formula OK
- power `egrow` base year: 2 cells, 1 formula OK
  - `=MAX(0,R829C[0]+R830C[0])`
- power `egrow` 2023-2040: 36 cells, 1 formula OK
  - `=MAX(0,R829C[0]+R830C[0]-R714C[-1]-(SUM(R794C[-1]:R802C[-1])-R794C[-1]))`
- power `egrow` column D: 1 cells, 1 formula OK
- power `egrow` column E: 1 cells, 1 formula OK
- power `egrow` column F: 1 cells, 1 formula OK
- power `egrow` column G: 1 cells, 1 formula OK
- power `ecmin` 2022-2040: 38 cells, 1 formula OK
  - `=MIN(R561C[0],R609C[0],R610C[0])`
- power `ecmin` column D: 1 cells, 1 formula OK
- power `ecmin` column E: 1 cells, 1 formula OK
- power `ecmin` column F: 1 cells, 1 formula OK
- power `ecmin` column G: 1 cells, 1 formula OK
- power `ecshr` 2022-2040: 38 cells, 1 formula OK
  - `=1/(1+EXP(-R[0]C4*(R609C[0]-R561C[0])/R805C[0])+EXP(-R[0]C4*(R610C[0]-R561C[0])/R805C[0]))`
- power `ecshr` column D: 1 cells, 1 formula OK
- power `ecshr` column E: 1 cells, 1 formula OK
- power `ecshr` column F: 1 cells, 1 formula OK
- power `ecshr` column G: 1 cells, 1 formula OK
- power `early` 2022-2040: 38 cells, 1 formula OK
  - `=MAX(0,MIN(R764C[0]*R[0]C4,R804C[0])*(1-R806C[0]))/R[0]C6*R[0]C5`
- power `early` column D: 1 cells, 1 formula OK
- power `early` column E: 1 cells, 1 formula OK
- power `early` column F: 1 cells, 1 formula OK
- power `early` column G: 1 cells, 1 formula OK
- power `kafter` 2022-2040: 38 cells, 1 formula OK
  - `=SUM(R764C[0]:R772C[0])-SUM(R794C[0]:R802C[0])`
- power `kafter` column D: 1 cells, 1 formula OK
- power `kafter` column E: 1 cells, 1 formula OK
- power `kafter` column F: 1 cells, 1 formula OK
- power `kafter` column G: 1 cells, 1 formula OK
- power `need` 2022-2040: 38 cells, 1 formula OK
  - `=MAX(R676C[0]-R808C[0],0)`
- power `need` column D: 1 cells, 1 formula OK
- power `need` column E: 1 cells, 1 formula OK
- power `need` column F: 1 cells, 1 formula OK
- power `need` column G: 1 cells, 1 formula OK
- power `cmin` 2022-2040: 38 cells, 1 formula OK
  - `=MAX(_xlfn.MINIFS(R605C[0]:R613C[0],R815C[0]:R823C[0],1),R[0]C4)`
- power `cmin` column D: 1 cells, 1 formula OK
- power `cmin` column E: 1 cells, 1 formula OK
- power `cmin` column F: 1 cells, 1 formula OK
- power `cmin` column G: 1 cells, 1 formula OK
- power `deficit` 2022-2040: 38 cells, 1 formula OK
  - `=MAX(R809C[0]-SUM(R855C[0]:R863C[0]),0)`
- power `deficit` column D: 1 cells, 1 formula OK
- power `deficit` column E: 1 cells, 1 formula OK
- power `deficit` column F: 1 cells, 1 formula OK
- power `deficit` column G: 1 cells, 1 formula OK
- power `permit` 2022-2040: 342 cells, 1 formula OK
  - `=IF(AND(R[0]C4="Yes",R4C[0]>=R[0]C5),1,0)`
- power `permit` column D: 9 cells, 1 formula OK
- power `permit` column E: 9 cells, 1 formula OK
- power `permit` column F: 9 cells, 1 formula OK
- power `permit` column G: 9 cells, 1 formula OK
- power `maxinv` 2022-2040: 342 cells, 1 formula OK
  - `=R[0]C4*SUM(R764C[0]:R772C[0])*R[-10]C[0]`
- power `maxinv` column D: 9 cells, 1 formula OK
- power `maxinv` column E: 9 cells, 1 formula OK
- power `maxinv` column F: 9 cells, 1 formula OK
- power `maxinv` column G: 9 cells, 1 formula OK
- power `lgi` 2022-2040: 342 cells, 1 formula OK
  - `=R[-20]C[0]*EXP(-R[0]C4*R[-230]C[0]/R812C[0])`
- power `lgi` column D: 9 cells, 1 formula OK
- power `lgi` column E: 9 cells, 1 formula OK
- power `lgi` column F: 9 cells, 1 formula OK
- power `lgi` column G: 9 cells, 1 formula OK
- power `pni` 2022-2040: 342 cells, 1 formula OK
  - `=IF(SUM(R835C[0]:R843C[0])=0,0,R[-10]C[0]/SUM(R835C[0]:R843C[0]))`
- power `pni` column D: 9 cells, 1 formula OK
- power `pni` column E: 9 cells, 1 formula OK
- power `pni` column F: 9 cells, 1 formula OK
- power `pni` column G: 9 cells, 1 formula OK
- power `first` 2022-2040: 342 cells, 1 formula OK
  - `=MIN(R[-30]C[0],R809C[0]*R[-10]C[0])`
- power `first` column D: 9 cells, 1 formula OK
- power `first` column E: 9 cells, 1 formula OK
- power `first` column F: 9 cells, 1 formula OK
- power `first` column G: 9 cells, 1 formula OK
- power `space` 2022-2040: 342 cells, 1 formula OK
  - `=(R[-40]C[0]-R[-10]C[0])*R[-50]C[0]`
- power `space` column D: 9 cells, 1 formula OK
- power `space` column E: 9 cells, 1 formula OK
- power `space` column F: 9 cells, 1 formula OK
- power `space` column G: 9 cells, 1 formula OK
- power `second` 2022-2040: 342 cells, 1 formula OK
  - `=R[-60]C[0]*MAX(0,MIN(R[-10]C[0],R813C[0]-SUMPRODUCT((R605C[0]:R613C[0]+R875C4:R883C4<R[-270]C[0]+R[0]C4)*R865C[0]:R873C[0])))`
- power `second` column D: 9 cells, 1 formula OK
- power `second` column E: 9 cells, 1 formula OK
- power `second` column F: 9 cells, 1 formula OK
- power `second` column G: 9 cells, 1 formula OK
- power `nei` 2022-2040: 342 cells, 1 formula OK
  - `=R[-30]C[0]+R[-10]C[0]`
- power `nei` column D: 9 cells, 1 formula OK
- power `nei` column E: 9 cells, 1 formula OK
- power `nei` column F: 9 cells, 1 formula OK
- power `nei` column G: 9 cells, 1 formula OK
- power `nim` 2022-2040: 342 cells, 1 formula OK
  - `=R[-10]C[0]*R[0]C4`
- power `nim` column D: 9 cells, 1 formula OK
- power `nim` column E: 9 cells, 1 formula OK
- power `nim` column F: 9 cells, 1 formula OK
- power `nim` column G: 9 cells, 1 formula OK
- power `nimtot` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=SUM(R895C[0]:R903C[0])`
- power `nimtot` column D: 1 cells, 1 formula OK
- power `nimtot` column E: 1 cells, 1 formula OK
- power `nimtot` column F: 1 cells, 1 formula OK
- power `nimtot` column G: 1 cells, 1 formula OK
- power `usd` 2022-2040: 342 cells, 1 formula OK
  - `=R[-12]C[0]*R[-440]C[0]/1000`
- power `usd` column D: 9 cells, 1 formula OK
- power `usd` column E: 9 cells, 1 formula OK
- power `usd` column F: 9 cells, 1 formula OK
- power `usd` column G: 9 cells, 1 formula OK
- power `usdtot` 2022-2040 (data, sum or look-ahead: one formula): 38 cells, 1 formula OK
  - `=SUM(R907C[0]:R915C[0])`
- power `usdtot` column D: 1 cells, 1 formula OK
- power `usdtot` column E: 1 cells, 1 formula OK
- power `usdtot` column F: 1 cells, 1 formula OK
- power `usdtot` column G: 1 cells, 1 formula OK
- power `phi` 2022-2040: 342 cells, 1 formula OK
  - `=IF(R[-167]C[0]+R[-24]C[0]-R[-135]C[0]=0,0,R[-24]C[0]/(R[-167]C[0]+R[-24]C[0]-R[-135]C[0]))`
- power `phi` column D: 9 cells, 1 formula OK
- power `phi` column E: 9 cells, 1 formula OK
- power `phi` column F: 9 cells, 1 formula OK
- power `phi` column G: 9 cells, 1 formula OK
- power `prtx` 2022-2040: 342 cells, 1 formula OK
  - `=R[0]C5*R[-204]C[0]*Settings!R13C3*(INDEX(R407C[0]:R418C[0],R[0]C4)+INDEX(R394C[0]:R405C[0],R[0]C4))`
- power `prtx` column D: 9 cells, 1 formula OK
- power `prtx` column E: 9 cells, 1 formula OK
- power `prtx` column F: 9 cells, 1 formula OK
- power `prtx` column G: 9 cells, 1 formula OK
- power `prsub` 2022-2040: 342 cells, 1 formula OK
  - `=R[0]C5*R[-214]C[0]*Settings!R13C3*(INDEX(R420C[0]:R431C[0],R[0]C4))`
- power `prsub` column D: 9 cells, 1 formula OK
- power `prsub` column E: 9 cells, 1 formula OK
- power `prsub` column F: 9 cells, 1 formula OK
- power `prsub` column G: 9 cells, 1 formula OK
- power `prxc` 2022-2040: 342 cells, 1 formula OK
  - `=R[-431]C[0]*8.76*R[-247]C[0]`
- power `prxc` column D: 9 cells, 1 formula OK
- power `prxc` column E: 9 cells, 1 formula OK
- power `prxc` column F: 9 cells, 1 formula OK
- power `prxc` column G: 9 cells, 1 formula OK
- power `prnew` 2022-2040: 342 cells, 1 formula OK
  - `=R[-257]C[0]*8.76*R[-421]C[0]-R[-224]C[0]*R280C[0]*(1-R201C[0])`
- power `prnew` column D: 9 cells, 1 formula OK
- power `prnew` column E: 9 cells, 1 formula OK
- power `prnew` column F: 9 cells, 1 formula OK
- power `prnew` column G: 9 cells, 1 formula OK
- power `prxl` 2022-2040: 342 cells, 1 formula OK
  - `=R[-234]C[0]*(R[0]C4*R321C[0]+R[0]C5*R324C[0])`
- power `prxl` column D: 9 cells, 1 formula OK
- power `prxl` column E: 9 cells, 1 formula OK
- power `prxl` column F: 9 cells, 1 formula OK
- power `prxl` column G: 9 cells, 1 formula OK
- sheet Lambdas: an optional LAMBDA formula for each of the 2435 calculated rows (2435 listed, each calls a named LAMBDA) OK
- column groups, one level: B:G, I:K and 2030-2039 of each scenario (20 columns), rolled up; other columns visible: OK
- blank spacer column after each scenario group (group width 21): OK
- light beige (EEECE1): right column not marked (plain formulas) OK; 2023-2024 history block of sp, txo, rppre (base year darker beige, projection white) OK

## 3. LAMBDA encoding
- PATH: 6 parameters, 20 _xlpm. tokens, bare names [] OK
- PATHYEARS: 4 parameters, 9 _xlpm. tokens, bare names [] OK
- NOMFACTOR: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- ETSON: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- ETSPRICE: 5 parameters, 11 _xlpm. tokens, bare names [] OK
- OBRSHARE: 1 parameters, 2 _xlpm. tokens, bare names [] OK
- VOLPICK: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- VOLADJ: 3 parameters, 7 _xlpm. tokens, bare names [] OK
- COVERED: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- CAPCHANGE: 5 parameters, 13 _xlpm. tokens, bare names [] OK
- ETSCAP: 7 parameters, 14 _xlpm. tokens, bare names [] OK
- ETSSEMI: 6 parameters, 13 _xlpm. tokens, bare names [] OK
- ETSESTIMATE: 5 parameters, 15 _xlpm. tokens, bare names [] OK
- TAXEQUIV: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- REVFACTOR: 1 parameters, 3 _xlpm. tokens, bare names [] OK
- ETSCOVER: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- SHADOWSECTOR: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- SHADOWSHARE: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- SUPPLYCOST: 6 parameters, 14 _xlpm. tokens, bare names [] OK
- OTHERTAX: 6 parameters, 25 _xlpm. tokens, bare names [] OK
- RETAILPRICE: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- VATPAY: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- TAXPART: 1 parameters, 2 _xlpm. tokens, bare names [] OK
- SUBSIDYPART: 1 parameters, 2 _xlpm. tokens, bare names [] OK
- PERUNIT: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- SECTORTAX: 5 parameters, 10 _xlpm. tokens, bare names [] OK
- ECPPRICE: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- XCARBON: 5 parameters, 10 _xlpm. tokens, bare names [] OK
- SECTORETS: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- CARBONTAX: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- ETSCOST: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- NEWEXCISE: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- NEWPOLICY: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- POSTTAX: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- SHADOWTOTAL: 5 parameters, 10 _xlpm. tokens, bare names [] OK
- SHADOWGJ: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- KWHPOLICY: 6 parameters, 12 _xlpm. tokens, bare names [] OK
- FLOATTAX: 4 parameters, 9 _xlpm. tokens, bare names [] OK
- REVENUE: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- EMISSIONS: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- NEWREVRATE: 6 parameters, 12 _xlpm. tokens, bare names [] OK
- FUELUSE: 10 parameters, 22 _xlpm. tokens, bare names [] OK
- CAPEX: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- VINTAGE: 3 parameters, 7 _xlpm. tokens, bare names [] OK
- AMORTISED: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- LEVELISED: 5 parameters, 10 _xlpm. tokens, bare names [] OK
- VARCOST: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- PNEWREV: 5 parameters, 10 _xlpm. tokens, bare names [] OK
- GENCOST: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- WEIGHTED: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- STORAGEHOUR: 5 parameters, 10 _xlpm. tokens, bare names [] OK
- STORAGECOST: 7 parameters, 16 _xlpm. tokens, bare names [] OK
- POWERSUPPLY: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- POWERPRICE: 7 parameters, 15 _xlpm. tokens, bare names [] OK
- ENDUSERPOWER: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- SUBSIDYGAP: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- POWERREV: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- REBATE: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- VARTOTAL: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- INVCOST: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- CONSUMPTION: 5 parameters, 10 _xlpm. tokens, bare names [] OK
- PRICEMIX: 3 parameters, 7 _xlpm. tokens, bare names [] OK
- DEMANDMULT: 9 parameters, 18 _xlpm. tokens, bare names [] OK
- GROW: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- SCALE: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- RESIDUAL: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- FLOORMIN: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- MINDISPATCH: 4 parameters, 9 _xlpm. tokens, bare names [] OK
- LOGITW: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- DISPATCH: 7 parameters, 14 _xlpm. tokens, bare names [] OK
- INCREASE: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- SHARE: 2 parameters, 5 _xlpm. tokens, bare names [] OK
- CARRY: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- RETIRE: 6 parameters, 13 _xlpm. tokens, bare names [] OK
- EARLYROOM: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- LOWEST3: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- COALSHARE: 5 parameters, 13 _xlpm. tokens, bare names [] OK
- EARLYRET: 6 parameters, 12 _xlpm. tokens, bare names [] OK
- DIFF: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- PERMITTED: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- SCALEUP: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- FIRSTROUND: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- SPACE: 3 parameters, 6 _xlpm. tokens, bare names [] OK
- CHEAPFILL: 4 parameters, 8 _xlpm. tokens, bare names [] OK
- ADDED: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- INVSPEND: 2 parameters, 4 _xlpm. tokens, bare names [] OK
- VINTSHARE: 3 parameters, 10 _xlpm. tokens, bare names [] OK

## 4. Values vs independent Python recomputation
- A. as shipped (plain formulas, no LAMBDA), scenario 1: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 6.47e-14 (at P714) OK
- B. optional LAMBDA library (sheet Lambdas) copied over every calculated row, LAMBDAs expanded, scenario 1: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 6.47e-14 (at P714) OK
- A. as shipped (plain formulas, no LAMBDA), scenario 2: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 6.22e-14 (at AK714) OK
- B. optional LAMBDA library (sheet Lambdas) copied over every calculated row, LAMBDAs expanded, scenario 2: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 6.22e-14 (at AK714) OK

## 4b. Power method = legacy (Python reference with the legacy inputs vs cached legacy rows)
- baseline: 208 legacy rows (demand by subsector, required generation, capacity, generation, shares, fuel use, retirements, investment shares and both rounds, nameplate and USD investment, B1 phi, E1 investment cost, C7, B2, B6, E3, J, G1 prices, early retirement), 2022-2040: max relative diff 9.6e-14 (at second|nga 2026) OK
- carbon-tax run: 208 legacy rows (demand by subsector, required generation, capacity, generation, shares, fuel use, retirements, investment shares and both rounds, nameplate and USD investment, B1 phi, E1 investment cost, C7, B2, B6, E3, J, G1 prices, early retirement), 2022-2040: max relative diff 4.4e-13 (at second|nga 2040) OK

Effect of each bug fix on the legacy Egypt inputs (2040; % change against legacy):

| Fix | legacy rows | generation (base) | power CO2 (base) | generation (carbon tax) | VRE share (carbon tax) | investment USD 2022-2040 (base) |
|---|---|---|---|---|---|---|
| services | 3731 | +7.69% | +6.62% | +3.98% | +1.21% | +11.28% |
| aee | 3691, 3697, 3703, 3709 | +0.90% | +0.77% | +0.95% | +0.53% | +1.33% |
| income | 3692, 3698, 3704, 3710 | +3.22% | +2.74% | +3.66% | +1.81% | +4.60% |
| lt_alloc | 3911, 3912 | +0.00% | +0.00% | +0.00% | +0.00% | +0.00% |
| invcost | 4276-4284 | +0.00% | +0.00% | +0.00% | +0.00% | -11.26% |
| cmin | 4046, 4073-4081 | +0.00% | +0.00% | -0.07% | -1.40% | +0.00% |
| retlag | 3973 | +0.00% | +0.00% | +0.00% | +0.00% | +0.00% |
| coalret | 4009-4011 | +0.00% | +0.00% | +0.00% | +0.00% | +0.00% |
| rebate | 3407-3409 | +0.00% | +0.00% | +0.00% | +0.00% | +0.00% |
| kcheck | 4024 | +0.00% | +0.00% | +0.00% | +0.00% | +0.00% |
| all |  | +12.59% | +10.83% | +9.15% | +1.73% | +4.74% |

## 5. Scenario tests (copied MTInputs column + Mitigation group)
- scenario 3 (copy of 2) = scenario 2: 52877 numeric cells, max abs diff 0 OK
- scenario 4 (copy of 2, carbon price 0 in MTInputs) = scenario 1: 52839 numeric cells, max abs diff 0 OK
- pasted group 3: number 3, name "Test 3: copy of 2", code `egy.mit.ener.rod.coa.e.3` OK
- pasted group 4: number 4, name "Test 4: copy of 2, carbon price 0", code `egy.mit.ener.rod.coa.e.4` OK
- pasted group 5: number 5, name "Test 5: legacy Egypt carbon tax", code `egy.mit.ener.rod.coa.e.5` OK
- pasted group 6: number 6, name "Test 6: fuel price reform", code `egy.mit.ener.rod.coa.e.6` OK
- pasted group 7: number 7, name "Test 7: feebates", code `egy.mit.ener.rod.coa.e.7` OK
- pasted group 8: number 8, name "Test 8: IMF-WB*, High, nominal carbon price", code `egy.mit.ener.rod.coa.e.8` OK
- pasted group 9: number 9, name "Test 9: new ETS with cap (fast estimate) + carbon tax elsewhere", code `egy.mit.ener.rod.coa.e.9` OK
- pasted group 10: number 10, name "Test 10: ETS override blank, no volatility, full auctioning", code `egy.mit.ener.rod.coa.e.10` OK
- scenario 5 (legacy Egypt carbon tax): carbon price [0, 12.5, 50, 112.5] vs legacy row 8582, max abs diff 0 OK
- scenario 5 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 5.86e-13 (at DE804) OK
- scenario 6 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 6.47e-14 (at DQ714) OK
- scenario 7 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 6.47e-14 (at EL714) OK
- scenario 8 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 1.50e-13 (at FN869) OK
- scenario 9 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 2.19e-13 (at GN876) OK
- scenario 10 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 6.22e-14 (at GW714) OK
  - scenario 6, road gasoline new excise ($/GJ) 2026-2031: [0, 1.43, 2.86, 4.289, 5.719, 7.149]
  - scenario 7, road gasoline feebate shadow price ($/GJ) 2026-2031: [0, 0.662, 1.544, 2.427, 3.309, 4.191]
- scenario 7 vs scenario 1 (covered with a non-zero rate: cem, cst, ftr, irn, mac, mch, nfm, oen, omn, rod): other subsectors unchanged (max abs diff 0); covered subsectors in 2039: cem -16.85%, ftr -17.47%, irn -15.61%, oen -21.36%, omn -20.33%, rod -18.59% OK
  - scenario 8: crude oil gp 2024-2026 (real) [84.95, 109.03, 114.14], carbon price 2027-2030 (nominal 20 x infl) [19.581, 19.162, 18.752, 18.35]
- scenario 10 (new ETS on power and industry from 2027, override row blank = carbon price, volatility Zero, policy risk 0, benchmarks 0): nce, atp and fuel use equal scenario 2 (max abs diff 0); ETS cost > 0 (40.68, sum over rows 2039); carbon tax in cement 0 (0) OK
  - scenario 9 (cap -2% in 2027 to -15% in 2035 vs baseline covered emissions; benchmarks 1.0 -> 0.8; volatility Medium), 2026-2036:
    - baseline covered emissions (MtCO2): [0, 130.329, 129.805, 135.332, 137.83, 141.894, 144.906, 148.632, 152.098, 155.881, 159.634]
    - cap (MtCO2): [0, 127.722, 125.1, 128.227, 128.354, 129.833, 130.235, 131.168, 131.755, 132.499, 132.499]
    - volatility adjustment: [0.909, 0.909, 0.909, 0.909, 0.909, 0.909, 0.909, 0.909, 0.909, 0.909, 0.909]
    - effective semi-elasticity: [0, -0.002, -0.002, -0.002, -0.002, -0.002, -0.002, -0.002, -0.002, -0.002, -0.002]
    - fast estimate ($/tCO2): [0, 13.898, 24.651, 35.176, 45.293, 55.179, 64.751, 74.117, 83.243, 92.189, 105.596]
    - permit price used: [0, 13.898, 24.651, 35.176, 45.293, 55.179, 64.751, 74.117, 83.243, 92.189, 105.596]
    - OBR share, cement: [0, 1, 0.975, 0.95, 0.925, 0.9, 0.875, 0.85, 0.825, 0.8, 0.8]
    - covered emissions (MtCO2): [0, 128.524, 126.322, 129.907, 129.764, 130.729, 130.365, 130.375, 129.882, 129.451, 128.475]
    - covered / cap - 1: [0, 0.006, 0.01, 0.013, 0.011, 0.007, 0.001, -0.006, -0.014, -0.023, -0.03]
    - next price proposal: [0, 16.73, 28.722, 40.386, 49.225, 57.448, 65.056, 72.382, 79.379, 86.221, 97.81]
  - volatility adjustment at Medium = 1/1.1 as legacy row 2191 (0.909091) OK
- bco2.chk (baseline data vs scenario 1 CO2) in all 10 groups: max abs 0 OK
- price controls None (global), scenario 1 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 3.10e-13 (at R804) OK
- price controls None (global), scenario 2 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 4.19e-13 (at AR813) OK
- price controls Manual (global), scenario 1 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 1.31e-13 (at P714) OK
- price controls Manual (global), scenario 2 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 7.10e-14 (at AK714) OK
- MTOutputs with blocks 3-10 added by copying (scenario tests): 10 scenario blocks x 58 indicators, 11020 cells equal the Mitigation cell of their output code (max abs diff 0); block numbers 1..10 OK

## 5b. ETS goal seek (ets_goalseek_v2_2.py)
  iteration 1, scenario 3: worst |covered/cap - 1| = 0.0406, alpha 1.000, price 2030 45.29, 2035 92.19
  iteration 2, scenario 3: worst |covered/cap - 1| = 0.0306, alpha 1.000, price 2030 49.22, 2035 86.30
  iteration 3, scenario 3: worst |covered/cap - 1| = 0.0270, alpha 1.000, price 2030 50.49, 2035 85.13
  iteration 4, scenario 3: worst |covered/cap - 1| = 0.0233, alpha 1.000, price 2030 52.40, 2035 82.75
  iteration 5, scenario 3: worst |covered/cap - 1| = 0.0209, alpha 1.000, price 2030 53.89, 2035 81.17
  iteration 6, scenario 3: worst |covered/cap - 1| = 0.0182, alpha 1.000, price 2030 55.30, 2035 79.63
  iteration 7, scenario 3: worst |covered/cap - 1| = 0.0142, alpha 1.000, price 2030 56.54, 2035 78.35
  iteration 8, scenario 3: worst |covered/cap - 1| = 0.0122, alpha 1.000, price 2030 57.64, 2035 77.22
  iteration 9, scenario 3: worst |covered/cap - 1| = 0.0104, alpha 1.000, price 2030 58.59, 2035 76.22
  iteration 10, scenario 3: worst |covered/cap - 1| = 0.0089, alpha 1.000, price 2030 59.43, 2035 75.35
  iteration 11, scenario 3: worst |covered/cap - 1| = 0.0076, alpha 1.000, price 2030 60.15, 2035 74.59
  iteration 12, scenario 3: worst |covered/cap - 1| = 0.0065, alpha 1.000, price 2030 60.77, 2035 73.92
  iteration 13, scenario 3: worst |covered/cap - 1| = 0.0057, alpha 1.000, price 2030 61.31, 2035 73.35
  iteration 14, scenario 3: worst |covered/cap - 1| = 0.0050, alpha 1.000, price 2030 61.77, 2035 72.85
  iteration 15, scenario 3: worst |covered/cap - 1| = 0.0044, alpha 1.000, price 2030 62.16, 2035 72.41
- goal seek, scenario 3 (ETS test settings): worst |covered/cap - 1| 0.0044 after 15 iterations OK
  - price $/tCO2, fast estimate -> goal seek: 2027 13.9 -> 27.8, 2030 45.3 -> 62.2, 2035 92.2 -> 72.4, 2040 159.4 -> 124.3

## 5c. Existing carbon pricing with non-zero test data (v1.09)
- ECP test data, scenario 1 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 2.15e-13 (at S804) OK
- ECP test data, scenario 2 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 2.83e-13 (at AR869) OK
- ECP test data, scenario 3 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 5.81e-13 (at BM869) OK
- ECP test data, scenario 4 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 1.13e-13 (at CF869) OK
- ECP test data, scenario 5 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 2.15e-13 (at CY804) OK
- scenario 2 with the test data: change against the price data xcpd = 0 up to 2024 (max 0), > 0 in 2030 (data 2025 above 2024: max 0.284 $/GJ); level xcp > 0; power xck 2030 max 0.00468 $/kWh; existing carbon pricing revenue 2030 (all, power) [3722.5, 2189.0] USD m; scenario 4 (switched off): xcpd < 0 (min -0.954) OK
- scenario 5 (complementary): new carbon tax rate 2030 = carbon price 20.00 x (1 - existing ETS sector coverage): power 0.00, cement 8.00, road 20.00 OK
- check rows ener.chk, rev.chk, co2.chk in groups 1-5: max abs 0 OK

## 6. Regression vs v1.09 (shared output codes)
- 5556 output codes in both versions (0 intended changes), 2022-2040: max abs diff 0 OK
- variables added: 
- intended changes that differ: none

## 6b. MTOutputs and Charts
- MTOutputs as shipped: 2 scenario blocks x 58 indicators, 2204 cells equal the Mitigation cell of their output code (max abs diff 0); block numbers 1..2 OK
- Charts: 10 line charts (CO2 emissions from fuel combustion, Fuel use, all subsectors and power generation, Fiscal effect vs scenario 1 (total revenue change), Carbon price, Power generation, total, Power CO2 emissions, Variable renewable share of generation, Residential electricity price, Nameplate capacity, total, Electricity demand, total), one series per scenario block on the right MTOutputs rows OK

## 7. Labels and codes
- column H: 2796 cells, 1 formula OK
- column I: 2796 cells, 1 formula OK
- column J: 2796 cells, 1 formula OK
- code column K (scenario 1): 2796 cells, 1 formula OK
- code column AF (scenario 2): 2796 cells, 1 formula OK
- K1061: `egy.mit.ener.rod.gso.e.1` | Fuel use | Road | Gasoline OK
- K342: `egy.mit.sp.pow.coa.a.1` | Supply cost (pre-tax price) | Power | Coal OK
- AF387: `egy.mit.rpb.all.gso.a.2` | Retail price before new policies | All subsectors | Gasoline OK
- K337: `egy.mit.gp.int.oil.1` | International energy price (real; source and adjustment from MTInputs) | International | Crude oil OK
- K8: `egy.mit.infl.1` | Inflation index: US CPI, ResultsYear = 1 (nominal -> real for domestic prices and nominal policy inputs) OK
- AF997: `egy.mit.ets.rod.nga.a.2` | New ETS cost in the fuel price (tax-equivalent, auctioned part) | Road | Natural gas OK
- K427: `egy.mit.esub.all.die.a.1` | Existing consumer subsidy (negative part of txo, positive number) | All subsectors | Diesel OK
- AF199: `egy.mit.ets.p.2` | ETS permit price used OK
- AF1097: `egy.mit.rnew.rod.gso.2` | Revenue from new policies (carbon tax, excise, auctioned ETS, VAT on them) | Road | Gasoline OK
- K3092: `egy.mit.rsub.tra.all.1` | Cost of existing consumer subsidies (positive = cost) | Transport | All fuels OK
- AF3112: `egy.mit.rtot.chg.all.all.2` | Change in total revenue vs scenario 1 (fiscal effect) | All subsectors | All fuels OK
- K1071: `egy.mit.co2.rod.die.e.1` | CO2 emissions from fuel combustion | Road | Diesel OK
- AF3127: `egy.mit.co2.all.nga.e.2` | CO2 emissions from fuel combustion | All subsectors | Natural gas OK
- AF3136: `egy.mit.co2.pct.all.all.2` | Change in CO2 emissions vs scenario 1 | All subsectors | All fuels OK
- AF1043: `egy.mit.atp.rod.gso.e.2` | After-tax price | Road | Gasoline OK
- AF990: `egy.mit.ctxnew.rod.die.a.2` | New carbon tax | Road | Diesel OK
- K1052: `egy.mit.shp.rod.gso.1` | Shadow price on the efficiency margin | Road | Gasoline OK
- K985: `egy.mit.ener.rod.all.e.1` | Fuel use | Road | All fuels OK
- K984: `egy.mit.ener.tra.all.e.1` | Fuel use | Transport | All fuels OK
- AF12: `egy.mit.cptraj.2` | Carbon price trajectory used OK
- K48: `egy.mit.ctcov.oen.all.1` | Carbon tax coverage (Apply tax?) | Other energy use | All fuels OK
- AF3044: `egy.mit.ener.all.all.e.2` | Fuel use | All subsectors | All fuels OK
- K593: `egy.mit.gnc.pow.nga.1` | Generation cost before new policies (cost recovery) | Power | Natural gas OK
- AF482: `egy.mit.caxav.pow.sol.2` | Weighted average capital cost of the stock (vintages) | Power | Solar OK
- K602: `egy.mit.gncav.1` | Average generation cost (weighted by generation shares) OK
- AF625: `egy.mit.rp.res.ecy.2` | End-user electricity price | Residential | Electricity OK
- K749: `egy.mit.msc.1` | Marginal storage cost of variable renewables (wind, solar, other renewables) OK
- AF541: `egy.mit.ccp.pow.oop.2` | New policy cost per kWh generated (carbon tax, ETS, fuel price reform) | Power | Other oil products OK
- K708: `egy.mit.gen.pow.sol.1` | Generation | Power | Solar OK
- AF660: `egy.mit.dem.srv.ecy.2` | Electricity demand | Services (public & private) | Electricity OK
- K649: `egy.mit.mult.tra.ecy.1` | Growth multiple of electricity demand (efficiency, GDP, price; demand shock if switched on) | Transport | Electricity OK
- K3125: `egy.mit.co2.pow.all.e.1` | CO2 emissions from fuel combustion | Power | All fuels OK

## 8. Row outline
- row group buttons below the groups (summaryBelow): OK
- opens rolled up: every grouped row hidden; collapsed flag on the row after each group OK
- Mitigation is the first and active tab: OK; zoom 75%: OK
- carbon price (summary of section 1): 1 rows at level 0, visible OK
- section 1 inputs and paths: 299 rows at level 1, hidden OK
- section 2 retail prices before new policies: 12 rows at level 0, visible OK
- section 2 selectors, gp, sp, txo, vat, etx, esub, esubpu: 96 rows at level 1, hidden OK
- sector totals: 4 rows at level 0, visible OK
- subsector headings and electricity memo lines: 32 rows at level 1, hidden OK
- subsector variables: 1792 rows at level 2, hidden OK
- power: average generation cost and end-user prices: 6 rows at level 0, visible OK
- power: sub-headings and other rows: 459 rows at level 1, hidden OK
- results and electricity memo by sector: 37 rows at level 1, hidden OK
- revenue totals: 9 rows at level 0, visible OK
- revenue by sector: 20 rows at level 1, hidden OK
- CO2 total and change: 3 rows at level 0, visible OK
- CO2 by sector, fuel, scenario 1: 38 rows at level 1, hidden OK

## 9. MTInputs
- columns A:H, rows 1-415: 0 cells differ from the template; D:E hidden OK

## 10. Format
- section bands ['Scenario assumptions (', '1. Policies (scenario ', '2. Retail energy price', '3. Power sector: engin', '5. Transport sector', '6. Buildings sector', '7. Industrial sector', '8. Other energy use', '11. Results - energy c', '12. Revenues (USD mill', '13. Emissions - CO2 fr']: white bold text OK
- Prices_dom: 291 of 291 changed Egypt cells bright yellow with the new values; no other country marked OK
- tabs ['Mitigation', 'MTOutputs', 'StoredResults', 'ScenarioCompare', 'MacroCheck', 'Charts', 'ReadMe', 'LegacyDiff']; sheet LegacyDiff: 40 differences listed OK
- Inputs_prices: data changed by assumption marked bright yellow (columns ['Margin (mit.mar, base year, nominal pric', 'Raw pass-through coefficient (mit.ps, ba', 'VAT applies to final consumers? (assumpt', 'VAT rate assumption = general VAT rate (']; other oil products pass-through and margin) OK

## 11. Sanity and results
- scenario 1 vs 2 before 2027 (fuel use): max abs diff 0 OK
- other energy use not taxed (MCovOen FALSE): fuel use scenario 1 = 2, max abs diff 0 OK
- check row 3073: max abs 0; sector totals and power sum to total: max abs 5.53e-10 OK
- pass-through 0 (nga.pow, nga.res, nga.ind, gso.all, die.all, lpg.all, ker.all, oop.all): retail price before new policies stays at its 2024 value, max abs diff 0 OK

Retail price before new policies, scenario 1 (real $/GJ of ResultsYear), and chosen pass-through:

| Price fuel | pass-through | 2022 | 2024 | 2027 | 2030 | 2040 | VAT 2024 | existing tax 2024 | existing subsidy 2024 | subsidy per price unit 2024 |
|---|---|---|---|---|---|---|---|---|---|---|
| coa.pow | 1 | 13.301 | 5.912 | 5.677 | 5.540 | 5.492 | 0.000 | 0.000 | 0.000 | 0.000 |
| coa.res | 1 | 23.399 | 14.744 | 14.475 | 14.319 | 14.265 | 1.811 | 0.001 | 0.000 | 0.000 |
| coa.ind | 1 | 11.466 | 7.746 | 7.599 | 7.513 | 7.484 | 0.000 | 0.169 | 0.000 | 0.000 |
| nga.pow | 0 | 2.784 | 2.654 | 2.654 | 2.654 | 2.654 | 0.000 | 0.000 | 6.112 | 6.112 |
| nga.res | 0 | 4.634 | 2.331 | 2.331 | 2.331 | 2.331 | 0.000 | 0.000 | 12.056 | 12.056 |
| nga.ind | 0 | 5.336 | 5.087 | 5.087 | 5.087 | 5.087 | 0.000 | 0.000 | 4.483 | 4.483 |
| gso.all | 0 | 17.165 | 9.583 | 9.583 | 9.583 | 9.583 | 0.000 | 0.000 | 12.251 | 0.428 |
| die.all | 0 | 11.095 | 6.682 | 6.682 | 6.682 | 6.682 | 0.000 | 0.000 | 14.189 | 0.529 |
| lpg.all | 0 | 6.791 | 3.414 | 3.414 | 3.414 | 3.414 | 0.000 | 0.000 | 17.647 | 0.455 |
| ker.all | 0 | 10.962 | 6.461 | 6.461 | 6.461 | 6.461 | 0.000 | 0.000 | 11.855 | 0.448 |
| oop.all | 0 | 5.400 | 3.787 | 3.787 | 3.787 | 3.787 | 0.000 | 0.000 | 11.182 | 68.431 |
| bio.all | 1 | 14.471 | 14.476 | 14.476 | 14.476 | 14.476 | 0.000 | 0.000 | 0.000 | 0.000 |
- revenues: sectors and power sum to totals (max abs 1.02e-10); no new-policy revenue in scenario 1 (max abs 0) OK
- CO2: sectors incl. power and fuels sum to the total (max abs 9.66e-13); biomass 0; no change vs scenario 1 before 2027 outside power (max abs 4.26e-13) OK
  - power CO2 before 2027 differs by up to 0.132 MtCO2: investment looks ahead (levelised variable cost includes the future carbon price), as legacy

CO2 emissions from fuel combustion incl. power generation (MtCO2; no process emissions):

| Year | Scenario 1 | Scenario 2 | Change | Transport | Buildings | Industry | Other | Power (scen. 2) |
|---|---|---|---|---|---|---|---|---|
| 2022 | 224.9 | 224.9 | +0.0% | 53.4 | 16.8 | 64.3 | 0.3 | 90.1 |
| 2024 | 297.2 | 297.1 | -0.0% | 74.7 | 25.9 | 86.5 | 0.4 | 109.6 |
| 2027 | 312.8 | 284.1 | -9.2% | 72.3 | 22.0 | 79.1 | 0.5 | 110.3 |
| 2030 | 336.8 | 297.0 | -11.8% | 78.2 | 23.3 | 88.0 | 0.5 | 107.0 |
| 2035 | 382.5 | 334.5 | -12.5% | 87.8 | 25.3 | 103.2 | 0.6 | 117.7 |
| 2040 | 431.8 | 374.1 | -13.4% | 97.7 | 27.3 | 120.0 | 0.7 | 128.3 |

Revenues (USD million, real 2026):

| Year | existing taxes | existing subsidies | new policies | total | change vs scenario 1 |
|---|---|---|---|---|---|
| 2022 (scenario 1) | 25 | 46,656 | 0 | -46,631 | 0 |
| 2027 (scenario 1) | 17 | 53,189 | 0 | -53,172 | 0 |
| 2030 (scenario 1) | 18 | 48,107 | 0 | -48,089 | 0 |
| 2040 (scenario 1) | 22 | 61,726 | 0 | -61,705 | 0 |
| 2022 (scenario 2) | 25 | 46,656 | 0 | -46,631 | 0 |
| 2027 (scenario 2) | 15 | 48,942 | 5,673 | -43,254 | 9,918 |
| 2030 (scenario 2) | 16 | 42,678 | 5,930 | -36,732 | 11,357 |
| 2040 (scenario 2) | 19 | 54,513 | 7,467 | -47,028 | 14,677 |

| Year | Scenario 1 (ktoe) | Scenario 2 (ktoe) | Change |
|---|---|---|---|
| 2022 | 86,439 | 86,439 | +0.0% |
| 2023 | 107,441 | 107,385 | -0.1% |
| 2024 | 112,871 | 112,816 | -0.0% |
| 2026 | 115,426 | 115,374 | -0.0% |
| 2027 | 117,730 | 107,463 | -8.7% |
| 2030 | 127,028 | 112,166 | -11.7% |
| 2035 | 144,876 | 126,824 | -12.5% |
| 2040 | 164,175 | 142,205 | -13.4% |

## 11b. Power sector: properties and results
- scenario 1: no new policy cost in power (max abs 0); variable costs before policy equal in scenarios 1 and 2 (max abs 0); generation shares sum to 1 (max abs 8.88e-16) OK
- generation = required generation from 2023 (max abs 0 MWy); capacity = last year less retirements plus new nameplate investment (max abs 9.09e-11 MW); new effective investment >= 0 and at most the capacity needed (max excess 7.28e-12 MW) OK
  - scale-up limits leave investment unplaced (MW, scenario 1; gas then runs at a higher capacity factor, as legacy): 2023 2,379, 2024 414, 2025 342
- supply cost = average generation cost of the previous year + constant T&D add-on (max abs 5.13e-16) OK
- carbon price (scenario 2) shifts power towards renewables by 2040: VRE share 0.262 -> 0.287, power CO2 146.33 -> 128.30 MtCO2 OK

Power, scenario 1 (baseline) and 2 (carbon price $20 from 2027):

| Year | generation TWh (1) | (2) | gas share (1) | (2) | VRE share (1) | (2) | power CO2 Mt (1) | (2) | price res $/kWh (1) | (2) | capacity GW (1) | investment USD m (1) |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 2022 | 220.1 | 220.1 | 0.797 | 0.797 | 0.052 | 0.052 | 90.1 | 90.1 | 0.0614 | 0.0614 | 60.7 | 2,741 |
| 2023 | 269.3 | 269.3 | 0.828 | 0.827 | 0.049 | 0.050 | 109.8 | 109.6 | 0.0524 | 0.0524 | 60.8 | 23,118 |
| 2025 | 320.8 | 320.8 | 0.702 | 0.701 | 0.096 | 0.097 | 117.8 | 117.7 | 0.0487 | 0.0487 | 85.2 | 29,068 |
| 2027 | 327.7 | 327.7 | 0.612 | 0.608 | 0.133 | 0.136 | 110.8 | 110.3 | 0.0549 | 0.0631 | 100.8 | 5,855 |
| 2030 | 358.8 | 331.7 | 0.607 | 0.592 | 0.160 | 0.156 | 116.1 | 107.0 | 0.0533 | 0.0624 | 107.6 | 8,981 |
| 2035 | 441.6 | 410.8 | 0.586 | 0.556 | 0.218 | 0.230 | 130.6 | 117.7 | 0.0508 | 0.0590 | 132.8 | 11,149 |
| 2040 | 532.5 | 495.6 | 0.566 | 0.522 | 0.262 | 0.287 | 146.3 | 128.3 | 0.0489 | 0.0566 | 162.0 | 12,619 |

Generation cost by type, scenario 1, 2022 and 2040 ($/kWh): Coal 0.197 / 0.115, Natural gas 0.057 / 0.056, Oil 0.190 / 0.172, Nuclear 0.068 / 0.068, Wind 0.072 / 0.072, Solar 0.102 / 0.078, Hydro 0.090 / 0.090, Other renewables 0.078 / 0.078, Biomass 0.260 / 0.260

## 11c. Planned retirement switched on (v1.07)
- planned retirement on (gas 1,500 MW/yr from 2025, oil 400 MW/yr from 2024), scenario 1 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 3.88e-14 (at R714) OK
- planned retirement on (gas 1,500 MW/yr from 2025, oil 400 MW/yr from 2024), scenario 2 vs Python: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, 8 price-fuel and 14 subsector variables, 85 power variables, 45676 cells, max relative diff 3.49e-13 (at AK714) OK
- retirement of gas and oil = the planned data (capped at capacity), max abs diff 0 MW; gas capacity 2030 43,720 MW, new investment 2030 7,123 MW OK

## 11d. Electricity demand memo lines (v1.07)
- scenario 1: memo by subsector = power demand of the subsector (max abs 0); by sector = sum of its subsectors (8e-11); total = total power demand (0); demand by group sums to the total (6.18e-11) OK
- scenario 2: memo by subsector = power demand of the subsector (max abs 0); by sector = sum of its subsectors (9.09e-11); total = total power demand (0); demand by group sums to the total (1.16e-10) OK
  - electricity demand by sector, scenario 1 (ktoe), 2022 / 2030 / 2040: all 15,496 / 25,261 / 37,495; tra 48 / 79 / 126; bld 11,661 / 19,530 / 29,401; ind 3,787 / 5,652 / 7,969; oen 0 / 0 / 0

## 12. Scenario definitions, VBA module, stored results, comparison, macro run
- definition 3 (column L): "Carbon price $20/tCO2 from 2027", Run? Yes, 5 inputs changed from the template OK
- definition 4 (column M): "Package A: carbon tax to $50 by 2030 + feebates (transport, industry)", Run? Yes, 14 inputs changed from the template OK
- definition 5 (column N): "Package B: ETS on power and industry (cap -5% 2027 to -20% 2035) + carbon tax $25 elsewhere", Run? Yes, 11 inputs changed from the template OK
- J4 / K4: "Baseline (stored once)" / "Live (batch target)" OK
- VBA project: module CPATScenarios = CPATScenarios_v2_2.bas (582 lines); document modules ThisWorkbook + 30 sheets matching the sheet code names OK
  - public macros: RunAllScenarios, StoreBaseline, StoreLiveScenario, SolveETSLive, ClearStoredScenarios, RunAll, SolveETS, CheckBatchRun, CheckBatch
- StoredResults: scenario IDs [1, 3, 4, 5], 58 rows each; values only: OK
- stored ID 1 = MTOutputs block 1 of the shipped workbook recalculated (same inputs): max rel diff 0.00e+00 OK
- stored ID 3 = MTOutputs block 2 of the shipped workbook recalculated (same inputs): max rel diff 0.00e+00 OK
- stored ID 4 (Package A: carbon tax to $50 by 2030 + f...) vs independent Python recomputation (17 indicators: CO2, revenues, carbon and ETS price, covered CO2, cap, electricity prices, generation cost): max rel diff 6.33e-13 OK
- stored ID 5 (Package B: ETS on power and industry (ca...) vs independent Python recomputation (17 indicators: CO2, revenues, carbon and ETS price, covered CO2, cap, electricity prices, generation cost): max rel diff 1.18e-14 OK
- stored ID 5 (ETS package): covered emissions meet the cap, worst |covered/cap - 1| 0.0049 OK; ETS price 2027/2030/2035/2040: [123.7, 115.2, 105.9, 157.8]
- ScenarioCompare 2030: IDs [1, 3, 4, 5], names ['Baseline (no carbo', 'Carbon price $20/t', 'Package A: carbon ', 'Package B: ETS on ']; levels, differences and % differences equal StoredResults (max rel diff 2.10e-15) OK
- ScenarioCompare levels: 486 cells, 1 formula OK
- ScenarioCompare differences: 432 cells, 1 formula OK
- ScenarioCompare % differences: 432 cells, 1 formula OK
- ScenarioCompare names: 9 cells, 1 formula OK
  - CO2 2030 (MtCO2): 1: 336.8, 3: 297.0, 4: 251.8, 5: 289.1
- LibreOffice ran the embedded macro: VBAProject.CPATScenarios.RunAll returned 3 in 60.7 s
  - macro stored IDs [1, 3, 4, 5] with the same names; values vs the Python emulation: max rel diff 9.62e-15 at (5, 'mit.nimtot', 2034) OK
  - live column K restored (0 cells differ) and ETS override row empty again OK

## 13. First-run test macro CheckBatchRun (v1.04)
- sheet MacroCheck as shipped: "not run yet" OK
- LibreOffice ran CheckBatch: VBAProject.CPATScenarios.CheckBatch returned 9.237286721121211e-15 in 63.1 s
  - Result: PASS
  - Started: 2026-10-11 18:23:21 in Microsoft Excel 11.0
  - MTInputs: column of scenario 1 (baseline): OK (10)
  - MTInputs: column of scenario 2 (live): OK (11)
  - MTInputs: row of D_NewETS: OK (84)
  - MTInputs: row of D_ETSPriceOverride: OK (93)
  - Mitigation: base-year column of scenario group 2: OK (33)
  - Mitigation: row ets.ovr (ETS override prices): OK (198)
  - Mitigation: row ets.est (fast estimate): OK (197)
  - Mitigation: row ets.next (next price proposal): OK (3158)
  - Mitigation: row co2.gap (covered / cap - 1): OK (3156)
  - Mitigation: row co2.cap (ETS cap): OK (3155)
  - MTOutputs: block of scenario 1: OK (6)
  - MTOutputs: block of scenario 2: OK (66)
  - Stored rows before the run: 232
  - Baseline recalculated here vs stored ID 1: rows compared: 58
  - Baseline: largest relative difference: 4.5e-15
  - RunAll: definitions run and stored: 3
  - Stored cells compared after the run: 4408
  - Cells differing by more than 1E-06 (relative): 0
  - Largest relative difference (baseline and batch): 9.2e-15
  - Finished: 2026-10-11 18:24:24
- first-run test: Result PASS, 12 lookups found, largest relative difference 9.2e-15 OK
- negative test (row code ets.ovr broken): CheckBatch -1 in 1.5 s; Result "FAIL: 1 lookup(s) not found (see below)"; failed lookup: ['Mitigation: row ets.ovr (ETS override prices)'] OK

## 14. Consistency (v1.09, v2.2): check rows, identities, references, versions, stale texts
- check rows in groups [1, 2]: ener.chk 0, rev.chk 0, co2.chk 0, bco2.chk 0 OK
- identities in every column: rnet = rtx - rsub, rtot = rnet + rnew, fuel use = sector totals + power, power fuel use = fuel types of pe, power CO2 = CO2 by type = ETS sector power: max abs 5.7e-10 OK
- file references in the workbook texts (11): current ones exist in the folder: OK
- stale statements about power, revenues or interim data in the workbook texts: none OK
- versions: Settings log ends with 2.2, Mitigation title v2.2, folder README names the workbook and builder of v2.2 OK
- every current versioned file of the folder carries version 2.2 (27 files): OK
- issues register (issues_register_v2_2.py, Word register): 72 items, unique IDs; LegacyDiff rows of type Not yet, Simplification, To check (bucket 5) all included (17 of 17) OK
- one current version of each method note in the folder (8 notes): OK

**Overall: PASS**
