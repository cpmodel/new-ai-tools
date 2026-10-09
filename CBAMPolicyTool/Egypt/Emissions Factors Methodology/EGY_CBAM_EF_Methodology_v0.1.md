# Egypt CBAM goods – direct emission factors split four ways

**Methodology and results, v0.1** (companion to `EGY_CBAM_EF_v0.1.xlsx`, built by `build_ef_v0_1.py`, verified by `recalc_and_check.py`)

| | |
|---|---|
| Scope | Direct (Scope 1) emission factors, tCO₂e per tonne of product, for the eight CBAM goods/routes in the Egypt industry kernel (`CPAT_Industry_Kernel_Egypt_v0.12.xlsx`, sheet `Manual inputs`, rows 30–37) |
| Split | **fc** fuel combustion CO₂ · **fp** fuel-based process CO₂ (reductants / feedstock) · **np** non-fuel process CO₂ (carbonates, electrodes, anodes) · **no** non-CO₂ process gases (N₂O, PFCs) as CO₂e |
| Reference year | 2022 (plant fleet and production structure); factors are technology-based so they drift slowly |
| Status | v0.1 – methodology and audit chain complete; several Egypt-specific inputs are analyst estimates flagged **VERIFY** (see §9) |

---

## 1. Why a four-way split

CBAM reports a single direct-emissions number per tonne, but mitigation analysis needs to know *what kind* of carbon it is, because each component responds to a different lever:

| Component | Code | What it is | Main abatement lever | IPCC 2006 category | CBAM MRV category |
|---|---|---|---|---|---|
| Fuel combustion | **fc** | CO₂ from fuels burned for heat or steam (kiln fuel, reformer burners, EAF burners, dryers) | fuel switching, efficiency, electrification, biomass | 1A2 (Energy – manufacturing) | Combustion emissions (source streams, Method A) |
| Fuel-based process | **fp** | CO₂ from fuel carbon that enters the process as **reductant or feedstock** (NG reformed to H₂ in ammonia/DRI; coke and PCI in a blast furnace) | hydrogen, CCS/CCU, scrap substitution | 2B1 / 2C1 (IPPU – chemical & metal) | Process emissions – mass balance (Method B) |
| Non-fuel process CO₂ | **np** | CO₂ from carbonates (clinker, flux) and consumable carbon not used as fuel (electrodes, charge carbon, anodes) | clinker substitution, novel cements, inert anodes, CCS | 2A1 / 2C1 / 2C3 | Process emissions – mass balance |
| Non-CO₂ process | **no** | N₂O from nitric acid, PFCs (CF₄, C₂F₆) from aluminium anode effects, expressed as CO₂e | catalytic N₂O destruction, anode-effect control | 2B2 / 2C3 | Process emissions – N₂O / PFC |

The four components are **additive**: own-process EF = fc + fp + np + no. The split is an *attribution of the same carbon*, never a change in the total; the Checks sheet enforces carbon balances where attribution is involved (DRI).

## 2. Boundaries

1. **Own-process EF** – the emissions of the installation(s) making the good, per tonne of the good. Boundaries follow the CBAM "aggregated goods category" definitions: crude steel (melt shop; rolling and reheating excluded), grey clinker (kiln only; grinding is electricity), ammonia (synthesis complex), urea, ammonium nitrate (AN), primary (unwrought) aluminium.
2. **Chain EF** – own EF plus **precursor** emissions carried in component-by-component: urea carries 0.570 t NH₃/t; AN carries 0.79 t HNO₃/t and 0.215 t NH₃/t; HNO₃ carries 0.27 t NH₃/t. Component identity is preserved (ammonia *fp* stays *fp* when embedded in urea), so the chain split remains meaningful.
3. **Indirect emissions (Scope 2 electricity)** are out of scope here (the kernel holds them separately in column O).
4. **Purchased lime** for EAFs: its calcination CO₂ is emitted at the lime kiln and is outside the steel boundary (CBAM treats lime as a non-precursor); only limestone charged directly counts in steel *np*.
5. **Internal fuel gases** (BF gas, top gas) are not double-counted: their carbon is already inside the coke/NG carbon counted once, in *fp* or *fc*.

## 3. Attribution rules

### 3.1 fc versus fp – the "purpose of the carbon" rule
A fuel's carbon is **fp** if the fuel is charged *into the process* to act chemically (reducing agent or hydrogen feedstock) and **fc** if it is *burned outside the process stream* to provide heat. Carbon that leaves the plant in the product (C in DRI, in hot metal, in crude steel) is deducted from the step in which it entered and re-appears as *fp* where it is finally oxidised.

Applied:
- **Ammonia** – reformer *feed* gas (→ H₂ + concentrated process CO₂) = fp; reformer *burner* and auxiliary boiler gas = fc. Egypt's plants are all natural-gas steam reformers.
- **DRI (Midrex / HYL)** – the share of plant gas reformed to reducing gas = fp (net of carbon retained in the DRI); the share burned in the reformer burners = fc. The retained DRI carbon is oxidised in the EAF and appears there as *fp* (net of the small carbon content of steel).
- **Blast furnace** (reference only) – coke + PCI carbon net of hot-metal carbon = fp; purchased fuels for coke ovens, sinter and stoves = fc.

*Alternative rule considered and not adopted:* splitting DRI gas by the chemistry of reduction (CO vs H₂) and treating only the CO-path carbon as process. It gives a smaller *fp* but requires plant gas analyses, is not how the IPCC or CBAM treat reformer gas, and would make Egypt's numbers incomparable with the ammonia convention above.

### 3.2 np
Carbonate calcination computed stoichiometrically from oxide content (CBAM Method B: CaO 0.785, MgO 1.092 tCO₂/t oxide, derived from molar masses in the workbook) with a kiln-dust correction; carbon electrodes and charge/injection carbon via CBAM carbon contents (Guidance 3, Table 4-11); aluminium anode consumption net of sulphur and ash, plus anode-baking CO₂.

### 3.3 no
N₂O (nitric acid) and PFCs (aluminium) in kg/t multiplied by the selected GWP set. Default **CBAM/EU MRR set (AR5): N₂O 265, CF₄ 6630, C₂F₆ 11100**; switch `gwp_set = "AR6"` gives 273 / 7380 / 12400.

### 3.4 Urea and the ammonia CO₂ stream (convention switch)
CBAM rule (Guidance 5c): CO₂ consumed in urea synthesis is **counted as emitted at the ammonia installation** and is *not* deducted from ammonia or urea. Hence ammonia *fp* = all feedstock carbon, urea *np* = 0. The national-inventory (IPCC) convention nets the bound CO₂ (−0.733 tCO₂/t urea) and the kernel v0.12 currently uses this (urea np = −0.733). The workbook switch `urea_conv` (CBAM | IPCC) reproduces either; **CBAM is the default** for CBAM-cost analysis.

## 4. Equations (as implemented – every row is a live formula over named inputs)

Notation: `EF_f` = fuel emission factor tCO₂/GJ (IPCC 2006 Vol. 2 Table 2.2 = CBAM Annex VIII); `C→CO₂` = 44.009/12.011 = 3.664.

**P1 DRI (per t DRI)**
- fp = NG_total × feed_share × EF_NG − C_DRI × C→CO₂  (10.5 × 0.75 × 0.0561 − 0.0191 × 3.664 = **0.372**)
- fc = NG_total × (1 − feed_share) × EF_NG = **0.147**
- check: fc + fp + retained C = NG_total × EF_NG ✓

**P2 EAF, DRI route (per t crude steel)**
- fc = NG_EAF × EF_NG (0.6 × 0.0561 = 0.034)
- fp = DRI/t × C_DRI × C→CO₂ − C_steel × C→CO₂ (0.98 × 0.070 − 0.040 = 0.029)
- np = (electrodes × 0.8188 + charge C × 0.8297)/1000 × C→CO₂ + limestone × 0.440 = 0.038
- **DRI-EAF steel** = 0.98 × P1 + P2 → fc 0.178 / fp 0.393 / np 0.038

**P3 EAF, scrap route** – fc = 0.8 × 0.0561 = 0.045; fp = 0; np = 0.044.

**P4 BF-BOF (reference, no Egyptian production)** – fp = [(coke × C_coke + PCI × C_PCI − C_HM) × HM/t − C_steel] × C→CO₂; fc = other fuels × EF_NG; np = flux × 0.440.

**P5 Clinker (per t clinker)**
- fc = thermal × Σ(share_i × EF_i) = 3.5 × 0.0896 = **0.314** (fuel mix coal 45 % / petcoke 40 % / NG 5 % / HFO 2 % / alt. fuels 8 % at 50 % fossil)
- np = (CaO 0.65 × 0.785 + MgO 0.015 × 1.092) × CKD 1.02 = **0.537** (IPCC Tier 1: 0.52)
- memo: cement = (fc + np) × clinker ratio 0.85 = 0.723 tCO₂/t cement

**P6 Ammonia (per t NH₃)** – fp = feed 22.5 × 0.0561 = **1.262**; fc = (35.2 − 22.5) × 0.0561 = **0.712**; total 1.975 (IPCC Table 3.1 band: modern 1.694 – European average 2.104; EU default for Egypt in kernel 2.05).

**P7 Urea** – fc = 2.0 × 0.0561 = 0.112; np = IF(urea_conv = "IPCC", −0.733, 0); chain adds 0.570 × NH₃ EF.

**P8 Nitric acid (per t HNO₃ 100 %)** – no = [abated share × 2.5 + (1 − share) × 7.0] kg N₂O × GWP/1000 = 4.75 × 265/1000 = **1.259**; fc = 0 (net steam exporter); chain adds 0.27 × NH₃ EF.

**P9 Ammonium nitrate** – fc = 2.0 × 0.0561 = 0.112; chain adds 0.79 × HNO₃ chain EF + 0.215 × NH₃ EF.

**P10 Primary aluminium** – np = anode 0.44 × (1 − S 0.02 − ash 0.004) × C→CO₂ + baking 0.05 = **1.623**; no = (CF₄ 0.10 × 6630 + C₂F₆ 0.01 × 11100)/1000 = **0.774**; fc = 2.2 GJ × 0.0561 = 0.123.

## 5. Results (v0.1, CBAM conventions, tCO₂e per tonne)

| Product (kernel row) | fc | fp | np | no | **Own total** | Chain incl. precursors | Kernel v0.12 own | EU default (kernel) |
|---|---|---|---|---|---|---|---|---|
| DRI-EAF crude steel (30) | 0.178 | 0.393 | 0.038 | 0 | **0.609** | 0.609 | 0.730 | 1.82 |
| Scrap-EAF crude steel (31) | 0.045 | 0 | 0.044 | 0 | **0.089** | 0.089 | 0.155 | — |
| BF-BOF steel, reference (32) | 0.168 | 1.257 | 0.053 | 0 | **1.478** | 1.478 | 2.172 | — |
| Grey clinker (33) | 0.314 | 0 | 0.537 | 0 | **0.851** | 0.851 | 0.814 | — |
| Ammonia (34) | 0.712 | 1.262 | 0 | 0 | **1.975** | 1.975 | 1.851 | 2.05 |
| Urea (35) | 0.112 | 0 | 0 | 0 | **0.112** | 1.238 | −0.621 | — |
| Ammonium nitrate (36) | 0.112 | 0 | 0 | 0 | **0.112** | 1.953 | 1.082 | — |
| Primary aluminium (37) | 0.123 | 0 | 1.623 | 0.774 | **2.521** | 2.521 | 2.506 | — |

Exact values and the EU default/benchmark columns are in `Products` and `Summary`; cement memo 0.723 tCO₂/t cement.

**Differences from kernel v0.12 and why**
- *Urea*: kernel nets bound CO₂ (IPCC convention, −0.621 total); v0.1 applies the CBAM rule (0.112). Switch `urea_conv` reproduces the kernel.
- *Ammonium nitrate*: kernel 1.082 appears to include nitric-acid N₂O inside "own"; v0.1 keeps HNO₃ as a precursor (own 0.112, chain 1.953). Choose by whether AN plants are modelled as integrated with HNO₃ – both numbers are available.
- *DRI-EAF*: kernel 0.730 (fp 0.561) vs 0.609 (fp 0.393); the kernel's EF-input columns S:V (5 and 24 GJ/t) are inconsistent with its own H:I values, so the route was rebuilt from Midrex energy use with explicit carbon balance.
- *Ammonia*: 1.975 vs 1.851 – kernel used ~33 GJ/t; v0.1 uses 35.2 GJ/t (fleet mid-point, VERIFY). Sensitivity: ±1 GJ/t = ±0.056 tCO₂/t, all in *fc*.
- *Clinker, aluminium*: within 5 % of kernel; now fully traceable.

## 6. Data hierarchy and confidence

Every input row in `Egypt_Inputs` carries value, low, high, **tier**, source IDs and a note.

| Tier | Meaning | Used for |
|---|---|---|
| A | Egyptian plant data (company reports, CBAM installation reports, CDM PDDs) | none yet |
| B | Egyptian sector statistics (worldsteel, IFA, inventory) | production volumes (from kernel) |
| C | Technology default calibrated to Egypt's known plant fleet | most process parameters |
| D | International / IPCC default | BF-BOF route, N₂O rates, anode S/ash, ammonia total gas, PFC rates |

Source register (`Sources`): S01 IPCC 2006 Vol. 2 Table 2.2; S02 CBAM IR 2023/1773 Annex VIII & Guidance 3 Tables 4-7…4-12; S03 IPCC Vol. 3 Ch. 2; S04 Ch. 3; S05 Ch. 4; S06 AR6 GWPs; S07 Midrex; S08 worldsteel; S08b Ezz Steel; S09 GCCA GNR/IEA; S10 EFMA/IFA ammonia; S11 EFMA nitric acid/AN; S12 IAI; S13 kernel v0.12; S14 analyst fleet assessment; S15 stoichiometry. A research pass (2026) verified all IPCC table values directly from the PDFs; Egypt-specific sources were largely inaccessible (bot-blocked) and remain Tier C/D.

## 7. QA and reconciliation (sheet `Checks`)

1–4 structural: DRI carbon balance = 0; clinker fuel shares sum to 1; feedstock ≤ total gas; share bounds.
5–8 integrity: valid switches; own total = Σ components for every product; no negative component (except urea *np* under IPCC netting).
9–12 plausibility against IPCC Tier 1: ammonia within 1.694–2.104; clinker *np* 0.50–0.55; aluminium *np* 1.45–1.75; DRI total 0.7–1.1 × 0.70.
13–14 information: own totals vs EU default base; implied national direct emissions of the eight goods at kernel production volumes (≈ 50.7 MtCO₂e, of which clinker ≈ 42.6) for comparison with Egypt's BTR inventory (1A2 + 2A1 + 2B1 + 2B2 + 2C1 + 2C3) once entered.

`recalc_and_check.py` forces a full Excel recalculation, scans every sheet for error values, prints all checks and the results tables; the build is accepted only with 0 errors and Overall = OK.

## 8. Audit-trail rules (how to read / extend the workbook)

- Every number is either (i) a cited constant on `Factors`, (ii) a cited or flagged Egypt input on `Egypt_Inputs` (green = input; yellow = VERIFY), or (iii) a formula. No calculated cell holds a typed number.
- Every input and output has a **workbook-level defined name equal to its code** (e.g. `clk_thermal`, `ef_nh3_fp`); formulas read as text. `Derivation` column G shows each formula via FORMULATEXT.
- Component tags (fc/fp/np/no) sit next to each derivation row; `Products` sums by tag, so a reclassification is a one-cell change with a visible trail.
- Settings switches: `gwp_set`, `urea_conv`, `ref_year`. Conventions are echoed on `Summary`.
- Versioning per `NORMS.md`: `_vX.Y` in file name, version log on `ReadMe`, builder script kept alongside. Change an input → edit `INPUTS` in the builder → rebuild → recalc/check → bump version → log.

## 9. Known limitations and verification list (ordered by impact)

| # | Item | Current value | Impact if wrong | Where to get it |
|---|---|---|---|---|
| 1 | Ammonia total gas use (fleet) | 35.2 GJ/t | ±0.06 tCO₂/t per GJ (fc) | Plant CBAM installation reports; IFA benchmarking; MOPCO/Abu Qir/EFC reports |
| 2 | Nitric-acid N₂O abatement share | 50 % | ±0.6 tCO₂e/t HNO₃ (→ ±0.47 on AN chain) | UNFCCC CDM registry (Abu Qir N₂O projects), plant operators |
| 3 | Aluminium PFC rates (Egyptalum technology) | 0.10 / 0.01 kg/t | up to +2.0 tCO₂e/t at IPCC Tier 1 CWPB defaults | IAI anode-effect survey, Egyptalum |
| 4 | Clinker kiln fuel mix and thermal intensity | coal/petcoke-dominated, 3.5 GJ/t | ±0.03 tCO₂/t per 0.3 GJ; coal↔NG swap ≈ 0.13 | GCCA GNR, Egyptian cement producers' CBAM reports |
| 5 | DRI gas use and feed/fuel split | 10.5 GJ/t, 75 % | ±0.04 per GJ; the split moves fc↔fp only | Midrex World DR Statistics, Ezz Steel |
| 6 | No blast furnace in Egypt since EISCO 2021 | assumed | BF-BOF row stays a reference | worldsteel route data, trade press |
| 7 | EU default values / benchmarks | copied from kernel v0.12 | affects comparison columns only | Official Journal (definitive-period defaults) |
| 8 | Production volumes (reconciliation only) | kernel v0.12 | none on EFs | worldsteel, IFA, CAPMAS |

Other limitations: single national factor per route (no plant distribution); alternative-fuel fossil share generic; no biogenic-carbon deduction; electricity-related emissions excluded by design.

## 10. Mapping to the kernel

Workbook `Products!D:G` (own fc/fp/np/no) corresponds one-to-one to `'Manual inputs'!H30:K37`; `Products!J:M` (chain) is the basis for embedded-precursor columns (M). The kernel is **not** modified by this work; adoption of v0.1 values is a separate decision to be recorded in the kernel's version log, noting the urea convention and the AN precursor treatment.

---
*Version log*: v0.1 (2026) – first complete methodology, workbook, checks and research pass; Egypt-specific inputs Tier C/D pending verification.
