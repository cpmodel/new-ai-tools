# CAVEATS - log of completed tasks

Append-only log of every completed task in this repo: what was asked, what went in, what came out, and what the next person must know before building on it. One entry per completed task (or per workbook version increment). **Newest entry at the bottom.** Never edit or delete an earlier entry; if a caveat is later resolved, add a new entry that says so and names the entry it resolves.

The rule that every finished task gets an entry here, and the other bookkeeping that goes with it (workbook version log, task inventory, `TODO.md`), is in [`NORMS.md` section 6](NORMS.md#6-task-completion-caveats-log-and-bookkeeping). Repo map and key documents: [`README.md`](README.md).

## Entry template

Copy the block below to the end of the file and fill it in. Keep it short; point to the task spec, report or yaml notes for detail rather than repeating them.

```markdown
## YYYY-MM-DD - <task id / short name> (<version or increment, if any>)
- **Task:** what was asked and what was done (one or two sentences).
- **Inputs:** files, sheets/ranges, data sources, specs and prior versions used.
- **Outputs:** files created or changed (paths), version numbers, documents updated.
- **Caveats:** placeholders, assumptions, known gaps, deferred items, regression differences that were explained rather than zero, decisions still open. Write "None" only if there genuinely are none.
```

## Caveats recorded before this log existed

Earlier work was logged in the task documents themselves; they remain the detailed record:

- Kernel increments v0.3-v0.11 (Tasks E, A, B, C, K, D-scaffold, H, L): `notes` under `TASK-1` in [`egypt/instructions/instructions-egypt.yaml`](egypt/instructions/instructions-egypt.yaml) and each workbook's `Settings` version log.
- Process semi-elasticities (Task D): [`egypt/supporting/TASK-D_ProcessHalfElasticities_DropIn_v0.1.md`](egypt/supporting/TASK-D_ProcessHalfElasticities_DropIn_v0.1.md) section 4.
- Gap list for the Egypt kernel: [`egypt/instructions/EgyptTaskReference.md`](egypt/instructions/EgyptTaskReference.md).
- LAMBDA/VBA distribution workbook: [`cpat_excel_new/distribution/LESSONS_LEARNED.md`](cpat_excel_new/distribution/LESSONS_LEARNED.md); Python Distribution module gaps vs Excel: [`cpat_excel_new/distribution/README.md`](cpat_excel_new/distribution/README.md).

---

## Entries

## 2026-10-02 - TASK-D: IPCC-based process semi-elasticities (values ready, not applied)
- **Task:** Derive documented, IPCC-based process-emission semi-elasticities ("half-elasticities") for the eight CBAM products of the industry kernel, to replace the undocumented v0.8 `Manual inputs` E40:F47 values. Values are final; applying them to a kernel increment is queued as `TODO.md` T1.
- **Inputs:** IPCC AR6 WGIII Ch. 11 (Table 11.3 technology costs) and Ch. 12 (Table 12.3 2030 potentials) in `_research/`; CBAM Implementing Regulation (EU) 2023/1773 and guidance documents in `_research/`; kernel convention from `CPAT_Industry_Kernel_Egypt_v0.8.xlsx` (`Manual inputs` D38:F47) and the v0.9 layout (`Manual inputs` rows 53-60, selector `E50`); earlier method note in `egypt+mitigation/TASK-2a_AdHocCalculations_Pseudocode_v0.1.md` and `egypt+mitigation/InitialResultsAndIssues/Methodological Note.txt`.
- **Outputs:** `egypt+mitigation/TASK-D_ProcessHalfElasticities_DropIn_v0.1.md` (drop-in spec, sections 2.3 maps onto v0.9+); `egypt+mitigation/ProcessEmissions_CarbonPrice_Response/ProcessEmissions_CarbonPriceResponse_Report.md` and `.docx` (derivation report, Appendix D = drop-in table); `egypt+mitigation/ProcessEmissions_CarbonPrice_Response/ProcessEmissions_CarbonPriceResponse.xlsx` (live derivation, sheet `CPAT_v0.8_Table`); `TASK-D` entry in `instructions/instructions-egypt.yaml`; `EgyptTaskReference.md` item 1b marked partly done; `TODO.md` T1.
- **Caveats:** (1) Not applied: v0.9-v0.11 rows 53-60 hold placeholder anchors, so process response in those versions still equals v0.8. (2) Values are for process emissions only; v0.8 used one table for both fuel ER (rows 352-359) and process ER (363-370) - the applying task must separate them or accept the proxy explicitly. (3) The kernel books DRI reductant and SMR feedstock CO2 as `fp`, so `np` is ~0 for DRI-EAF and ammonia; the response then acts through `fp`. (4) AN assumes an unabated N2O baseline (`K36` = 0.97); use 0.3292 instead of 0.7210 if Egyptian nitric-acid plants are already abated. (5) Urea's own `np` is negative and floored at 0; the 11.6 % applies to embedded NH3 CO2 and is not applied separately. (6) Same values used for Egypt and non-Egypt. (7) IPCC supplies costs and potentials, not elasticities; the 2030 vs long-run split rests on labelled deployment assumptions. (8) USD2019 basis (100 USD2019 ~ 122 USD2024) versus the kernel's nominal/2024 price path - deflate or note when applying.

## 2026-10-02 - Docs: task-completion norm, caveats log and cross-references
- **Task:** Make the task-completion process explicit: every finished task is appended to this file (task, inputs, outputs, caveats), and `README.md`/`NORMS.md` reference the norms, processes and key documents.
- **Inputs:** `README.md`, `NORMS.md`, `TODO.md` (per-task conventions), `instructions/instructions-egypt.yaml`, `instructions/context-egypt.md`, `egypt+mitigation/EgyptTaskReference.md`, `egypt+mitigation/TASK-D_ProcessHalfElasticities_DropIn_v0.1.md`, `cpat_excel_new/distribution/*.md`, current folder listing.
- **Outputs:** `CAVEATS.md` (new, this file); `NORMS.md` (new "Key documents" table and section 6 "Task completion: caveats log and bookkeeping"; legacy workbook path corrected to `cpat_excel_original\`); `README.md` (layout block, "When a task is done", "Key documents", folder sections brought up to date with the current tree).
- **Caveats:** (1) Only Task D is back-filled above; caveats for kernel increments v0.3-v0.11 stay in the yaml `TASK-1` notes and the workbook `Settings` logs. (2) Not changed here and still stale: `TODO.md` header says the latest mainline is v0.9 (disk and yaml say v0.11, with `build_stream2_v2.py`/`..._Stream2_v2.xlsx` also present); `instructions/context-egypt.md` key-files table points at the v0.3 prototype and at `egypt+mitigation/Methodology/` (the methodology note now sits in `egypt+mitigation/InitialResultsAndIssues/`); the yaml, context file and TASK-D spec still write the legacy folder as `original_cpat_excel/` although the folder on disk is `cpat_excel_original/`. (3) Working-tree reorganisation (`cpat_excel_new/standalone-final/` -> `standalone_initial_prototypes/`, legacy folder rename, `NORMS.md`/`TODO.md`/`_research/` untracked) is uncommitted; nothing was committed in this task.

## 2026-10-02 - T2: merge Stream 2 (Tasks F, I, J) into the mainline (v0.12)
- **Task:** Integrate the Stream-2 branch (v0.7 -> `Stream2_v1/v2/v3` = Task F energy CO2, Task I output-based rebate, Task J revenue fund) with the mainline v0.11 (Tasks A-E, H, K, L, D-scaffold) into one workbook, v0.12.
- **Inputs:** `Old\CPAT_Industry_Kernel_Egypt_v0.11.xlsx`; branch `Old\CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v3.xlsx` (reference); builders `build_stream2_v1.py`, `build_stream2_v2.py`, `build_stream2_v3.py` (imported, not modified).
- **Outputs:** `cpat_excel_new\standalone_working_version\CPAT_Industry_Kernel_Egypt_v0.12.xlsx` and `build_v0_12.py`. New sheets Data_EF, Emissions_Industry, Rebate_Industry, Fund_Industry; Check sections F, I, J appended after row 1516; Settings version-log row v0.12. v0.11 and the three branch workbooks moved to `Old\`. Updated `TODO.md` (T2 ticked, header), `NORMS.md`, yaml, `context-egypt.md`, `EgyptTaskReference.md`.
- **Regression:** for every bundle (LEGACY, 1A, 2A, 2B, 3A, 3B, 3C) all Check "Max |" summaries are 0, except the comparison rows 2-4, 584/585 and 1381, which are compared with v0.11 instead. LEGACY, 1A, 2A, 2B and 3A are identical to v0.11 (Mitigation_Industry and Check rows 1-1516, max abs diff 0). Only 3B (theta = 1, rebate) and 3C (phi = 1, fund) change, by design. No error values on any sheet.
- **Caveats:**
  1. **Fund_Industry process abatement adapted.** It now uses the v0.9 per-category ERmax form: wm = Q x cov x max(EF, 0) x ERmax / 1000, wp = wm x e^(-beta P), with 16 rows (np/no). The branch used the v0.7 single-beta form. The fund is still calibrated on the block curves, and the fuel side is unchanged. Fund 3C results therefore differ from the branch file.
  2. **`+o4` added to the process price term.** It is in o118-o125 (np and no) and in o15/o16. In v0.11, o15/o16 are still in the v0.7 form (beta.np from E363:E370 for both np and no, no ERmax), which is inconsistent with o118 (v0.9). This is pre-existing and was not changed here.
  3. **Task F check (d) re-pointed.** It used block o1 (row 645), which cascades to the year row (o29), not to a price. In the branch the check never bound because output was exogenous. Under Task H (v0.10), 2027 scenario-2 coal/gas CO2 is about 3e-6 Mt above scenario 1 at a zero explicit price. The check now conditions on o2 + o3 (rows 646/647). Block o1 itself is unchanged and stays a label cascade.
  4. **Rebate is counted on two paths.** Task I's rebate uses EU benchmark x obrrb, while Task H's output response (v0.10) deducts min(obrrb, price) x covered EF. Task I's ppin is kept as a memo so the rebate does not enter output twice; the two bases are not reconciled.
  5. **ssc rows now hold 1 in every bundle.** Data_Prices `egy.mit.ssc.<sector>.2` is 1 for all bundles (ssc_fund default); it has no effect while shp.<fuel>.ind.2 = 0 (phi = 0).
  6. **Ex-post shortfall in 3C.** In 3C the ex-post shortfall is negative (e.g. -152 USD mn in 2030): actual o155 revenue is below F, because F is phi x pre-fund revenue rev0. This is the Stream-2 design that avoids circularity.
  7. **T1 still not applied.** Rows 53-60 remain placeholders.

## 2026-10-02 - Task G: IPPU replacement and fuel-CO2 reconciliation (v0.13)
- **Task:** Task G (EgyptTaskReference 5.1/5.2). Replace CPAT's proportional IPPU row with the block's process emissions to remove the double count, and reconcile block product fuel CO2 with sector energy CO2. Produces v0.13.
- **Inputs:** `Old\CPAT_Industry_Kernel_Egypt_v0.12.xlsx`. Legacy `cpat_excel_original\CPAT 1.0pre_456_NoPropData.xlsb`, Mitigation rows 7566-7570 / 13883-13887 (IPPU co2/ch4/n2o/fga/tot) and 6590 (ind CO2), values stored. CPAT method: IPPU(t) = IPPU(t-1) x ind CO2(t)/ind CO2(t-1) from 2024.
- **Outputs:** `CPAT_Industry_Kernel_Egypt_v0.13.xlsx` and `build_v0_13.py`. Manual inputs rows 84-95: switch E85 `IppuOther` NONE (default) / CPAT, plus a per-product "fp counted as IPPU" flag (E88:E95; ammonia = 1). New sheet `IPPU_Industry`: CPAT IPPU; CPAT-method IPPU; block np/no/proc; fp as IPPU; non-block remainder; kernel IPPU = block + non-block; double-count correction; sector fuel-CO2 reconciliation. Check section Task G after row 1827 (summary D1831). v0.12 moved to `Old\`.
- **Regression:** for all bundles, every Check "Max |" row is 0 (Task G identities included), and all prior sheets are identical to v0.12 (max diff 0). There are no error values. The IppuOther = CPAT switch test passes. 2030 s2, kernel vs CPAT-method IPPU (Mt): LEGACY 87.67 vs 82.79; 1A/3A 84.75 vs 85.27; 2A/2B 86.74 vs 85.27; 3B 86.99 vs 86.29; 3C 80.01 vs 81.91. Non-block 50.75 in all.
- **Caveats:**
  1. **Report-only.** Kernel IPPU is not yet fed into a national GHG total; this needs the CPAT links (Task M). There is no calibration to CPAT.
  2. **Non-block baseline.** Non-block IPPU in s1 = CPAT IPPU - block. Under NONE it is held at s1 in s2; under CPAT it follows CPAT's s2/s1 ratio.
  3. **CPAT-method s2 is an approximation.** It reproduces CPAT exactly only in s1. In s2 it applies CPAT's rule to the kernel's own industrial energy CO2 change, with the non-kernel sectors (omn, ftr) held at baseline.
  4. **fp flag judgement.** Only ammonia SMR feedstock fp is booked as IPPU (IPCC 2B1; CPAT mch energy CO2 = 0). DRI reductant fp stays in energy.
  5. **Reconciliation gap (Info rows).** Block fuel CO2 net of fp-as-IPPU exceeds kernel sector energy CO2 in every year for cem, nfm and mch, and in 11 (s1) / 16 (s2) years for irn. Cement is about 13.5 Mt (50 Mt clinker x 0.2883) against 9.5 Mt in 2022. Block fuel EFs or CPAT sector energy balances need review; not resolved here.
  6. **Data source.** CPAT values come from the legacy xlsb (as in Task F), not from the Egypt csv run.
  7. **T1 still not applied** (rows 53-60 are placeholders).

## 2026-10-02 - Task M: links to main CPAT and Table 2 assembly (v0.14)
- **Task:** Task M (EgyptTaskReference 5.6). Link the kernel to CPAT national outputs (GHG for coverage %, revenue, deaths avoided, recycling) and assemble Table 2 in the kernel, porting the TASK-2b rebuild composition. Produces v0.14.
- **Inputs:** `Old\CPAT_Industry_Kernel_Egypt_v0.13.xlsx`; `egypt+mitigation\AdHocRebuild\cpat_outputs_egypt_2022_2041.csv` (CPAT runs EG1-EG4, 32 codes); rebuild method `MethodologyNote_v0.1.md`, `build_adhoc_rebuild_v0_1.py` (Results).
- **Outputs:** `CPAT_Industry_Kernel_Egypt_v0.14.xlsx` and `build_v0_14.py`. New sheets `CPAT_National` (csv stored) and `Table2_Industry`: A bundle -> run mapping (D6:D12; 1A/2A EG1, 2B EG2, 3A-3C EG3, LEGACY '-') + active bundle, scope (Scenarios L), summary year; B CPAT series of the mapped run + deltas; C kernel series; D Table 2 columns J K L M N O P Q R T U AR per year; E live summary row + stored 2030 snapshot of all bundles with rebuild v0.1 K. Check section Task M after row 1857. v0.13 moved to `Old\`.
- **Method (deltas only):** K = dGHG_CPAT + (dIPPU_kernel - dIPPU_CPAT) + (dInd_kernel - dInd_CPAT); dInd_kernel = d energy CO2 of the kernel sectors + non-kernel industry change (All sectors: non-kernel baseline x CPAT run's ind2/ind1 - 1; Industry only: 0). J = (All sectors: CPAT energy CO2; Industry only: encov x kernel-sector energy CO2 + flag x block process baseline) / CPAT GHG baseline. P = gross - rebates (Rebate_Industry) - fund F (Fund_Industry); gross = All: CPAT six-fuel carbon-tax receipts + cptraj x encov x (dInd_kernel - dInd_CPAT) + block process revenue; Industry only: cptraj x encov x kernel-sector energy CO2 (policy) + process revenue. Q = CPAT deaths avoided x dEnergy CO2 (kernel composition) / dEnergy CO2 (CPAT). AR = CPAT dNet new revenue - CPAT carbon-tax receipts + P. M/N/O/T/U = block cbcov/cbintch/cbobchu/emrt over baseline emis/cbqch.
- **Regression:** every Check "Max |" row is 0 for all bundles (Task M included), all prior sheets identical to v0.13 (max diff 0), no error values; live row = stored snapshot.
- **2030 results (kernel / rebuild v0.1):** K Mt 1A -27.7/-31.6, 2A -25.7/-27.7, 2B -27.1/-29.3, 3A -9.7/-18.0, 3B -6.6/-12.8, 3C -17.3/-22.4. J % 63.0, 57.3, 57.3, 8.5, 8.5, 8.5. P $bn 6.90, 6.27, 6.24, 0.91, 0.00, -0.37. Q 1441, 1441, 1497, 278, 241, 397. AR $bn 11.10, 10.47, 10.67, 2.67, 1.76, 1.39.
- **Caveats:**
  1. **T1/T3 not done.** TODO T4 asked for Task M after T1/T3; built now on request. Process semi-elasticities are still placeholders, so K, N, T change once T1 is applied.
  2. **Data vintage.** The csv runs differ from the CPAT values stored in the kernel (xlsb; 2030 baseline IPPU 85.97 vs 87.67, industrial energy CO2 88.97 vs 75.38). Only deltas within one source are composed; J mixes csv GHG with kernel-sector levels.
  3. **Smaller K than the rebuild,** mainly for 3x: the kernel's own industry response replaces CPAT's (no kappa, D_obr, F_fund add-ons); kernel-sector energy CO2 is smaller than kappa x CPAT industry CO2 (J 8.5 % vs 14 %).
  4. **Negative P in 3C (-0.37).** Fund F = phi x block revenue (CBAM products), which exceeds kernel-sector carbon revenue because block fuel CO2 > sector energy CO2 (Task G reconciliation gap). Left unclamped.
  5. **LEGACY** is not a Table 2 bundle: national-linked columns are 0; block metrics shown as is.
  6. **Recycling effects:** the csv has no recycling outputs; reported as revenue use (Scenarios K) + AR only.
  7. **Run mapping** is a stored input (Table2_Industry D6:D12); EG3 prices kappa = 0.54 of industry, so its dGHG residual (non-industry) is that run's.
  8. **Snapshot** rows are stored at default Manual inputs; Check compares live = snapshot only in 2030.

## 2026-10-02 - TASK-EF: Egypt CBAM emission factors split four ways (v0.1, values ready, not applied)
- **Task:** A transparent, auditable derivation of Scope 1 emission factors for the eight kernel CBAM goods, split into fuel combustion (fc), fuel-based process/reductant (fp), non-fuel process CO2 (np) and non-CO2 (no).
- **Inputs:** IPCC 2006 Vol.2 Table 2.2 and Vol.3 Tables 2.4, 3.1, 3.3, 4.1, 4.10 and 4.15; CBAM Implementing Regulation (EU) 2023/1773 and its default-value conventions; `Manual inputs` H30:K37 and S:V of kernel v0.12/v0.13; Egypt plant information (Tier C/D).
- **Outputs:** `egypt+mitigation/EmissionFactors/EGY_CBAM_EF_v0.1.xlsx`, `build_ef_v0_1.py`, `recalc_and_check.py` (0 errors, 14 checks OK) and `EGY_CBAM_EF_Methodology_v0.1.md`; `TASK-EF` in `instructions/instructions-egypt.yaml`; `EgyptTaskReference.md` section 1; `TODO.md` T5.
- **Caveats:**
  1. Not applied to the kernel. Adoption is queued as T5.
  2. Conventions differ from the kernel:
     - urea np = 0 under the CBAM rule, against the kernel's -0.733 (the `urea_conv` switch reproduces the kernel);
     - AN keeps HNO3 as a precursor (own 0.112, against the kernel's folded-in 0.97);
     - DRI-EAF is rebuilt from a carbon balance (0.609, against 0.730).
  3. Egypt-specific inputs are Tier C/D and carry VERIFY flags: ammonia GJ/t, the nitric-acid N2O abatement share, Egyptalum PFC rates, the kiln fuel mix, DRI gas use and the EISCO BF closure. Egypt portals (UNFCCC, CDM, GNR, IEA) were not reachable.
  4. Clinker fc 0.314 (> 0.2883) widens the v0.13 cement fuel-CO2 reconciliation gap.
  5. EU default values are copied from the kernel, not re-verified.

## 2026-10-02 - Integrated Egypt methodology document (v1.0)
- **Task:** Bring the whole Egypt thread into one main methodology document. The EF derivation is Appendix A and the process semi-elasticity work (Task D) is Appendix B.
- **Inputs:**
  - `README.md`, `instructions/context-egypt.md`, `CAVEATS.md`, `TODO.md`, `NORMS.md`, `EgyptTaskReference.md`, `instructions/instructions-egypt.yaml`;
  - `TASK-D_ProcessHalfElasticities_DropIn_v0.1.md` and the `ProcessEmissions_CarbonPriceResponse_Report.md` report;
  - `AdHocRebuild/MethodologyNote_v0.1.md`;
  - `EmissionFactors/EGY_CBAM_EF_Methodology_v0.1.md`;
  - the `InitialResultsAndIssues/` docx files (Technical note, MajorIssues, EgyptResultsInitial, MACC audit trail);
  - kernel v0.13 (`Manual inputs` H30:K37 and rows 53-60).
- **Outputs:** `egypt+mitigation/EGYPT_Methodology_v1.0.md` and `.docx`. The docx is generated from the md by a python-docx script, which is not in the repo. Pointers added in `README.md` (Key documents, egypt folder), the `NORMS.md` Key documents table, the `instructions/context-egypt.md` key files, `EgyptTaskReference.md`, and `TASK-METHOD` in `instructions/instructions-egypt.yaml`.
- **Caveats:**
  1. The document describes the current state. The Table 2 results come from the TASK-2b rebuild, which uses kernel EFs, not EF v0.1, and has open flags: kappa(EG3) = 0.54, and eps_Q, P_EU and sigma are placeholders.
  2. Kernel v0.13 still holds the Task D placeholders (T1 not applied), so the App. B values are not yet live.
  3. App. B.5 recommends routing beta to fp for DRI, BF and ammonia, together with np/no. This is a recommendation, not implemented.
  4. The earlier MACC audit-trail docx (ammonia 85 %, AN 89 % at $100) is treated as superseded by Task D.
  5. If the md is edited, regenerate the docx or the two will drift.

## 2026-10-02 - T1 + T5: Task D semi-elasticities and Egypt CBAM EF v0.1 applied (kernel v0.15, AdHoc rebuild v0.2)
- **Task:** Apply the updated emission factors (TASK-EF v0.1) and the IPCC process-emission semi-elasticities (Task D, TODO T1) to both the kernel prototype and the TASK-2b ad hoc rebuild.
- **Inputs:** `Old\CPAT_Industry_Kernel_Egypt_v0.14.xlsx`; `egypt+mitigation\EmissionFactors\EGY_CBAM_EF_v0.1.xlsx` (Products D:G); `TASK-D_ProcessHalfElasticities_DropIn_v0.1.md` and `ProcessEmissions_CarbonPriceResponse.xlsx` (IPCC central ER at $100); `AdHocRebuild\AdHocCalculations_Rebuild_v0.1.xlsx` builder.
- **Outputs:** `CPAT_Industry_Kernel_Egypt_v0.15.xlsx` + `build_v0_15.py` (Manual inputs H30:K37 EFs, M35/M36 embedded NH3, S:V memo, AD/AE sources; E50 = IPCC; rows 53-60 ERmax / P* = 100 / ER*; Check section T1/T5 at row 1881; Settings log row 42). `AdHocRebuild\AdHocCalculations_Rebuild_v0.2.xlsx` + `build_adhoc_rebuild_v0_2.py`, `recalc_and_check_adhoc_v0_2.py` (report `_v0_2_report.txt`), `MethodologyNote_v0.2.md/.docx`, new switch `Inputs!EFSet` (EGY_EF_V01 default / PROTOTYPE). v0.14 moved to `Old\`.
- **Regression:** kernel: every Check "Max |" row 0 for all 7 bundles, no error values (stored Task J v2 table and Task M Table 2 snapshot refilled; baseline-independent rows unchanged except via EFs). AdHoc: ALL PASSED; PROTOTYPE FULL/NOPHASE = v0.11 132/132 (EFSet forced PROTOTYPE); REBUILD = Python mirror 534/534; check 8 (E_base 61.16) is n/a unless EFSet = PROTOTYPE (E_base 2030 now 62.42 Mt).
- **2030 K (Mt):** kernel v0.15 1A -29.8, 2A -25.8, 2B -27.1, 3A -11.8, 3B -8.6, 3C -23.9 (v0.14 -27.7/-25.7/-27.0/-9.7/-6.6/-17.3); rebuild v0.2 -31.8/-27.8/-29.3/-18.2/-12.7/-22.6 (v0.1 -31.6/-27.7/-29.3/-18.0/-12.8/-22.4).
- **Caveats:**
  1. **Bug fixed (kernel):** EF-category rows o15/o16 (Mitigation_Industry 260/261, 659/660) applied the np beta to no and ignored ERmax; now np = F x (1 - exp(-E x P)), no = H x (1 - exp(-G x P)), as in o118-o125. Latent while the placeholders were equal.
  2. **AN:** HNO3 N2O folded into AN own no (0.79 x 1.2588 = 0.9944); ER at $100 = N2O-weighted blend of unabated 0.7210 and abated 0.3292 at the EF workbook's 50 % abatement share: (3.5 x 0.7210 + 1.25 x 0.3292) / 4.75 = 0.6179.
  3. **Urea np = 0** (CBAM rule; was -0.733, floored to 0 in the process factor anyway). Ammonia/urea beta = 0: routing the CCS lever to fp (EF App. B.5) is not implemented.
  4. **Aluminium:** combined ER 0.2164 applied to both np (anode CO2) and no (PFC) instead of the non-additive split 0.1282 / 0.1078.
  5. **P* = 100 in USD2019,** no deflation to the kernel price unit. Fuel ER (rows 40-47) and the fund path unchanged; 3C fund shadow price 2028 falls 72.4 -> 44.5 $/t because the higher beta buys more abatement per dollar.
  6. **Cement fc 0.2883 -> 0.3136** widens the Task G block-vs-sector fuel CO2 gap (info rows in Check still non-zero counts).
  7. **EF VERIFY list** (EF App. A.6) still open; values not final.
  8. Table 2 snapshot rebuild column now "Rebuild v0.2 K"; Task M not otherwise redesigned (T3 and the EG3 full-coverage re-run still pending).

## 2026-10-02 - AdHoc rebuild v0.3: P* deflation and EG3 full-coverage approximation
- **Task:** Build `AdHocCalculations_Rebuild_v0.3.xlsx` without modifying v0.2. Added `PStar=122` for IPCC beta sets and made `KappaMode=SCALE` the default to approximate full industrial coverage for EG3 bundles 3A/3B/3C.
- **Inputs:** `AdHocCalculations_Rebuild_v0.2.xlsx`, `build_adhoc_rebuild_v0_2.py`, `recalc_and_check_adhoc_v0_2.py`, `MethodologyNote_v0.2.md`, `cpat_outputs_egypt_2022_2041.csv`, `prototype_v0_11_results.json`, Egypt EF v0.1 and Task-D IPCC ER100 values.
- **Outputs:** `egypt+mitigation\AdHocRebuild\AdHocCalculations_Rebuild_v0.3.xlsx`, `build_adhoc_rebuild_v0_3.py`, `recalc_and_check_adhoc_v0_3.py`, `recalc_and_check_adhoc_v0_3_report.txt`, `MethodologyNote_v0.3.md` and `.docx`. Verification passed: PROTOTYPE FULL/NOPHASE 132/132 each; REBUILD Python mirror 588/588; no Excel errors.
- **Caveats:** `KappaMode=SCALE` is an approximation of a full-coverage EG3 run pending a CPAT re-run. CPAT fuel receipts are scaled by 1/kappa; CPAT delta net new revenue itself is not independently re-run. `PStar=122` has Medium confidence and applies only to IPCC beta sets; PROTOTYPE keeps /100.

## 2026-10-02 - Worked example results workbook (v0.1, quick draft)
- **Task:** One-example-per-level walkthrough of the ad hoc Table 2 rebuild v0.3 (parameters -> CPAT inputs -> clinker in 1A -> 8-product block -> national 1A metrics -> 3A/3B/3C variants), with a column noting where the main version (kernel v0.15) differs.
- **Inputs:** `AdHocCalculations_Rebuild_v0.3.xlsx` (Inputs, Results), `cpat_outputs_egypt_2022_2041.csv` (2030), `MethodologyNote_v0.3.md`, `EGYPT_Methodology_v1.1.md` section 5.1.
- **Outputs:** `egypt+mitigation\AdHocRebuild\ExampleResults_v0.1.xlsx` (sheets Example, Inputs, Block), builder `build_example_results_v0_1.py`. All 30 check values match v0.3 to 4 dp (max diff 5e-5); no Excel errors.
- **Caveats:** Quick draft for iteration. Inputs copied as values (not linked to v0.3). Only REBUILD / NOPHASE / SCALE / EGY_EF_V01 shown. Kernel column quotes v0.15 headline values and method differences only; kernel steps are not recomputed.

## 2026-10-03 - Kernel v0.16: run-2 fixes (P* deflation, EG3 1/kappa scaling, block fuel CO2 reallocation, 3C fund fixed point)
- **Task:** Apply the four user-agreed fixes to the final prototype (decisions recorded in the session; mirrors AdHoc rebuild v0.3 for items 1-2) and refresh the `Egypt Final results` deliverables.
- **Inputs:** `Old\CPAT_Industry_Kernel_Egypt_v0.15.xlsx`; `egypt+mitigation\AdHocRebuild\rebuild_v0_3_K2030.json` (rebuild v0.3 K for the Table 2 comparison column).
- **Outputs:** `CPAT_Industry_Kernel_Egypt_v0.16.xlsx` + `build_v0_16.py`: Manual inputs F/J 53-60 P* = 122 (Check T1/T5 ER rows use 122); `Table2_Industry` row 54 kappa, rows 59-60 block fuel CO2 reallocation (ek1' = sum max(block, sector)), row 64 non-kernel response /kappa, rows 84/91 coverage and gross revenue at kappa = 1 for Industry only, row 94 deaths /kappa, row 96 AR; `Fund_Industry` rows 425-433 stored fund R* per bundle (fixed point R = post-abatement block payments o155, secant per year, tol 1e-9), row 329 reads it; new Check section; Table 2 rebuild column "Rebuild v0.3 K". `Egypt Final results\`: kernel v0.16, rebuild v0.3, MethodologyNote_v0.3, ResultsComparison_Table2_v0.3, EgyptResultsInitial_UpdatedResults_v0.3_tracked.docx (built by `AdHocRebuild\make_results_page_v0_3.py`), EGYPT_Methodology_v1.2, EGYPT_FinalCaveats_v1.1; superseded files in `Egypt Final results\Old\`.
- **Regression:** every Check "Max |" row 0 for all 7 bundles, no error values; 3C fund fixed point residual 5e-12 (5 iterations); other bundles have no fund.
- **2030 (kernel v0.16 / v0.15):** K 1A -29.0/-29.8, 2A -25.5/-25.8, 2B -26.7/-27.1, 3A -21.1/-11.8, 3B -17.4/-8.6, 3C -32.9/-23.9 Mt; J 3A-3C 18.5/8.6 %; P 1A 6.89, 3A 1.88/0.89, 3B 0.98/-0.02, 3C 0.73/-0.47 $bn; Q 3A 816/278, 3B 746/239, 3C 985/366. Rebuild v0.3 K -31.4/-27.8/-29.3/-25.1/-19.6/-30.6.
- **Caveats:**
  1. 1/kappa scaling is a linear approximation of a full-coverage EG3 CPAT run (non-kernel industry responds at CPAT's average priced-industry rate).
  2. 3C fund = block's own post-abatement payments; non-block industry revenue (0.73 $bn) stays in P. The rebuild sends all industrial revenue to the fund (P = 0): convention difference.
  3. Reallocation removes the block-above-sector excess (cement 19.05 vs 10.91 Mt in 2030) from non-kernel industry; the underlying CPAT energy-balance mismatch is not resolved.
  4. Data-vintage mix remains (stored industry CO2 75.4 vs csv 89.0 Mt): about 1.9 Mt of the prototype-rebuild K gap at kappa = 1, about 3.5 Mt for EG3.
  5. fp routing (ammonia/urea beta = 0) and the block fuel-intensity channel remain unimplemented (kept as caveats).
  6. The stale `EgyptResultsInitial_UpdatedResults_tracked.docx` (v0.1) may still be open in Word; move it to `Egypt Final results\Old\` if it remains in place.

## 2026-10-02 - Repo tidy: final deliverables vs Old/ (no model changes)
- **Task:** Tidy the Egypt output and working folders so only the final deliverables sit at top level; everything else moved to `Old/`. The prototype kernel v0.16 (not the ad hoc rebuild) is treated as the final model.
- **Inputs:** `Egypt Final results\`, `egypt+mitigation\`, `cpat_excel_new\standalone_working_version\`.
- **Outputs:** `Egypt Final results\` and `egypt+mitigation\` now both hold the identical final set at top level (hash-checked): `CPAT_Industry_Kernel_Egypt_v0.16.xlsx`, `EGYPT_Methodology_v1.2`, `EGYPT_FinalCaveats_v1.1`, `ResultsComparison_Table2_v0.3`, `EgyptResultsInitial_UpdatedResults_v0.3_tracked.docx` (+ `EgyptTaskReference.md` in the working folder). Moved to `Egypt Final results\Old\`: rebuild v0.3, MethodologyNote_v0.3, `CAVEATS.md` snapshot; exact duplicates of files already in `Old\` removed. Moved to `egypt+mitigation\Old\`: `AdHocRebuild\`, `EmissionFactors\`, `InitialResultsAndIssues\`, `ProcessEmissions_CarbonPrice_Response\`, TASK-2a/TASK-D specs, Methodology v1.0/v1.1, ResultsSummary v1.0, kernel v0.1. `standalone_working_version\`: builders other than `build_v0_16.py` moved to `Old\`, `__pycache__` deleted; `build_v0_16.py` paths updated (adds `Old\` to `sys.path`, reads v0.15 and `rebuild_v0_3_K2030.json` from their `Old\` locations; imports verified). Removed an empty stray root folder named `EGYPT_FinalCaveats_v1.1.md\`. Paths updated in `README.md`, `NORMS.md`, `TODO.md`, `instructions\context-egypt.md`.
- **Caveats:**
  1. `EGYPT_Methodology_v1.2` Appendix C (file map) and earlier entries in this log / `instructions-egypt.yaml` still cite pre-tidy paths; prefix `Old\` for anything not in the final set.
  2. `egypt+mitigation\EGYPT_ResultsSummary_v1.0.docx` was locked (open in Word); a copy is in `Old\`, delete the top-level one once closed.

## 2026-10-02 - Repo tidy follow-up
- **Task:** Resolve both caveats of the previous entry.
- **Inputs:** EGYPT_Methodology_v1.2.md/.docx.
- **Outputs:** Appendix C paths now point to Old\ in the .md and .docx (table text only, formatting kept), copied to both Egypt folders; locked egypt+mitigation\EGYPT_ResultsSummary_v1.0.docx removed (copy in Old\).
- **Caveats:** Earlier log entries and instructions-egypt.yaml still cite pre-tidy paths (historical record).

## 2026-10-02 - Updated results text v0.4 (final prototype values)
- **Task:** Put the final prototype (kernel v0.16) numbers into the tracked-changes results text; v0.3 carried the ad hoc rebuild values.
- **Inputs:** Egypt Final results\Old\EgyptResultsInitial_UpdatedResults_v0.3_tracked.docx; kernel v0.16 2030 values from ResultsComparison_Table2_v0.3.md.
- **Outputs:** EgyptResultsInitial_UpdatedResults_v0.4_tracked.docx in both Egypt folders (builder egypt+mitigation\Old\AdHocRebuild\make_results_page_v0_4.py); tracked changes still relative to the original text; Table 2 checked against v0.16 (P, K, J, N, O, Q). References updated in README, FinalCaveats v1.1, Methodology v1.2, ResultsComparison v0.3. v0.3 moved to Old\.
- **Caveats:** O uses NOPHASE (as Table 2); the prototype's default FULL roughly doubles it. 2A/2B intensity shows 0% because the prototype holds block fuel intensity fixed (stated in the text). Narrative re-ranked: 3C now has the largest cut (32.9 Mt), 2B the largest health gain (1,473). Methodology v1.2 section text still describes the rebuild-based results page.

## 2026-10-02 - CBAM carve-out of Table 2 (v0.1)
- **Task:** Original CPAT run results everywhere except the CBAM block; block from kernel v0.16 with growth matched to CPAT sectors. 3A-3C use EG3 as run (no 1/kappa).
- **Outputs:** EGYPT_CarveOut_Table2_v0.1.md/.docx in both Egypt folders; builder egypt+mitigation\Old\AdHocRebuild\make_carveout_v0_1.py with kernel extracts carveout_kernel_*_2030.json (Excel COM, Settings!B10 per bundle).
- **Result (2030 K, Mt):** 1A -36.1, 2A -31.7, 2B -33.4, 3A -21.5, 3B -17.8, 3C -31.9.
- **Caveats:** K = CPAT dGHG - r x block + kernel block change, with r = CPAT industry energy CO2 % change (also CPAT's IPPU %). Non-CBAM IPPU (about 50 Mt) keeps CPAT's response; prototype v0.16 held it fixed (main cause of its low 1A). Growth rescaled 2024-2030: irn x0.976, cem x1.001; mch (mining & chemicals = fertilisers) x1.157 and nfm x0.956 follow total CPAT industry because CPAT sector energy is 0 for Egypt. M/N/O kept from v0.16 (O NOPHASE). Deaths scaled by block fuel adjustment. 1A remains 5.5 Mt below the original: about 2.7 Mt was the original's process double count, the rest is the block responding less than CPAT's industry average (fuel intensity fixed). Tracked doc v0.4 not updated to the carve-out.

## CBAM carve-out v0.2 (CPAT fuel-intensity response), updated text v0.5, Methodology v1.3 (2026-10-03)

- **Task:** Final Table 2 = original CPAT runs everywhere; only the CBAM block replaced, using CPAT's own emissions-intensity reduction plus the CBAM output/process adjustment, with no double counting.
- **Method:** K = dGHG_CPAT - r*B + dB. CPAT's implied block change (r*B) removed in full; dB = kernel output response + kernel process abatement (where charged) + CPAT fuel-intensity response (1+r)^(1/3)-1 (s_int = 1/3, Methodology section 3): EG1 -4.9%, EG2 -5.1%, 3A-3C use -4.9%. N/O = kernel v0.16 plus the fuel-intensity term (O x0.8 NOPHASE; 3B x1.0). Check: the N formula reproduces kernel N to within 0.2 pt with i_f = 0.
- **Results 2030 (1A/2A/2B/3A/3B/3C):** K -37.4/-33.1/-34.8/-22.8/-19.1/-33.0; P 6.9/6.3/6.2/1.5/0.6/0.4; N -5.8/-2.1/-2.3/-5.8/-5.8/-23.1; O -24.0/-12.2/-12.3/-24.0/-5.1/-36.2; Q 1,501/1,456/1,517/552/491/646; J 63/57/57/14/14/14.
- **Outputs:** EGYPT_CarveOut_Table2_v0.2.md/.docx (builder Old\AdHocRebuild\make_carveout_v0_2.py); EgyptResultsInitial_UpdatedResults_v0.5_tracked.docx (make_results_page_v0_5.py, tracked relative to the original; rankings rechecked); EGYPT_Methodology_v1.3 (new section 5.0). Both Egypt folders identical. v0.4 text, carve-out v0.1 and Methodology v1.2 moved to Old\. References updated in README, NORMS, TODO, context-egypt.md, FinalCaveats v1.1, ResultsComparison v0.3.
- **Caveats:**
  1. The 3B rebate and the 3C fund are not re-solved for the lower block fuel payments (under 0.05 $bn).
  2. O on NOPHASE, not yet confirmed by the user.
  3. 3A-3C use EG3 as run (54% industry coverage), so their non-block industry cut is diluted.
  4. The CPAT intensity share (1/3) is the CPAT efficiency/fuel-switch split; it is not product-specific.

## CBAM obligations on FULL convention: carve-out v0.3, updated text v0.6 (2026-10-03)

- **Task:** Report CBAM obligations (O) on the FULL convention (EU CBAM factor 0.485 in 2030, full credit for the domestic price; kernel headline) instead of NOPHASE.
- **Method:** O = kernel v0.16 FULL value + 0.59 x the fuel-intensity change in N ((48.5 - 20)/48.5); 3B x1.0. NOPHASE kept as a memo row. All other values as in carve-out v0.2.
- **Results 2030 O FULL (1A/2A/2B/3A/3B/3C):** -44.2/-22.9/-23.0/-44.2/-5.0/-53.1 (NOPHASE -24.0/-12.2/-12.3/-24.0/-5.1/-36.2).
- **Outputs:** EGYPT_CarveOut_Table2_v0.3.md/.docx (make_carveout_v0_3.py); EgyptResultsInitial_UpdatedResults_v0.6_tracked.docx (make_results_page_v0_6.py; O cells and O narrative only); Methodology v1.3 section 5.0 edited in place. v0.5 text and carve-out v0.2 moved to Old\. Both Egypt folders identical; references updated.
- **Caveats:**
  1. The FULL percentages are not like-for-like with the original Table 2, which used NOPHASE; the comparison table shows both.
  2. The results text does not state the convention explicitly; add a footnote to Table 2 if needed.

## Egypt finalisation: final set docx-only, FinalCaveats v1.2, convention note (2026-10-03)

- **Task:** Finalise the Egypt deliverables, archive the old versions and the .md files, and tie off loose ends.
- **Final set (top level of both `Egypt Final results\` and `egypt+mitigation\`, hash-identical):** EGYPT_CarveOut_Table2_v0.3.docx (final Table 2), EgyptResultsInitial_UpdatedResults_v0.6_tracked.docx, EGYPT_Methodology_v1.3.docx, EGYPT_FinalCaveats_v1.2.docx, CPAT_Industry_Kernel_Egypt_v0.16.xlsx. The working folder also keeps EgyptTaskReference.md (workflow tracker, per NORMS).
- **Changes:**
  - FinalCaveats v1.2 adds section F (carve-out caveats F1-F7) and updates items 8-9.
  - The results text v0.6 now says obligations are "under the 2030 CBAM phase-in" (tracked).
  - The .md sources and ResultsComparison_Table2_v0.3 moved to Old\. make_carveout_v0_3.py now writes its md to Old\.
  - References updated in README, NORMS, TODO, context-egypt.md and instructions-egypt.yaml (FINAL note under TASK-1).
- **Caveats:**
  1. The egypt+mitigation\Old copy of FinalCaveats v1.1 was restored from the Egypt Final results\Old copy, which has the older file references (superseded anyway).
  2. The kernel xlsx could not be hash-checked because it is open in Excel.

## Kernel v0.17: CBAM carve-out in the kernel (2026-10-03)

- **Task:** Put the final Table 2 (CBAM carve-out v0.3, O on FULL) into the kernel with live formulas.
- **Outputs:** `CPAT_Industry_Kernel_Egypt_v0.17.xlsx` + `build_v0_17.py` (Excel COM, from v0.16). New sheet `CarveOut_Table2`:
  - A: inputs (live links to Table2_Industry, Manual inputs, CPAT_National, Emissions/IPPU_Industry);
  - B: rates r and i_f;
  - C: block by sector (mch, irn, nfm, cem; growth factors);
  - D: carve-out;
  - E: Table 2 columns for the active bundle (Settings!B10);
  - F: stored 6-bundle snapshot and a live-vs-stored check.
  - Also: a Check sheet section, a pointer in Table2_Industry B112, and a Settings version row. No other formulas changed.
- **Copies:** in both `Egypt Final results\` and `egypt+mitigation\` (hash-identical). v0.16 and `build_v0_16.py` moved to `Old\`.
- **Verification:** each bundle set active in turn, CalculateFull. Live = stored within 0.006 (largest in 3A-3C J: kappa 0.537 is rounded in the script). Kernel N rebuilt from the block table is within 0.19 pt.
- **Fix:** Methodology v1.3 section 5.0 and FinalCaveats v1.2 section F had symbols (minus, delta, times, section sign, kappa, CO2 subscript) saved as "?" (encoding slip in the earlier edit). Repaired, and both docx regenerated.
- **References updated:** README, TODO, context-egypt.md, instructions-egypt.yaml, Methodology (5.0, 6, App. C, version log), FinalCaveats header.
- **Caveats:**
  1. The NOPHASE memo uses the stored v0.16 NOPHASE O per bundle (hard input, rounded to 0.1), not a live convention switch.
  2. With Settings!B10 = LEGACY the carve-out sheet shows n/a; pick a bundle from 1A to 3C.

## 2026-10-04 - Tracked results text v0.7 (terminology; author Stephen Stretton)
- `EgyptResultsInitial_UpdatedResults_v0.7_tracked.docx` (both Egypt folders) built from v0.6 by `egypt+mitigation/Old/AdHocRebuild/make_results_page_v0_7.py`; v0.6 moved to `Old/`.
- Fixes: 3C CBAM obligation cut 30% -> 53% (summary paragraph, missed in v0.6); 3B "undermines both revenue and CBAM relief" -> "raises less revenue and delivers the smallest CBAM relief"; 3C "forgoes revenue and health benefits" -> "raises less revenue and yields smaller health benefits than broad-based pricing"; added note that EG3 prices only about half of industrial energy CO2 (14% coverage, not about 21%).
- Terminology standardised: "industrial abatement rebate", "CBAM-sector emission intensity", "CBAM-sector emissions", "national GHG emissions"; en-dashes in numeric ranges.
- All tracked changes and docProps (creator/lastModifiedBy) of the final .docx files now attributed to Stephen Stretton; Lelia Croitoru's comments kept. References v0.6 -> v0.7 updated (README, instructions, Methodology/FinalCaveats md; docx regenerated).
- Consistency check: final numbers identical across tracked text, CarveOut v0.3, Methodology v1.3 §5.0 and kernel v0.17 snapshot; both folders hash-identical. Not verified by opening in Word (automation hung); XML validated.


## 2026-10-04 - Egypt final v1.0 release (Stephen Stretton)

- All final deliverables relabelled v1.0 and placed identically (hash-checked) at the top of `Egypt Final results\` and `egypt+mitigation\`: CPAT_Industry_Kernel_Egypt_v1.0.xlsx (= v0.17 relabelled, no formula changes), EGYPT_CarveOut_Table2_v1.0, EgyptResultsInitial_UpdatedResults_v1.0_tracked, EGYPT_Methodology_v1.0, EGYPT_FinalCaveats_v1.0. Numbers unchanged (2030 K: -37.4/-33.1/-34.8/-22.8/-19.1/-33.0 Mt).
- Proofread: tracked text fixes (double spaces, punctuation, "CBAM obligations", "trade-offs", 3C label, table-header units), all revisions and docProps by Stephen Stretton. Methodology: final-status header, §5 pointer to §5.0, coverage row in §5.0, kernel v0.17/v1.0 rows, file map updated.
- Md sources in `Old\` (both folders); earlier drafts and md (Methodology v1.3, FinalCaveats v1.2, CarveOut v0.3, text v0.7, kernel v0.17) archived in `Old\Superseded\` and `standalone_working_version\Old\`.
- Not validated by opening in Word (Word COM hangs in this environment); XML checked. Open items unchanged: full-coverage EG3 CPAT run; 3B published deaths (345) untraced; 3B rebate and 3C fund not re-solved.


## 2026-10-05 - Egypt results text v1.1: single scenario term (Stephen Stretton)

- Reviewer comment: the chapter mixed "policy options", "families", "bundles", "choices" and "designs" for the scenarios. `EgyptResultsInitial_UpdatedResults_v1.1_tracked.docx` (builder `egypt+mitigation\Old\AdHocRebuild\make_results_page_v1_1.py`) makes 22 minimal tracked replacements, all by Stephen Stretton, so that every reference to 1A-3C uses "scenario(s)" (for example "Policy bundle" -> "Policy scenario", "The second family" -> "The second set of scenarios", "The six bundles" -> "The six scenarios"). Uses that do not refer to scenarios are kept: "Revenue Recycling Option", "recycling options/choice(s)", "policy design". No numbers changed.
- v1.0 text archived in `Old\Superseded\` (both folders); v1.1 copied to both folders (hash-identical). Methodology and FinalCaveats v1.0 updated only to point to the v1.1 text file. Not opened in Word (COM hangs); XML checked.


## 2026-10-05 - Table 2 workbook with live CBAM obligations (Stephen Stretton)

- New `EGYPT_Table2_Final_CBAMcalc_v1.0.xlsx` (both Egypt folders; builder `egypt+mitigation\Old\AdHocRebuild\make_table2_cbam_xlsx.py`, data `kernel_products_2030.json` from `extract_kernel_products_2030.py`). Rows J, P, K, M, N, Q are hard-coded from the CBAM carve-out v1.0. Row O is computed live by product: obl = max(0, CBF x P_EU x EI - S x d), weighted by 2024 EU exports.
- Inputs are on the Assumptions sheet: EU price in 2030 (USD 100, placeholder kept from the initial analysis); convention (default **Full implementation, CBF = 1**, as in the initial table; alternative 2030 phase-in, CBF = 0.485); deduction scaled by CBF (No); net the 3C fund (No).
- Product-level O reproduces the kernel's O exactly (all six scenarios). It differs from the carve-out v1.0 linear shortcut by <= 0.4 pt. Default O: 1A -24.3, 2A -12.5, 2B -12.6, 3A -24.3, 3B -5.4, 3C -36.6 (2030 phase-in: -44.4/-23.1/-23.2/-44.4/-5.4/-53.4). Sensitivities: net the 3C fund -> 3C -20.7; P_EU 120 -> 1A -21.2; P_EU 80 -> 1A -29.1.
- **Pending:** the v1.0/v1.1 docx (CarveOut, tracked text, Methodology, Caveats) still quote O on the 2030 phase-in convention. Update them once the convention is agreed.


## 2026-10-06 - CBAM obligations restated as CBAM-product intensity change (Stephen Stretton)

- Table 2 row O is now the change in CBAM obligations per tonne exported = export-weighted change in embedded (fuel + process) emission intensity of CBAM products only (2024 EU export weights). The Egyptian carbon-price deduction is not applied, so the EU price and phase-in cancel. Final O: 1A -5.4, 2A -2.6, 2B -2.7, 3A -5.4, 3B -5.4, 3C -20.7% (was FULL -44.2/-22.9/-23.0/-44.2/-5.0/-53.1).
- Updated: `EGYPT_Table2_Final_CBAMcalc_v1.0.xlsx` (O live, deduction variants as memo rows only); results text **v1.2** tracked (`make_results_page_v1_2.py`; O cells, range "3 to 21%", 1A/3A 5%, 2A/2B 3%, 3C 21%, 3B sentence); CarveOut, Methodology §5.0, FinalCaveats F5/item 9; README and instructions references. v1.1 text and `EGYPT_CBAM_ObligationNote_v0.1` (deduction-based) moved to `Old\Superseded\`.
- Not changed: kernel v1.0 sheet `CarveOut_Table2` still reports O on FULL (the workbook governs). The user's `Egypt Final results\EGYPT_Methodology_v1.1 - Edited.docx` and the renamed note in `CBAM Guess - Needs Carolyn Input\` were left untouched and still quote the old O.

## 2026-10-07 - All Egypt deliverables aligned at v1.3

- Methodology v1.3 is built from the author's edited master (`EGYPT_Methodology_v1.1 - Edited.docx`, archived in `Egypt Final results/Old/`). Edits: §2 output and §4.5 now define Table 2 O as intensity-only (export-weighted, no deduction); the kernel's deduction-based Task L obligation and the NOPHASE/FULL/SCALED conventions are kept only as a memo; the P_EU/CBF parameter row and the glossary entry are marked memo-only; "bundle" changed to "scenario" (as in the tracked text). The edited master has no Markdown source, so it was edited directly in Word XML.
- CBAM obligation note rewritten to v1.3 (intensity-only; source `egypt+mitigation/Old/EGYPT_CBAM_ObligationNote_v1.3.md`, via pypandoc).
- CBAM workbook, CarveOut, FinalCaveats, tracked text and kernel are all at v1.3; content is unchanged apart from version references and "scenario" terminology. Kernel v1.3 `CarveOut_Table2`: column M80:M85 stores the intensity-only O; row 71 is relabelled as the FULL deduction memo (live O is still FULL there; the workbook governs).
- v1.0/v1.2 copies archived (FIN `Old/`, WRK `Old/Superseded/`). FIN and WRK copies are byte-identical. README, instructions and EgyptTaskReference now point to v1.3.

## 2026-10-08 - Egypt folder reorganisation and doc consistency fixes
- **Task:** merged `Egypt Final results/`, `egypt+mitigation/` and `instructions/` into a single root folder `egypt/` (`final/`, `supporting/`, `archive/`, `instructions/`); fixed documentation inconsistencies (latest kernel is v1.3 everywhere; legacy folder is `cpat_excel_original/`; methodology master is the edited `.docx`).
- **Inputs:** the three old folders (byte-identical duplicates found by hash; the two folders held the same v1.3 deliverables and mirrored archives).
- **Outputs:** `egypt/` (new `README.md`); duplicates removed, one copy of each file kept; tracked `__pycache__` files removed; path references updated in README, NORMS, TODO, AGENTS, the instructions yaml/context, the TASK-D/TASK-2a specs and the Python builders (`egypt/supporting/AdHocRebuild/*`, `cpat_excel_new/standalone_working_version/Old/build_v0_14-17.py`). Entries above this one still quote the old paths: map `Egypt Final results\`/`egypt+mitigation\` to `egypt\final\`, `Old\` and `Old\Superseded\` to `egypt\archive\` or `egypt\supporting\`.
- **Caveats:** builders were not re-run (they need Excel COM on Windows); path edits are untested. Final CBAM workbook/note renamed to `..._v1.3_NeedsCarolynConfirmation` (from the `CBAM Guess - Needs Carolyn Input/` folder name). Two variants that differed were kept with `_alt1` (FinalCaveats v1.0, Methodology v1.0 md) and the pre-edit Methodology v1.3 renamed `_preedit`, all in `egypt/archive/`; the two `CAVEATS.md` copies from the old `Old/` folders are `egypt/archive/CAVEATS_snapshot_*.md`. `extract_kernel_products_2030.py` still points at the (now archived) v1.0 kernel and old Windows root. Historical versions in `egypt/archive/` were left with their original path text.

## 2026-10-04 - Decisions: 3B rebate scope and O convention (no model change)
- **Task:** resolved the two open TODO T4 decisions on the user's instruction (Stephen Stretton): the 3B rebate goes to all covered industry (`Inputs!ThetaOther` = 1, P approx. 0); the CBAM obligation O is on the FULL convention (CBAM factor phase-in, domestic carbon price deducted in full). NOPHASE (CBF = 1) is kept as a memo only.
- **Inputs:** `TODO.md` T4; `build_v0_11.py` / `build_v0_17.py` convention definitions; CAVEATS 2026-10-03 kernel v0.17 entry (kernel already reports O on FULL).
- **Outputs:** `TODO.md` T4 item updated; this entry. No workbook, builder or yaml changes.
- **Caveats:** decisions only, not yet applied. The AdHoc rebuild defaults (`ThetaOther`, convention) must be set to these values in a new builder version and re-verified with `recalc_and_check_adhoc.py` (Windows/Excel COM). Any Table 2 numbers built under `ThetaOther` = 0 or NOPHASE are superseded once that is done.

## 2026-10-04 - AdHoc rebuild v0.4 builder and verifier drafted (not built, not run)
- **Task:** applied the 2026-10-04 decisions to the AdHoc rebuild defaults: `Conv` = FULL (NOPHASE kept as a memo) and `ThetaOther` = 1. `ThetaOther` is forced to 0 in PROTOTYPE mode so the v0.11 reproduction is unchanged. No formula changes.
- **Inputs:** `build_adhoc_rebuild_v0_3.py`, `recalc_and_check_adhoc_v0_3.py`; CAVEATS 2026-10-04 decisions entry.
- **Outputs:** `egypt/supporting/AdHocRebuild/build_adhoc_rebuild_v0_4.py` and `recalc_and_check_adhoc_v0_4.py` (docstrings, defaults, ReadMe version log; the verifier now saves with Conv = FULL). v0.3 files untouched. Both compile (`py_compile`).
- **Caveats:** drafted on Linux. `AdHocCalculations_Rebuild_v0.4.xlsx` has NOT been built, and the verifier (Excel COM, Windows) has NOT been run. Run `python build_adhoc_rebuild_v0_4.py` then `python recalc_and_check_adhoc_v0_4.py` and check the report. Still to do after a clean run: MethodologyNote v0.4 and ResultsComparison_Table2 v0.4 (results change with ThetaOther = 1 and FULL), the yaml / `EgyptTaskReference.md` / `context-egypt.md` bookkeeping, and whether the final Table 2 deliverables need regenerating.

## 2026-10-04 - T3 builder drafted (kernel v1.4, not built, not run)
- **Task:** drafted `build_v1_4.py` for T3: move 2024 production, growth, EU exports and pre-policy prices out of `Mitigation_Industry` (both blocks) into a new `Manual inputs` section and link them. Values are read from the workbook, so no number should change.
- **Inputs:** `CPAT_Industry_Kernel_Egypt_v1.3.xlsx` (layout and values inspected read-only with openpyxl: both blocks identical, rows 97-108 of `Manual inputs` empty); TODO T3 spec (corrected: the spec's row 62 and v0.9 references are stale).
- **Outputs:** `cpat_excel_new/standalone_working_version/build_v1_4.py`; TODO T3 note. No workbook produced.
- **Caveats:** drafted on Linux, never run; Excel COM needed. New section at `Manual inputs` rows 98-108 (title 98, header 100, products 101-108); links F/G/I rows 276-283 and 675-682, G rows 288-295 and 687-694; total exports (column H, blank placeholder) deliberately not linked. The builder aborts without saving unless the max abs difference over pre-existing cells is 0 and the new Check literal count is 0, then writes the version-log row and moves v1.3 to `Old/`. Check sheet rows 595-1020 hold CPAT values that equal the literals (e.g. 5100); they are comparison data, left as is. Pending after a clean run: yaml (TASK-T3 + TASK-1 notes), `EgyptTaskReference.md`, `context-egypt.md`, the kernel path references in other documents (still say v1.3), tick T3.

## 2026-10-04 - AdHoc rebuild v0.4 built and verified; housekeeping and reconciliation (no model change)
- **Task:** (1) resolves the entry "AdHoc rebuild v0.4 builder and verifier drafted (not built, not run)": v0.4 was built and verified on Windows (`recalc_and_check_adhoc_v0_4_report.txt`: ALL PASSED; PROTOTYPE 132/132 for FULL and NOPHASE, REBUILD 588/588 against the Python mirror). (2) Wrote `MethodologyNote_v0.4` (current method only), `ResultsComparison_Table2_v0.4` and `VersionNotes_AdHocRebuild.md` (change history, kept out of the methodology). (3) Added NORMS section 7 and an AGENTS.md rule: methodology and version notes are always separate documents. (4) Reconciled stale pointers in README, `instructions-egypt.yaml` (TASK-2a, 2b, D, EF, QUEUE statuses), `context-egypt.md`, `EgyptTaskReference.md` (current-status table at the top) and `TODO.md` (header, T4, Final steps plan). (5) Archived superseded AdHoc files to `egypt/archive/AdHocRebuild/` (workbooks, builders and verifiers v0.1-v0.2; methodology notes and comparisons v0.1-v0.3; make_carveout v0.1-v0.2; make_results_page v0.3-v0.6 and v1.0). Files still needed by live scripts stay in `supporting/AdHocRebuild/` (rebuild v0.3 workbook/builder for the example-results builder; make_results_page v0.7, v1.1, v1.2; make_carveout_v0_3).
- **Inputs:** CAVEATS 2026-10-04 decisions; v0.4 workbook and report; v1.3 kernel and final documents (read-only checks).
- **Reconciliation findings:** (a) The final numbers agree across `EGYPT_CarveOut_Table2_v1.3.md`, the Table 2 CBAM-calc workbook and the kernel `CarveOut_Table2` stored snapshot (K, P, Q, J, N, T checked). (b) The two 2026-10-04 decisions are in the rebuild only, not the final Table 2: final 3B still has the block-only rebate (P 0.64, K -19.1), and final row O is the intensity-only measure, so the FULL decision affects the kernel/rebuild memo only. (c) The CBAM-calc workbook's "Full implementation" (CBF = 1) is what the kernel calls NOPHASE. (d) The v1.3 kernel's Settings title and last log row say v1.0 (no v1.1-v1.3 rows). (e) `EGYPT_FinalCaveats_v1.3` still names rebuild v0.3; the Methodology docx carries before/after material. (f) The CBAM-calc workbook's Table 2 rows are hard-coded.
- **Caveats:** no model or deliverable changed. `.docx` versions of the new notes were produced with `md_to_docx.py` and not opened in Word. The final-steps plan in `TODO.md` lists the decisions needed (D1-D4) before the single confirmation workbook is built.

## 2026-10-04 - Decisions applied everywhere: 3B rebate to all covered industry, one CBAM convention, single confirmation workbook (final set v1.5; builders drafted, not run)
- **Task:** (D1) apply the 3B rebate decision to the final Table 2 and every document; (D2) apply one CBAM convention consistently; (D3) put the single confirmation workbook inside the kernel; (D4) keep EG3 as run for 3A-3C.
- **Method for 3B (new; review it):** `make_carveout_v0_4.py`. In non-block covered industry the output share (1 - s_int = 2/3) of CPAT's response (fuel + proportional IPPU, NB = dInd + dIPPU - r x block) is removed from K (+7.2 Mt); its fuel part enters the deaths adjustment; the rebate tau x (kappa x IND1 - block post-policy fuel) = USD 0.33bn is deducted from revenue, with no feedback onto CPAT receipts. Result, 3B: K -19.1 to -12.0 Mt, P 0.6 to 0.3 $bn, Q 491 to 330; all other scenarios and columns unchanged. 3B net revenue is 0.3, not 0, because EG3 prices only 54% of industry.
- **CBAM convention (D2, interpretation):** Table 2 row O stays the CBAM-product intensity change (EU price and phase-in cancel). The deduction-based memo obligation is on FULL (2030 phase-in, CBAM factor 0.485) with NOPHASE (no phase-in, factor 1) as the other memo, under those names in the kernel, rebuild and documents. The old workbook label "Full implementation" (CBF = 1) was NOPHASE.
- **Outputs:** `egypt/final/` v1.5 documents: `EGYPT_Methodology_v1.5.docx` (3B scope, new section 4.7, history removed), `EGYPT_CarveOut_Table2_v1.5`, `EGYPT_FinalCaveats_v1.5`, `EGYPT_CBAM_ObligationNote_v1.5_NeedsCarolynConfirmation`, `EgyptResultsInitial_UpdatedResults_v1.5_tracked.docx`, `EGYPT_VersionNotes_v1.5` (new, change history), md sources in `md_sources/`. Builders/scripts in `egypt/supporting/AdHocRebuild/`: `make_carveout_v0_4.py` (+ `carveout_v1_5_results.json`), `make_results_page_v1_5.py`, `edit_methodology_v1_5.py`, `check_final_documents.py`. Kernel builders (`cpat_excel_new/standalone_working_version/`): `build_v1_4.py` (T3) then `build_v1_5.py` (3B rule, `Manual inputs` E112, `CarveOut_Table2` section G, new sheet `Table2_Final`). The v1.3 document set and the CBAM-calc workbook (with its script and data) are in `egypt/archive/`.
- **Verification:** `check_final_documents.py`: every Table 2 figure printed in the carve-out note and the results text equals the independent Python mirror (15 rows, 0 failures). The kernel builders have NOT been run (Excel COM, Windows). `build_v1_5.py` refuses to save unless the live kernel equals the mirror, the stored snapshot, the live row O and the printed figures, and the regression over all other cells is 0.
- **Caveats:** (1) `egypt/final/` documents say kernel v1.5, which exists only after `build_v1_5.py` has run (until then `final/` holds kernel v1.3). (2) The new `.docx` files were generated with pandoc / python-docx and not opened in Word. (3) The kernel's prototype composition (`Table2_Industry`, `Rebate_Industry`) still rebates the CBAM block only (reference). (4) The row O definition is still pending Carolyn's confirmation. (5) The 3B non-block rule is an application of CPAT's usage/efficiency split, not a CPAT re-run. (6) Open as before: true full-coverage EG3 run, EF VERIFY list, fp routing.

## 2026-10-04 - `egypt-final/` simplified hand-over folder
- **Task:** new top-level folder `egypt-final/` with a plain, version-free set of the final Egypt deliverables: `1_Summary.docx` (2 pages, plain language), `2_Results_Table.docx`, `3_Methodology.docx` (simplified rewrite), `Egypt_Model.xlsx`, `Egypt_CBAM_Obligations.xlsx`.
- **Inputs:** `egypt/final/` (kernel v1.3, Table 2 carve-out, CBAM calc workbook, methodology v1.3, caveats).
- **Outputs:** the five files above. Nothing in `egypt/` was changed or moved.
- **Caveats:** the two workbooks are byte copies renamed; internal sheet text (ReadMe/Settings version log, builder names) still carries version labels, because Linux cannot recalculate or rebuild them. The three Word files are new text; numbers were copied from `EGYPT_CarveOut_Table2_v1.3.md` and the methodology v1.3 and not independently recomputed. The CBAM-bill workbook is still pending Carolyn's confirmation (stated in the summary). Page counts checked via LibreOffice (summary 2 pages).

## 2026-10-04 - Kernel v1.4 (T3) and v1.5 (3B decision, Table2_Final) built and verified on Windows
- **Task:** resolves the entries "T3 builder drafted" and the v1.5 caveat that the builders had not been run. `build_v1_4.py` and `build_v1_5.py` were run by the user (commits "Minor changes", "Rebuilt to v1.5").
- **Checked after the pull (read-only, openpyxl on the saved v1.5):** sheets include `Table2_Final`; Settings log has rows v1.4 and v1.5; `Manual inputs` E112 = 1 and rows 101-108 hold the T3 values; Check section v1.4 literal count = 0 and v1.5 printed-vs-workbook mismatches = 0 (`Table2_Final` D42 = 0). Final Table 2 stored values: 3B K -11.976, P 0.315, Q 329.8, J 13.921 (others unchanged from v1.3: 1A K -37.39, 2A -33.08, 2B -34.83, 3A -22.77, 3C -33.04). Section D memo: FULL -44.2/-22.9/-23.0/-44.2/-5.0/-53.1, NOPHASE -24.0/-12.2/-12.3/-24.0/-5.1/-36.2.
- **Outputs:** `egypt/final/CPAT_Industry_Kernel_Egypt_v1.5.xlsx` (= `cpat_excel_new/standalone_working_version/` copy); v1.4 in `Old/`; v1.3 in `Old/` and `egypt/archive/`.
- **Caveats:** the builders' own printed regression output was not available to me, so the "regression 0" gate is inferred from the fact that they saved (they abort otherwise). 3B coverage J is 13.921 in the kernel (live kappa) against 13.916 in the documents (rounded to 14); no printed figure changes. Still open: the `egypt-final/` hand-over folder (built from v1.3, shows the old 3B numbers), Word check of the new .docx files, Carolyn's confirmation of row O, EG3 full-coverage run, EF VERIFY list.

## 2026-10-04 - `egypt-final/` updated to v1.5 (3B rebate to all covered industry)
- **Task:** brought the hand-over folder in line with the final set v1.5. `update_egypt_final_v1_5.py` (in `egypt/supporting/AdHocRebuild/`) edited the three Word files in place and replaced the model workbook.
- **Changes:** 3B figures in all three documents (K -19.1 to -12.0 Mt, -3.2 to -2.0% of GHG, revenue 0.6 to 0.3 USD bn, deaths 491 to 330); 3B described as free allowances for all covered industry; new row in the "how the cut is built" table (+7.2 Mt) and a note; a 3B paragraph in the methodology; pointers to the CBAM calculation now say `Egypt_Model.xlsx`, sheet `Table2_Final`. `Egypt_Model.xlsx` = kernel v1.5 (byte copy of `egypt/final/`). `Egypt_CBAM_Obligations.xlsx` removed (retired; its content is in `Table2_Final`).
- **Caveats:** the Word files were edited with python-docx and not opened in Word; the summary was 2 pages before and gained one sentence (page count not re-checked, LibreOffice does not open files in this environment). The 3B column of "how the cut is built" shows rounding (-21.5 + 6.1 - 3.7 + 7.2 = -11.9 against -12.0). The workbook still carries version labels internally (Settings log).

## 2026-10-04 - `egypt-final/simple/Egypt_Simple.xlsx` (plain-words workbook)
- **Task:** a very small workbook for a non-specialist reader: sheet "How it works" (five plain steps), "Results" (six scenarios; final cut built live from four visible lines; CBAM bill as before/after index), "Try it" (one example product, yellow cells = the seven material parameters, green = worked out), and a hidden sheet "Other parameters" holding everything else. Builder: `egypt/supporting/AdHocRebuild/build_simple_workbook.py` (openpyxl, formulas).
- **Inputs:** `carveout_v1_5_results.json` (same source as the v1.5 documents).
- **Caveats:** simplified picture only; the Results numbers are typed from the full model (kernel v1.5), only the sums and the "Try it" example are live. The "Try it" example (cement clinker, USD 20) uses rounded inputs (process saving 6.5%, fuel saving 4.9%) and is illustrative, not a reproduction of the kernel. Not recalculated in Excel here (checked the formulas in Python: final cut equals K for all six scenarios; example gives -5.9% emissions per tonne, -12.4% total).

## 2026-10-04 - Output elasticity: evidence note and recommendation (no model change)
- **Task:** looked for a better basis for the uniform output elasticity -0.5 (`Manual inputs` rows 66-73). Wrote `egypt/supporting/OutputElasticity_Note_v0.1.md/.docx`.
- **Findings:** only cement matters (3.6 of the 3.8 Mt output-channel cut in 1A, because the carbon cost is 15% of the clinker price). Evidence for cement: demand elasticity about -0.02 to -0.16, pass-through 20-40%; steel demand -0.2 to -0.3, pass-through very wide (about 0.5 central); EU ETS firm studies find no detectable output fall. Recommended: cement -0.10 (range -0.03 to -0.30), steel and fertilisers -0.40, aluminium -0.50 (unchanged). Effect (approximate, outside the kernel): 1A cut -37.4 to -34.5 Mt, 3A -22.8 to -19.9, 3C -33.0 to -30.1; 3B unchanged.
- **Caveats:** values are judgements from sourced components (Low-Medium confidence), from web search summaries; the papers could not be opened (egress blocked). Nothing in the kernel or deliverables changed. Decision pending: apply cement -0.10 (kernel rebuild, new carve-out numbers, document updates) or keep -0.5 as a documented placeholder.

## 2026-10-08 - Final set v1.6: product-specific output elasticities (documents on mirror numbers; kernel builder and rebuild v0.5 drafted, not run)
- **Task:** apply cement -0.10 and the full recommended set (steel and fertilisers -0.40, aluminium -0.50; was -0.5 for all) per `OutputElasticity_Note_v0.1`, everywhere. Decisions: full set; one pass with approximate numbers; rebuild v0.5 included; set labelled v1.6.
- **New numbers (mirror, `make_carveout_v0_5.py`), 2030 K Mt old to new:** 1A -37.4 to -34.7; 2A -33.1 to -31.9; 2B -34.8 to -33.7; 3A -22.8 to -20.0; 3B -12.0 unchanged; 3C -33.0 to -30.8. Deaths: 1A 1,501 to 1,440; 2A 1,456 to 1,431; 2B 1,517 to 1,492; 3A 552 to 509; 3C 646 to 607. Revenue 3A 1.5 to 1.6; others within rounding. N 1A/3A -5.8 to -5.7, 2B -2.3 to -2.2, 3C -23.1 to -23.0; T (block emissions) 1A/3A -11.6 to -7.2, 2A -4.7 to -2.9, 2B -4.8 to -3.0, 3C -27.8 to -24.3. O (row O), J, M unchanged.
- **Mirror method:** emissions scale linearly with output, so the stored kernel block rows (policy rows 773-788, product outputs 675-682) and the block process revenue / 3C fund items are rescaled per product by (1 + dp)^eps_new / (1 + dp)^-0.5 and the carve-out arithmetic is rerun. Approximations: the 3C fund shadow price is not re-solved; kernel N is shifted by the change in the output weights; O is unchanged.
- **Outputs:** documents v1.6 in `egypt/final/` (methodology, carve-out note, final caveats, CBAM obligation note, tracked results text, version notes; v1.5 set in `egypt/archive/`), `egypt-final/` (three Word files; `simple/Egypt_Simple.xlsx` with cement default -0.1), scripts `make_carveout_v0_5.py`, `make_documents_v1_6.py`, `edit_methodology_v1_6.py`, `update_egypt_final_v1_6.py`, `build_v1_6.py` (kernel), `build_adhoc_rebuild_v0_5.py` and `recalc_and_check_adhoc_v0_5.py` (rebuild).
- **Verification:** `check_final_documents.py` (printed figures = mirror: 15 rows, 0 failures). Kernel v1.6 and rebuild v0.5 have NOT been built or run (Excel COM). `egypt/final/` still holds kernel v1.5 and `egypt-final/Egypt_Model.xlsx` is still v1.5 until `build_v1_6.py` runs; the documents state v1.6 numbers that come from the mirror.
- **Caveats:** the exact kernel numbers can differ from the mirror (largest expected for 3C, whose fund is re-solved in the kernel); the builder saves anyway, prints "DOCUMENTS MUST BE REGENERATED" if the printed figures differ, and the documents are rebuilt from `carveout_v1_6_results_live.json` (TODO, Final steps). Rebuild v0.5 notes and results are not written yet. Elasticity values are judgements (Low-Medium confidence); citations unverified against the papers.

## 2026-10-08 - Final set v1.6 finalised on the live kernel results
- **Task:** after `build_v1_6.py` and the rebuild v0.5 run (both on Windows, user), regenerated every document from `carveout_v1_6_results_live.json`, finalised the supporting notes and tidied the final folder.
- **Live kernel against mirror:** K, P, Q, T and O agree to rounding; the printed figures that changed are N for 1A/3A (-5.74 mirror to -5.79 live, still printed -5.8), 2B (-2.26, printed -2.3) and 3C (-23.10, printed -23.1), 3C deaths 606 (mirror 607), 3C K -30.82. Kernel `Table2_Final` mismatch count 0; `check_final_documents.py` 15 rows, 0 failures; kernel coverage J is 13.921 (live kappa) against 14 printed.
- **Outputs:** `egypt/final/` (v1.6): kernel, methodology, carve-out note, final caveats, CBAM obligation note, tracked results text, version notes (md sources in `md_sources/`); `egypt-final/` three Word files, `Egypt_Model.xlsx` (= kernel v1.6) and `simple/Egypt_Simple.xlsx` from the live numbers. Final caveats sections A-D now use the rebuild v0.5 and prototype v1.6 figures and no history wording (NORMS section 7). Rebuild notes: `MethodologyNote_v0.5`, `ResultsComparison_Table2_v0.5` (shortened: tables plus qualitative reasons; the old numeric decomposition of the prototype gap was dropped because it could not be recomputed outside Excel); v0.4 notes archived.
- **Caveats:** the new .docx files were generated with pandoc / python-docx and not opened in Word; row O still needs Carolyn's confirmation; evidence-note citations unchecked against the papers; the rebuild/prototype figures in the final caveats are as of the v0.5 / v1.6 runs.

## 2026-10-08 - TODO.md simplified
- **Task:** rewrote `TODO.md` to open work only: three checks for the user (Word check, Carolyn's confirmation of row O, citation check), seven modelling items (full-coverage EG3 run, EF VERIFY list, fp routing, block fuel-intensity channel, fuel-CO2 reconciliation, aligning the kernel's own composition with the 3B decision, items under review) and a short conventions list. Completed T1-T5, D1-D4 and the v1.4-v1.6 steps were removed (history is in this log). README and yaml pointers to TODO updated.
- **Caveats:** none.

## 2026-10-08 - Evidence check, EG3 full-coverage preparation, kernel-increment specs (no model change)
- **Task:** worked through the open modelling list as far as possible without Excel or CPAT.
- **Citations (`OutputElasticity_Note_v0.2`, replaces v0.1 in the archive):** checked by search summaries: EC / CE Delft-Oeko 2016 pass-through (cement 20-40%, steel 55-85%, correcting the earlier "about 0.5"), Ganapati-Shapiro-Walker 2020 (about 0.59 average), Colmer et al. 2025 (14-16% emission cut, no detectable output fall), cement demand (aggregate -0.02 to -0.04, median industry-level -0.10; the -0.16 attribution unconfirmed). Single-source: GTAP Armington values. Not found: a steel demand study and a numeric fertiliser demand elasticity. All fetches of the papers were blocked by the network policy. Methodology (parameter table) and `egypt-final` methodology text corrected for steel pass-through; recommended values unchanged.
- **Methodology wording:** Section 3 and the channel table said the proportional IPPU row "is removed"; now: removed for the CBAM block, kept for non-block IPPU. Typo fixed.
- **EG3 full coverage:** `EG3_FullCoverage_RunSpec_v0.1` (CPAT settings, acceptance check kappa about 1, export, regeneration, statements to rewrite); `cpat_run_constants.py` (reproduces every constant behind the final Table 2 from the csv); `make_carveout_v0_6.py` (mirror on a new csv; reproduces v1.6 on the current csv to 0.006); `build_v1_7.py` (kernel refresh of `CPAT_National`, stored snapshots and `Table2_Final`; drafted, not run). Indicative effect on the final Table 2 (1/kappa proxy): 3A -20.0 to about -29 Mt, 3B -12.0 to -15, 3C -30.8 to -40.
- **Kernel increments:** `KernelIncrements_Spec_v0.1`: fp routing about -0.3 Mt (optional), block fuel-intensity channel (none; double-counting risk with the carve-out; not recommended), 3B alignment of the kernel's own composition (reference sheet; label instead). No builders written for these three.
- **Caveats:** the indicative EG3 numbers are a proxy, not a CPAT result; `build_v1_7.py` has not been run and needs the new csv; the text of `Manual inputs` F66:F68 in the kernel still carries the old steel pass-through wording.

## 2026-10-05 - Kernel v1.7: CPAT_National linked to the new CPAT_Results sheet
- **Task:** in the user's `egypt/final/CPAT_Industry_Kernel_Egypt_v1.7.xlsx` (v1.6 plus a new raw CPAT export sheet `CPAT_Results`, runs EG1-EG5, 6,420 rows), replaced the stored values of `CPAT_National!L6:AE133` (2022-2041) with `=INDEX(CPAT_Results!$A:$AD,MATCH($H6&"_"&$C6,CPAT_Results!$F:$F,0),MATCH(L$5,CPAT_Results!$1:$1,0))` (key = CPAT code & "_" & run in `CPAT_Results` column F; year from row 5). Source column F and the section A band text updated. Done by Excel COM; `build_v1_7.py` was not used.
- **Verification:** all 128 keys found; 2,560 linked cells equal `CPAT_Results` exactly, no errors. 16 keys occur twice in `CPAT_Results` (co2.tot, rev.new.usd); the duplicates are identical, the first is used.
- **Caveats:** the new export changes the baselines as well as EG3 (about 2,060 of 2,560 cells differ from v1.6 in EG1, EG2, EG3 and EG4; e.g. `egy.air.mort` by about 31,000), so the new run is not an EG3-only replacement. EG5 is in `CPAT_Results` but not used. `CarveOut_Table2` (42 cells, from D22/D23: `Table2_Industry!G31/G32 & "|" & run` not found in `CPAT_National`) and `Table2_Final` (16 cells) show #N/A / #DIV/0!; these errors were already present in the v1.7 file before the links and are unchanged. Stored snapshots, documents and `egypt-final/` not refreshed.

## 2026-10-08 - Mitigation copy-paste prototype (v0.1)
- **Task:** first build of a copy-pasteable replacement for the CPAT mitigation module, price -> fuel use only, design 2 (scenario > year across; block > subsector > fuel down; per-block parameter sets on the left picked by scenario number). Goal, spec and decisions agreed in conversation (README in the folder).
- **Inputs:** legacy `CPAT 1.0pre_456_NoPropData.xlsb` (`Elasticities` rows 138-282 by income group; `Mitigation` rows 980-1009 EFs and unit conversions); `CPAT_PriceModule_Data_1.xlsx` `DomesticPrices` (221 countries, 2021-2024); kernel v1.6 `Data_Energy` (legacy Mitigation rows 583-602) and `Data_Macro` (legacy row 2446); CPAT public documentation (github.com/cpmodel/cpat_public, Mitigation 3.3.3 and 3.3.5); `cpat_coded` `ec.py` (subsector groupings, coal/gas price rule).
- **Outputs:** `cpat_excel_new/mitigation_copypaste/`: `CPAT_Mitigation_CopyPaste_v0.1.xlsx`, `build_v0_1.py`, `extract_data_v0_1.py`, `check_v0_1.py`, `check_report_v0.1.md`, `data/*.csv`, `README.md`. Checks (LibreOffice recalculation): 0 errors; one R1C1 formula per block across 128 rows x 2 groups; copied group = source group (diff 0) and copied group with carbon price 0 = baseline (diff 0); independent Python recomputation max relative diff 5e-15. Result: $20/t from 2027 lowers total fuel use by 10.7 % in 2027-2030 (Egypt's low, subsidised prices make the relative price rise large).
- **Caveats:** (1) The CPAT documentation tables (Figures 54, 57) differ from the legacy workbook: the workbook applies the developing-country adjustment to the usage elasticity (LIC/LMIC/UMIC = HIC - 0.1), not to the efficiency elasticity as the docs say, and several base values differ (e.g. residential gas income elasticity 0.3 vs 0.6). Legacy values used (user decision). (2) Simplifications vs CPAT: no per-capita GDP adjustment of income elasticities (user decision), no Covid factor, no shadow prices or additional efficiency gains, no self-generated renewables, jet fuel or road biofuel split, no electricity; pre-tax prices flat at the 2022 supply cost (growth 0 %, REVIEW); base tax = retail - supply cost (includes excise, subsidies and VAT) held constant. (3) Assumptions to review: EF sector for foo and srv = res, for transport subsectors = rod, for oen = ind; elasticity sector for foo and oen = ind (as `cpat_coded`); IIASA EFs without the Egypt inventory adjustment (0.93) for the tax base; carbon price taken as $/tCO2 in the same (2022 current) dollars as prices. (4) Base-year energy use, GDP growth and EFs are Egypt only; other countries return #N/A. (5) Built with openpyxl, not Excel COM (NORMS/TODO convention); verified in LibreOffice, not yet opened in Excel. (6) Years start after the key and parameter columns, not in column K (NORMS 1.1 deviation, noted in the ReadMe). (7) Not yet compared with legacy CPAT results (bucket 4).

## 2026-10-08 - Mitigation copy-paste prototype (v0.2)
- **Task:** revise v0.1 after review: lookups separated from formulas (new `Inputs` sheet, one row per fuel|subsector, elasticities expanded to that taxonomy); four parameter columns D:G, hidden by default, between Sector and Description; parameters global (scenario-specific parameter sets and scenario numbers removed; scenarios differ only in the carbon price); base year 2022 in column L as in legacy Mitigation; 2023-2034 plain formulas and 2035 named LAMBDAs (`PRETAX`, `TAX`, `POSTTAX`, `FUELUSE`, orange) that can be dragged back over the row.
- **Inputs:** v0.1 (`Old/`), `data/*.csv` (unchanged), `cpat_excel_new/distribution/LESSONS_LEARNED.md` (LAMBDA encoding), kernel v0.1b LAMBDA (encoding pattern).
- **Outputs:** `cpat_excel_new/mitigation_copypaste/CPAT_Mitigation_CopyPaste_v0.2.xlsx`, `build_v0_2.py`, `check_v0_2.py`, `check_report_v0.2.md`, README; v0.1 files moved to `Old/`; Settings version log row added. Checks: one R1C1 formula per block for base year, plain years and the LAMBDA year; LAMBDA encoding (`_xlfn.LAMBDA`, every parameter `_xlpm.`); LAMBDA bodies expanded and plain formula dragged into 2035 both match the independent Python recomputation (5e-15); scenario-copy test diff 0; regression vs v0.1 diff 0 on all 14,336 block cells.
- **Caveats:** (1) LibreOffice 24.2 cannot evaluate LAMBDA, so the 2035 column was verified by expansion, not by Excel; LESSONS_LEARNED warns that hand-encoded LAMBDA names can be stripped by some Excel builds - if Excel shows a repair prompt or #NAME? in 2035, drag 2034 forward (or re-create the names in Name Manager from the ReadMe). Not yet opened in Excel. (2) Further NORMS 1.1 deviation: Description..Output Code move to H:K and Input Code, Note and Helper are dropped. (3) All v0.1 caveats on simplifications and assumptions still apply.

## 2026-10-08 - Mitigation copy-paste prototype (v0.3)
- **Task:** full CPAT output codes with country and scenario, repeated for each scenario: column A now holds the variable code (`gdp.pos.pct`, `cptraj`, `sp`, `tax`, `atp`, `ener`, `ener.chk`, `ener.ref`, `ener.pct`); Description, Unit and Source (H:J) are looked up from a new `Variables` sheet; each scenario group starts with a code column (K for scenario 1, base year 2022 still in L) that builds `country.mit.<variable>[.<subsector>.<fuel>][.<suffix>].<scenario>` (e.g. `egy.mit.ener.rod.gso.e.1`, as legacy row 5430). The scenario number sits at the top of the code column; scenario 1 is typed and every later group is previous + 1, so a pasted group numbers itself.
- **Inputs:** v0.2 (`Old/`); legacy Mitigation output codes (rows 2498 `sp...a.N`, 5416 `atp...e.N`, 5430 `ener...e.N`, 2446 `gdp.pos.pct.N`; `cptraj.N` from the CPAT outputs csv).
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.3.xlsx`, `build_v0_3.py`, `check_v0_3.py`, `check_report_v0.3.md`, README; v0.2 files moved to `Old/`; Settings version log row. Checks: one formula per label column and per code column (542 rows), expected codes for 8 sample cells, pasted groups number themselves (3, 4), regression vs v0.2 diff 0, all v0.2 checks still pass.
- **Caveats:** (1) Codes follow the legacy pattern but use the subsector where legacy prices use the price sector (`egy.mit.sp.ind.coa.a.1`) and legacy `atp` is in $/liter for liquids while ours is $/GJ; `tax` and the totals/check codes (`ener.chk`, `ener.ref`, `ener.pct`, `...rod.all...`) are new, not legacy. (2) Labels for total rows rely on the `all` rows added to Mapping (rows 21 and 32). (3) All v0.1/v0.2 caveats still apply; still not opened in Excel.

## 2026-10-08 - Mitigation copy-paste prototype (v0.4)
- **Task:** roll up the sections: a row outline groups each block's rows under its band row (+/- on the band; summary row above), blocks collapsed by default, totals grouped but open.
- **Inputs:** v0.3 (`Old/`).
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.4.xlsx`, `build_v0_4.py`, `check_v0_4.py`, `check_report_v0.4.md`, README; v0.3 files moved to `Old/`; Settings version log row. Checks: outline levels and collapsed state per block, all v0.3 checks, regression vs v0.3 diff 0.
- **Caveats:** (1) openpyxl drops `sheetFormatPr/@outlineLevelCol`; the builder adds it after saving (`fix_outline_levels`) so the hidden D:G column group gets its button - v0.2/v0.3 lacked it. (2) Copying a scenario group or dragging across includes collapsed (hidden) rows, which is intended; a filtered view would not. (3) Still not opened in Excel.

## 2026-10-08 - Mitigation copy-paste prototype (v0.5)
- **Task:** white text in the green bands; a white summary line directly under each band, visible when the block is rolled up (the +/- now sits on that line): fuel use = total fuel use (`egy.mit.ener.all.all.e.N`); tax = policy carbon price (new variable `cptraj.ref`, repeating row 7); pre-tax and after-tax price = no total (prices are not additive, line left empty with a note). The totals section's own grand-total row is dropped (it duplicated the fuel-use summary line); its check row now tests totals by subsector and by fuel against the summary line.
- **Inputs:** v0.4 (`Old/`).
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.5.xlsx`, `build_v0_5.py`, `check_v0_5.py`, `check_report_v0.5.md`, README; v0.4 files moved to `Old/`; Settings version log. Checks: all earlier checks, white bold band text, summary lines equal the carbon price and the sum of the 128 fuel-use rows, outline under the summary lines, regression vs v0.4 diff 0.
- **Caveats:** (1) Every block is one row taller (band + summary line + 128 rows + blank), so all row numbers below row 9 shift by 1-4 against v0.4. (2) Bug found and fixed: the builders v0.3 and v0.4 wrote earlier version-log rows with the current version number (shipped v0.3 shows the v0.2 row as 0.3; v0.4 shows the v0.2 and v0.3 rows as 0.4); v0.5 hard-codes each row's version. Old workbooks are left as shipped. (3) `cptraj.ref` and the totals-section codes are new, not legacy. (4) Still not opened in Excel.

## 2026-10-08 - Mitigation copy-paste prototype (v0.6, bucket 6a: MTInputs and carbon price)
- **Task:** policy inputs per scenario. New sheet `MTInputs`: the template (`templates/MTInputs_template.xlsx`, legacy row structure, names, units, `NameOfParameter`) copied cell by cell into A:H, plus one *Used for calculation* column per scenario from column J (number row 5, auto-numbered; name row 6; Scenario ID and name rows 10/12 link to them). `Mitigation` gets a scenario-number row (5), the scenario name from MTInputs (row 6), a carbon-tax band with lookup rows for `CPIntro`, `CPLevelStart`, `CPLevelTarget`, `CPOutro`, `ExtendCarbonPriceBeyondOutro` (rows 9-13, rolled up) and the carbon price trajectory as a formula (row 14, as legacy rows 2248/2251), replacing the typed carbon prices. Decisions: MTInputs columns A:H unchanged; scenario columns from J; ETS left out for now; `MCovOen` = FALSE.
- **Inputs:** v0.5 (`Old/`); `templates/MTInputs_template.xlsx`; legacy Mitigation rows 2188-2252 and 8579-8582 (carbon tax settings and trajectory).
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.6.xlsx`, `build_v0_6.py`, `check_v0_6.py`, `check_report_v0.6.md`, README; v0.5 files moved to `Old/`. Checks: MTInputs A:H = template (0 cells differ, D:E hidden); a scenario with the legacy Egypt settings (2026, 0 -> 50 by 2030, Linear*) reproduces legacy row 8582 (0, 12.5 ... 50, 62.5 ... 112.5); copied MTInputs columns and Mitigation groups number and name themselves; scenario 4 with zero carbon price in MTInputs = scenario 1; regression vs v0.5 diff 0 (blocks and carbon price); all earlier checks pass.
- **Caveats:** (1) Scenario columns start from the template's *Used for calculation* values (an Egypt run: carbon tax 0 -> 50 by 2030 from 2026), overridden for our scenarios: scenario 1 carbon price 0, scenario 2 $20 from 2027 (`CPOutro` 2030, start = target), `MCovOen` FALSE in both; overrides are in red. (2) Only the carbon-tax parameters are used so far; coverage still comes from `Mapping` (oen 0) until bucket 6b switches to the `MCov*` switches. (3) Nominal vs real (`NomorReal`), the exponential option (`CTaxTrajectoryType`) and CPI conversion (legacy row 2249) are not implemented: the trajectory is the real, linear path. (4) Lookup ranges cover MTInputs scenario columns J:AZ (41 scenarios). (5) Rows below 8 shift by 7 against v0.5. (6) Still not opened in Excel.

## 2026-10-08 - Mitigation copy-paste prototype (v0.7, bucket 6b: policy wedges, legacy section layout)
- **Task:** (1) legacy CPAT visual structure: numbered sections 1. Policies, 3. Power sector (placeholder), 5. Transport, 6. Buildings, 7. Industrial, 8. Other energy use, 11. Results; within a sector section subsector > variable > fuel (as legacy rows 5538-5560), with a total fuel-use line per sector and per subsector. (2) Policy wedges per fuel x subsector: `ctxnew` (carbon price x EF x MTInputs fuel and sector coverage `MCov*`), `ntx` (fuel price reform path / GJ per unit), `nce` (total new policy), `tax` = base tax + `nce`, and the feebate shadow price `shp` (calculated, not yet used in fuel use). Section 1 reads all inputs from MTInputs by row and scenario number and computes the carbon price, fuel price reform and feebate paths. (3) The carbon-price line repeated under the tax block (v0.5-v0.6, `cptraj.ref`) is removed. Decisions: no ETS yet; `MCovOen` FALSE.
- **Inputs:** v0.6 (`Old/`); `templates/MTInputs_template.xlsx` (rows 18-47, 53-82, 140-169, 265); legacy Mitigation section layout (rows 2174-2447, 4302, 5538-5560, 5665-5690, 6063-6087) and codes (`ctxnew` row 2510, `ntx` 2512, `nce` 2514, `shp` 2377).
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.7.xlsx`, `build_v0_7.py`, `check_v0_7.py`, `check_report_v0.7.md`, README; v0.6 files moved to `Old/`; Settings version log. Checks: one formula per variable (base, plain years, 2035) across 128 rows x 2 scenarios, per D:G column, Section-1 inputs and paths, labels and codes; values vs independent Python recomputation (5e-15) for scenarios 1-2 and test scenarios 5 (legacy carbon tax = legacy row 8582), 6 (fuel price reform), 7 (feebates: shadow price non-zero, fuel use unchanged); copied scenarios = source; regression vs v0.6 on 1,094 shared output codes (max diff 1e-15); outline levels; MTInputs = template.
- **Caveats:** (1) Fuel price reform and feebate paths are assumed linear from start to target year and flat afterwards; the legacy formulas could not be read (values are zero in the Egypt run). Fuel price reform increases are in the MTInputs price units ($/GJ, $/liter, $/bbl) and converted with the GJ-per-unit factors; electricity rows are read but unused. (2) Feebate sector for food & forestry and services = residential, other energy use = industry (assumptions). (3) Biomass has no carbon-tax switch in MTInputs; its coverage row is a typed FALSE (biomass EF is 0 anyway). (4) `ctxnew`, `ntx` and `shp` pick a Section-1 row by position (`INDEX(range, position)`); positions come from Inputs (data step). (5) The shadow price is not yet in the fuel-use equation; ETS, subsidy phase-out, price liberalisation, methane fee and CBAM process emissions are not implemented. (6) Power sector is a placeholder section. (7) Row numbers change completely against v0.6 (codes are stable). (8) Still not opened in Excel.

## 2026-10-08 - Mitigation copy-paste prototype (v0.8: shadow price on the efficiency margin)
- **Task:** feed the feebate shadow price into fuel use as legacy CPAT: the efficiency term uses `((atp + shp)/(atp_prev + shp_prev))^eps_F` (all ^(1+eps_U)), the usage term `atp` only (`cpat_coded` `ec.py`, legacy rows 2374-2419). Section 1 gains shadow prices by sector ($/tCO2: power, transport, residential, industry = feebate path; legacy rows 2345-2349) and the share of the shadow price impacting efficiency by subsector (`ssc` = feebate coverage x adjustment 1.0; legacy rows 2403-2419; `cpat_coded` default for feebates). `shp` = sector shadow price x EF x `ssc`; `FUELUSE` LAMBDA takes the shadow prices (10 arguments). User decisions: one default for all paths (linear continuation after the target year; the carbon price keeps its MTInputs switch); feebate grouping as the legacy main grouping (buildings = res, foo, srv -> residential rate); biomass carbon-tax coverage TRUE like the other fuels (EF 0); positional INDEX picks accepted but kept rare.
- **Inputs:** v0.7 (`Old/`); `cpat_coded/cpat_model/components/energy_consumption/ec.py` and `policies/shadow_prices.py`; legacy Mitigation rows 2345-2420.
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.8.xlsx`, `build_v0_8.py`, `check_v0_8.py`, `check_report_v0.8.md`, README; v0.7 files moved to `Old/`; the MTInputs template path in `Old/build_v0_6.py`, `Old/build_v0_7.py` and `build_v0_8.py` now searches upward (archived builders run from `Old/`; output unchanged). Checks: all v0.7 checks; shadow-price paths uniform; feebate test: fuel use matches Python (5e-15) and falls only in subsectors with feebate coverage and a non-zero rate (road -11.1 %, iron & steel -12.3 % in 2035 at 10 -> 50 / 5 -> 25 USD/tCO2 from 2027 to 2030, continuing linearly); regression vs v0.7 on 2,322 shared codes diff 0 (intended change: biomass coverage switch).
- **Caveats:** (1) The scenario columns copy the template's feebate coverage switches (TRUE for most industrial subsectors and other energy use); with zero feebate rates they have no effect. (2) Other energy use takes the industry feebate rate (legacy groups it under 'oth'). (3) Elasticity grouping unchanged: food & forestry uses industry elasticities (as `cpat_coded` EC_SECTORS) although it sits in the buildings section - open question for the user. (4) Regulations (MTInputs rows 247-250 adjustments) and ETS are not yet shadow-price sources; additional policy-induced efficiency gains (rows 242-245) not yet in the fuel-use equation. (5) Still not opened in Excel.

## 2026-10-08 - Mitigation copy-paste prototype (v0.9: domestic price projection in real terms)
- **Task:** Implement in one step the legacy domestic price projection, and write out the real/nominal assumptions explicitly. The user decided not to track values after 2030, and to keep one index only if it matches legacy up to 2030.
- **Changes:**
  - New section 2 (12 price fuels) with:
    - `gp` (international prices: MTInputs source and High/Low adjustment, country gas market);
    - `sp` (data to 2024, then fixsp + floating part x gp ratio);
    - `txo` (data, then the legacy pass-through rule);
    - `rpb` = (sp + txo)(1 + VAT).
  - Top rows `infl` (US CPI index) and `defl` (US GDP deflator index), both = 1 in ResultsYear 2026.
  - `NomorReal` in section 1: a nominal carbon price is multiplied by `infl`.
  - Per subsector: `atp` = max(rpb + nce(1 + VAT), 0.01); `sp` and `tax` removed.
  - New sheets `Inputs_prices` (data step), `Prices_int` and `PriceAssump`; Settings C8-C12.
  - LAMBDAs `SUPPLYCOST` and `OTHERTAX` replace `PRETAX` and `TAX`; `POSTTAX` now takes the VAT rate.
- **Index decision:** The CPI and deflator indices differ before 2030 by up to 1.2% (2022: 1.137 vs 1.124; 2030: 0.918 vs 0.929). One index would not reproduce legacy, so both are kept: CPI for domestic data and nominal inputs, deflator for international prices (as legacy).
- **Inputs:** v0.8 (`Old/`); `cpat_coded/cpat_model/components/prices/prices.py`, `domestic_prices.py` and `international_prices.py`; legacy Mitigation rows 452/455/456/467/725 and 633-695; legacy `Prices_int` regional assumptions; `extract_data_v0_2.py` -> `data/prices_int.csv`, `weo_us.csv`, `price_assumptions.csv`.
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.9.xlsx`, `build_v0_9.py`, `check_v0_9.py`, `check_report_v0.9.md`, `extract_data_v0_2.py`, `PriceProjection_Method_v0.2.md` (assumptions A1-A10), README. v0.8 files, method note v0.1 and `extract_data_v0_1.py` moved to `Old/`.
- **Checks:**
  - Every new block has one R1C1 formula, with history and projection in the same formula.
  - gp, sp, txo, rpb and all subsector variables match an independent Python recomputation from the CSVs (max relative diff 1e-14). This covers scenarios 1-2; price source IMF-WB* with High adjustment and a nominal carbon price; and global price controls None and Manual.
  - Pass-through 0 keeps rpb at its 2024 value.
  - Regression vs v0.8: 1,286 shared non-price codes diff 0. atp, ener and sp change by intent.
- **Results (Egypt):**
  - Baseline fuel use: 49.7 Mtoe (2022), 62.5 (2023), 68.9 (2024), 78.8 (2027), 105.3 (2030). In v0.8 it was 55.4 (2027) and 60.7 (2030).
  - $20/t from 2027: -17.5% (2027) and -30.1% (2030) against the baseline. In v0.8 it was -10.7% in 2027.
- **Caveats:**
  1. Legacy rules are copied as they are, including two quirks:
     - (a) pass-through 0.5 or 0.8 makes `txo` jump in the first projected year;
     - (b) other oil products have pass-through 1 hardcoded but keep their negative fixed tax. Their retail price falls with the oil price: 3.79 $/GJ (2024), 0.23 (2030), 0.01 floor from 2031. Baseline oop fuel use therefore rises 10x by 2030 (3.0 to 29.2 Mtoe) and explodes after it. This drives most of the higher baseline and the large carbon-price effects (other manufacturing, fuel transformation, navigation). Proposed fix pending user decision.
  2. Real prices fall sharply from 2022 to 2024 in the data (devaluation; gasoline 15.1 to 8.4 $/GJ real), so modelled fuel use jumps +26% in 2023 and +10% in 2024. Legacy also models from the 2022 base year.
  3. Historical oil-product retail prices follow legacy (sp + txo; Egypt VAT rate on oil products is blank = 0), so they are 14% below the dataset `rp`, which includes VAT.
  4. Margins and production costs (2022 nominal) are converted with `infl(2022)`; `cpat_coded` applies no conversion. To check in bucket 5.
  5. ResultsYear and price controls are global (scenario 1 column). Source, adjustment and NomorReal are per scenario.
  6. Scenario 1 takes the template's *Used for calculation* source `IMF` (the legacy default is `IMF-WB*`).
  7. The legacy NoPropData price forecasts are #N/A, and data vintages differ, so the projected prices cannot yet be validated against legacy numbers.
  8. Still not opened in Excel.

## 2026-10-08 - Mitigation copy-paste prototype (v0.10: other oil products, VAT assumption)
- **Task (user decisions):** (1) Treat other oil products like the other oil products. (2) Make an assumption for Egypt's VAT rate. (3) Mark changed data clearly in a bright colour.
- **Changes:**
  - Other oil products: pass-through and margin now come from the IMF dataset (Egypt: 0 and 7.95 $/bbl) instead of the legacy hardcodes 1 and 0.
  - VAT rate: the dataset rate where filled; otherwise the general rate `VAT_WEO` (Egypt 14%) for residential coal and gas and the all-sector oil products, and 0 for power, industry (credited to firms) and biomass (informal).
  - The changed cells in `Inputs_prices` (VAT consumer flag, VAT assumption, oop pass-through and margin) are bright yellow with red bold text; the ReadMe and the method note (A10, A11) say so.
- **Evidence for the VAT assumption:** the Egypt dataset's own retail prices imply it: rp = (sp + txo) x 1.14 for residential gas, gasoline, diesel, LPG, kerosene and other oil products, and rp = sp + txo for power and industry.
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.10.xlsx`, `build_v0_10.py`, `check_v0_10.py`, `check_report_v0.10.md`, `PriceProjection_Method_v0.3.md`, README. v0.9 files and method note v0.2 moved to `Old/`.
- **Checks:** PASS. The price chain and fuel use match the independent Python recomputation (5e-15), and the bright marking is checked. Regression vs v0.9: 1,300 unaffected codes diff 0; 646 price and fuel-use codes change by intent.
- **Results (Egypt):**
  - Other oil products: retail price held at 4.32 $/GJ from 2024 (v0.9: 0.23 in 2030). Their fuel use is 4.9 Mtoe in 2030 (v0.9: 29.2).
  - Gasoline retail price 9.58 $/GJ (v0.9: 8.40, without VAT).
  - Baseline total 74.1 Mtoe (2027) and 81.0 (2030).
  - $20/t from 2027: -15.1% in 2027 and 2030.
- **Caveats:**
  1. The VAT rule is an assumption. It reproduces Egypt's data but is not a statement of Egyptian tax law: petroleum products may be under a schedule tax rather than standard VAT. The fiscal split between VAT and excise does not affect fuel use, but it will matter for revenue.
  2. The VAT assumption applies to any country with blank dataset rates. Check it per country.
  3. The legacy `txo` jump at pass-through 0.5 or 0.8 is still copied.

## 2026-10-09 - Mitigation copy-paste prototype (v0.11: ETS, shadow prices, existing taxes and subsidies, 2040, LAMBDA right column)
- **Task (user decisions):**
  1. Food & forestry takes the buildings elasticities. Services is used, because legacy groups food & forestry with services as Commercial.
  2. Differences with legacy go in a new tab.
  3. Existing taxes and subsidies, (a) ready for revenue and (b) as an input for a later per-fuel policy.
  4. ETS and feebates with sectoral shadow prices; the ETS is as effective as a carbon tax.
  5. Complete the prices section and prepare revenues.
  6. Separate column blocks for history and projection. The LAMBDA (which may hold the IF) goes in the right column of every calculated row.
  7. Extend to 2040 with LAMBDAs for the baseline and the policy scenario.
- **Changes:**
  - Horizon 2022-2040.
  - Section 1:
    - New ETS inputs: MTInputs rows 84-92 and coverage rows 99-115.
    - Effective ETS coverage `etsc`.
    - ETS permit price `ets.p` = the carbon price path once the ETS applies.
    - Auctioned share `ets.a`: linear, constant after the target year.
    - Regulation shadow prices `reg`: placeholder 0.
    - `shps` = feebates + regulations.
  - Subsectors:
    - New variable `ets` = permit price x EF x effective coverage.
    - `ctxnew` x (1 - ETS coverage), so there is no double pricing.
    - `nce` = ctxnew + ets + ntx.
  - Section 2 adds, per price fuel:
    - `vat`: VAT payment before new policies;
    - `etx` = max(txo, 0): existing tax;
    - `esub` = max(-txo, 0): existing consumer subsidy;
    - `esubpu`: the subsidy per MTInputs price unit.
  - Section 12 Revenues: band with the revenue recipe and the rows it uses.
  - Sheet `LegacyDiff` (second tab, 21 rows).
  - 19 named LAMBDAs; every calculated row calls one in 2040. `SUPPLYCOST` and `OTHERTAX` hold the history/projection IF. Plain blocks: history 2022-2024 and projection 2025-2039, with no switch.
  - Method note v0.4.
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.11.xlsx`, `build_v0_11.py`, `check_v0_11.py`, `check_report_v0.11.md`, `PriceProjection_Method_v0.4.md`, README. v0.10 files and method note v0.3 moved to `Old/`.
- **Checks:** PASS.
  - One R1C1 formula per variable and column block.
  - The right column has a LAMBDA on exactly the calculated rows (data lookups and sums stay plain).
  - Python recomputation matches for three variants (LAMBDA expanded; 2039 dragged; LAMBDA copied back over whole rows) and for test scenarios 5-9 and price controls None/Manual.
  - ETS test (scenario 9: power and industry in the ETS, same price): nce, atp and fuel use equal the carbon-tax scenario exactly, and the cement carbon tax is 0.
  - Regression vs v0.10, 2022-2035: 1,902 codes diff 0. Only food & forestry fuel use and its aggregates change (food & forestry 2027: 1,740 to 1,441 ktoe).
- **Results (Egypt):**
  - Baseline 73.8 Mtoe (2027), 92.1 (2035), 104.4 (2040).
  - $20/t: -14.9% throughout.
  - Existing consumer subsidies, 2024 ($/GJ real): gas power 6.2, residential 12.4, industry 4.6; gasoline 13.4 (0.47 $/liter); diesel 15.0 (0.56 $/liter); LPG 18.2; kerosene 12.6; other oil products 11.2 (68 $/bbl).
- **Caveats:**
  1. ETS price is not cap-based: the cap needs emissions (next buckets). Users set the expected permit price through the carbon price inputs.
  2. The full permit price enters the price whatever the auction share: free allocation keeps the marginal incentive.
  3. The existing tax/subsidy split is by sign of txo. Legacy splits into fixed tax, fixed subsidy and floating part, with the same totals.
  4. Changing the last historical year (Settings C10) now also needs the history/projection blocks re-dragged; the LAMBDA column adapts by itself.
  5. 2031-2040 compute but are not a legacy target.
  6. Feebates are revenue-neutral; revenue is not yet calculated.

## 2026-10-09 - Mitigation copy-paste prototype (v0.12: revenues)
- **Task (user):** Revenues, with existing taxes and subsidies and new revenues as separate calculations.
- **Changes:** Three new variables per subsector and fuel, in USD million real 2026 (Settings C13 = 0.041868 PJ/ktoe):
  - `rtx` = fuel use x PJ/ktoe x (existing tax `etx` + VAT `vat` of the price fuel);
  - `rsub` = fuel use x PJ/ktoe x existing subsidy `esub` (cost, positive);
  - `rnew` = fuel use x PJ/ktoe x (`ctxnew` + `ntx` + `ets` x auctioned share + `nce` x VAT rate). Feebates are revenue-neutral.

  Each has a plain block plus a 2040 LAMBDA (`REVENUE`, `NEWREVRATE`). Section 12 shows totals and splits by sector section for each, `rnet` = rtx - rsub, `rtot` = rnet + rnew, `rtot.ref` (scenario 1) and `rtot.chg` (fiscal effect). LegacyDiff gains a revenue-structure row (22 rows).
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.12.xlsx`, `build_v0_12.py`, `check_v0_12.py`, `check_report_v0.12.md`, README. v0.11 files moved to `Old/`.
- **Checks:** PASS.
  - Revenue rows: one formula per block, with a LAMBDA in 2040.
  - Per-row revenues and section-12 totals match the independent Python recomputation (5e-15) in all variants and test scenarios, including ETS auction shares in scenario 9.
  - Sector splits sum to totals; no new revenue in scenario 1.
  - Regression vs v0.11: 2,390 shared codes diff 0.
- **Results (Egypt, USD million real 2026):**

  | | Existing taxes | Existing subsidies | New revenue | Fiscal effect vs baseline |
  |---|---|---|---|---|
  | Baseline 2022 | 1,952 | 27,875 | - | - |
  | Baseline 2030 | 1,919 | 29,036 | - | - |
  | $20/t, 2027 | 1,529 | 26,058 | 3,818 | +8,533 |
  | $20/t, 2030 | 1,656 | 24,372 | 4,168 | +8,568 |

  The $20/t effect is mostly subsidy savings from lower use of subsidised fuels.
- **Caveats:**
  1. Subsidies are price-gap subsidies against supply cost (IMF method), not budget outlays.
  2. Excludes power and electricity, including gas to power. Gas for power has a subsidy of 6.2 $/GJ but no power fuel use yet.
  3. No producer subsidies, no GDP feedback, and existing taxes on fuels only.
  4. Revenue in years after 2030 is not a legacy target.

## 2026-10-09 - Mitigation copy-paste prototype (v0.13: CO2 emissions)
- **Task (user):** Emissions next, CO2 first.
- **Changes:**
  - New variable `co2` per subsector and fuel = fuel use (ktoe) x 0.041868 PJ/ktoe x EF (tCO2/GJ, IIASA, `EF_GHG`; biomass 0), in MtCO2. It has a plain block and the 2040 LAMBDA `EMISSIONS`.
  - Section 13: total, by sector section, by fuel, scenario 1 reference, and change vs scenario 1 in MtCO2 and %.
  - LegacyDiff gains an emissions row (23 rows).
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.13.xlsx`, `build_v0_13.py`, `check_v0_13.py`, `check_report_v0.13.md`, README. v0.12 files moved to `Old/`.
- **Checks:** PASS.
  - `co2`: one formula per block, with a LAMBDA in 2040.
  - Per-row CO2 and the section-13 total match the independent Python recomputation (5e-15) in all variants and test scenarios.
  - Sectors and fuels sum to the total; biomass is 0; there is no change vs scenario 1 before 2027.
  - Regression vs v0.12: 3,196 shared codes diff 0.
- **Results (Egypt, MtCO2, fuel combustion excluding power):**

  | Year | Baseline | $20/t from 2027 |
  |---|---|---|
  | 2022 | 134.8 | 134.8 |
  | 2024 | 188.9 | 188.9 |
  | 2027 | 203.5 | 172.5 (-15.2%) |
  | 2030 | 222.3 | 188.6 (-15.2%) |
  | 2040 | 287.5 | 244.0 (-15.1%) |
- **Caveats:**
  1. 2022-2024 rises +40%. Fuel use in 2023-2024 is modelled from the 2022 base with the large real price falls in the data (devaluation; gasoline -44% real), not taken from energy balances. This is likely too high against observed emissions. Options: use observed energy use for 2023-2024 when available, or start the response from the last price year.
  2. No inventory adjustment of EFs.
  3. No power sector (gas and oil to power are a large share of Egypt's CO2).
  4. No process emissions, CH4, N2O or local pollutants yet.
  5. The ETS price is still not cap-based. Emissions now exist, so a cap can be added next.

## 2026-10-09 - Check: does legacy CPAT show the 2022-2024 jump? (no model change)
- **Source:** a real legacy CPAT run for Egypt, `egypt/supporting/AdHocRebuild/cpat_outputs_egypt_2022_2041.csv` (EG1 baseline, CPAT Outputs sheet, proprietary data), compared with v0.13 scenario 1 (MtCO2, energy-related).
- **Findings:**

  | Sector | Run | 2022 | 2023 | 2024 | 2027 | 2030 |
  |---|---|---|---|---|---|---|
  | Residential | legacy | 15.8 | 17.4 | 22.3 | 23.3 | 24.7 |
  | Residential | v0.13 | 13.9 | 19.0 | 22.3 | 23.3 | 24.6 |
  | Industry | legacy | 59.6 | 58.9 | 69.8 | 80.2 | 89.0 |
  | Industry | v0.13 | 64.3 | 80.3 | 87.2 | 95.4 | 106.2 |
  | Transport | legacy | 53.0 | 57.3 | 48.1 | 72.8 | 116.4 |
  | Transport | v0.13 | 53.4 | 67.4 | 74.8 | 79.8 | 86.3 |

  - Residential: legacy shows the same jump (+41% by 2024), and the trajectories coincide from 2024.
  - Industry: legacy +17% by 2024; v0.13 +36%.
  - Transport: legacy falls in 2024 and then rises steeply; v0.13 rises early and more slowly.
  - Total energy CO2 incl. power: legacy +10% (210.9 to 232.3).
- **Caveats:**
  1. The legacy run used different price data (proprietary vintage; its 2023-2024 prices are not in the NoPropData file).
  2. Sector groupings may differ: legacy industry may exclude fuel transformation; v0.13 industry includes it.
  3. Legacy transport may include dynamics not modelled here.
  4. To confirm in bucket 5 with the legacy price inputs.

## 2026-10-09 - Mitigation copy-paste prototype (v0.14: corrected Egypt price block)
- **Task (user):** The user supplied "the whole price block for Egypt, with correct data" (`Egypt_Price_Data.xlsx`, legacy Prices_dom layout, 2019-2024, plus info columns on sources, reliability and last data year).
- **Changes:**
  - File stored as `data/source/Egypt_Price_Data_2026-10-09.xlsx`.
  - `update_prices_egypt_v0_1.py` replaces the Egypt rows of `data/prices_dom.csv` for the shared columns and adds 2019-2020. Non-Egypt rows are byte-identical; the extra info columns are not used.
  - 291 changed or added cells are listed in `data/prices_dom_changes.csv` and marked bright yellow with red text in the `Prices_dom` sheet.
  - VAT rule: the dataset rate wherever the cell is filled (explicit 0 included); the v0.10 VAT_WEO assumption applies only to blank cells.
  - LegacyDiff gains an "Egypt price data" row (24 rows).
- **What the correction changes:**
  1. Supply costs and retail prices: changes within rounding, except 2024 gas (-1%) and some coal values.
  2. Excise and other taxes of gas and oil products: retail price = supply cost + txo, with VAT rates explicitly 0 (only residential coal 14%). Any VAT sits inside txo, so the v0.10 VAT assumption would have counted VAT twice; under the new rule it no longer applies to Egypt.
  3. Electricity supply costs are now filled (not used yet).
- **Checks:** PASS.
  - Prices, taxes, fuel use, revenue and CO2 match the independent Python recomputation from the updated CSV, including the VAT blank/explicit rule.
  - All 291 changed cells are marked with the new values; no other country is marked.
  - Regression vs v0.13: 1,648 codes that do not depend on Egypt prices diff 0; 1,836 price-dependent codes change by intent.
- **Results (Egypt):**

  | | v0.13 | v0.14 |
  |---|---|---|
  | Baseline CO2 2022 / 2024 / 2030 (Mt) | 134.8 / 188.9 / 222.3 | 134.8 / 187.5 / 220.7 |
  | $20/t effect on CO2 | -15.2% | -13.9% (no VAT on new policies for gas and oil products) |
  | Existing tax revenue 2022 (USD m) | 1,952 (mostly assumed VAT) | 25 |
  | Existing subsidies 2022 (USD m) | 27,875 | 26,077 |
  | $20/t fiscal effect 2030 (USD m) | +8,568 | +7,683 |
- **Caveats:**
  1. The 2022-2024 jump in fuel use and CO2 remains: the corrected block has the same large real price falls. The gap to the legacy run in industry and transport therefore has other causes (to investigate in bucket 5).
  2. Other oil products keep the legacy retail-price rule (sp + txo); the dataset rp is 14% higher for them.
  3. 2019-2020 are stored but not used (history starts 2021).

## 2026-10-09 - Mitigation copy-paste prototype (v0.15: layout)
- **Task (user):**
  - Group columns B:C and I:K.
  - Row grouping at the bottom rather than the top.
  - Mitigation as the first tab, at 75% zoom.
  - Open rolled up.
- **Changes:**
  - Columns: B:C (fuel, sector) and I:K (unit, source, scenario-1 code column) at outline level 1, hidden. The parameter columns D:G sit between them and are nested at level 2, so expanding level 1 shows B:C and I:K and level 2 adds D:G. Adjacent same-level groups would merge in Excel.
  - Rows: summaryBelow, so the + button sits below each group. Every grouped row is hidden on open, and the collapsed flag is on the row after each group.
  - Mitigation is the first and active tab, at zoom 75%.
  - ReadMe roll-up text and version log updated.
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.15.xlsx`, `build_v0_15.py`, `check_v0_15.py`, `check_report_v0.15.md`, README. v0.14 files moved to `Old/`.
- **Checks:** PASS, including the scenario-copy tests. Regression vs v0.14: 3,484 codes diff 0. New layout checks: column levels, rolled-up rows and collapsed flags, first tab, zoom.
- **Caveats:**
  1. Grouping I:K also hides scenario 1's output-code column (K) when rolled up. Scenario 2's code column (Z) stays visible.
  2. With the + buttons below, a subsector's variables expand from the button on the next subsector's heading line (Excel convention for summary rows below).
  3. Not yet opened in Excel.

## 2026-10-09 - Mitigation copy-paste prototype (v0.16: layout)
- **Task (user):**
  - One level of column grouping, not nested; combine the early column groups.
  - Group 2030-2039 for both scenarios.
  - Two beiges: the darker one for the base year; a lighter one for areas that are copy-pasteable within themselves but not draggable to white or the other beige. Apply the lighter one to the LAMBDA formulae and the historical price range.
  - A blank column between scenarios.
- **Changes:**
  - Column groups at one level, all rolled up: B:G (fuel, sector, parameters), I:K, and 2030-2039 in each scenario group.
  - Light beige `EEECE1` (Excel tan, lighter than the base-year `DDD9C4`) on the right-column LAMBDA cells of every calculated row (dark red text kept) and on the 2023-2024 history block of `sp` and `txo`.
  - Each scenario group now ends with a blank spacer column (group width 21). Scenario 2 is `AF:AZ`; copy a whole group, spacer included, to add a scenario.
  - ReadMe colours and roll-up text, README "Add a scenario" and the version log updated.
- **Outputs:** `CPAT_Mitigation_CopyPaste_v0.16.xlsx`, `build_v0_16.py`, `check_v0_16.py`, `check_report_v0.16.md`, README. v0.15 files moved to `Old/`.
- **Checks:** PASS, including the scenario-copy tests with the spacer column (groups 3-9). Regression vs v0.15: 3,484 codes diff 0. New layout checks: one outline level; the 20 grouped year columns; spacer columns empty; light-beige LAMBDA and history blocks (base year darker, projection white).
- **Caveat:** Not yet opened in Excel.

## 2026-10-09 - CPAT-AI-Mitigation-MVP v1.00 (rename) and overview deck
- **Task (user):**
  - A very short PowerPoint (about 7 slides, pitch and user guide combined) explaining in simple terms the purpose, scope and definition of the Excel. Styled like the Excise Diagnostic user guide and the Excise-Fiscal proposal; screenshots from the workbook where possible.
  - Rename the prototype CPAT-AI-Mitigation-MVP-vX.XX and upgrade to v1.0.
- **Changes:**
  - Workbook renamed `CPAT-AI-Mitigation-MVP-v1.00.xlsx`. `build_v1_00.py` (NAME, VERSION 1.00; sheet titles and version log), `check_v1_00.py`, `check_report_v1.00.md`. Content as v0.16.
  - The uploaded v0.16 was byte-identical to ours (no user edits to merge).
  - The folder name `mitigation_copypaste/` and the v0.x file names in `Old/` are kept.
- **Deck:** `presentation/CPAT-AI-Mitigation-MVP_Overview_v1.0.pptx`, 7 slides:
  1. Title, with the rolled-up overview.
  2. Purpose and scope: AI-generated Excel; main mitigation equations; more readable. Definition, in scope, not yet, and the section chain.
  3. Four problems and four fixes: readability; fixed rows -> scenarios left to right; sector/fuel-specific quantities -> left parameter section; scenarios -> copy a block plus an MTInputs lookup.
  4. Scenarios side by side, with three steps to add one and the MTInputs screenshot.
  5. The left section (B:C, D:G).
  6. Two formula styles: direct formulas vs LAMBDA with IF, both draggable.
  7. Next steps: store data, charts and scenario settings in CPAT; collect outputs for many scenarios in an updated MTOutputs (outputs, not inputs, for now); power, other gases, legacy check. Also a "getting started" box.
- **Built by:** `presentation/build_deck_v1_0.js` (pptxgenjs; theme navy 1E2761 / teal 0F7B6C / amber E08A1E / gold D4A24C, Cambria headings, Calibri body, 13.33 x 7.5 in as the Excise decks). Screenshots are rendered from the real workbook by `presentation/make_screenshots_v1_0.py`: LibreOffice recalculation with LAMBDAs expanded, values copy of one sheet, selected cells showing formula text, PDF -> PNG.
- **Checks:** workbook PASS; regression vs v0.16: 3,484 codes diff 0. Deck passes `validate.py`; every slide rendered and inspected (no overflow or overlap).
- **Caveats:**
  1. Screenshots come from LibreOffice, not Excel (fonts and gridlines differ slightly).
  2. The formula screenshot shows the road-gasoline rows only.
  3. pptxgenjs is not a repo dependency: install it (`npm install pptxgenjs`) to rebuild the deck.

## 2026-10-09 - CPAT-AI-Mitigation-MVP v1.01: MTOutputs and Charts
- **Task:** User offline, "press on". First item of the agreed next steps: store results per scenario (slide 7 of the overview deck).
- **Changes:**
  - New sheet `MTOutputs` (second tab): 23 key indicators per scenario block, read from Mitigation by output code.
    - Indicators: fuel use total and by sector, CO2 total and by sector with change vs scenario 1, the revenue components and fiscal effect, carbon and ETS price, retail prices of gasoline, diesel and residential gas, and road-gasoline after-tax price.
    - Lookup: the scenario's code column is the first cell of Mitigation row 5 holding the scenario number; the year columns are read by offset.
    - Adding a scenario: copy the last block below (its number = previous + 1). Adding an indicator: a code stem in column B plus the row formulas.
  - New sheet `Charts`: 4 line charts (CO2, fuel use, fiscal effect, carbon price), one series per scenario block.
  - ReadMe, README and version log updated.
- **Outputs:** `CPAT-AI-Mitigation-MVP-v1.01.xlsx`, `build_v1_01.py`, `check_v1_01.py`, `check_report_v1.01.md`. v1.00 files moved to `Old/`.
- **Checks:** PASS.
  - MTOutputs as shipped: 874 cells equal the Mitigation cell of their code.
  - With blocks 3-9 added by copying alongside the scenario-copy tests: 3,933 cells equal, block numbers 1-9.
  - Charts reference the right rows.
  - Regression vs v1.00: 3,484 Mitigation codes diff 0.
- **Caveats:**
  1. Charts need a new series added by hand after adding a scenario block.
  2. MTOutputs search area on Mitigation is A1:ZZ5000 (room for about 30 scenario groups).
  3. The overview deck still shows v1.00 screenshots (no MTOutputs slide).
  4. Not yet opened in Excel.

## 2026-10-09 - Legacy cached baseline vs MVP; ETS cap design options (no model change)
- **Task:** User offline, "press on". Investigate the 2022-2024 fuel-use jump against legacy, and plan the cap-based ETS price.
- **Finding:** `CPAT 1.0pre_456_NoPropData.xlsb` keeps the cached results of its last full run with real data. That gives fuel use and after-tax prices for all 16 subsectors x 8 fuels, and CO2 by subsector. `legacy_compare_v0_1.py` extracts them (`data/legacy_cached_baseline.csv`) and compares them with MVP v1.01 scenario 1 in `legacy_comparison_v0_1.md`.
- **Results:**
  1. The 2022 base year matches exactly (fuel use, e.g. road gasoline 7,577.3 ktoe).
  2. Legacy prices show no devaluation drop: road gasoline 0.612 / 0.612 / 0.604 $/liter in 2022-2024 against the corrected data's 0.600 / 0.411 / 0.335. So the legacy run used an older price vintage.
  3. With the last historical price year set to 2022 (prices held), the mean absolute gap in fuel use by subsector and fuel for 2023-2030 falls from 45.6% to 17.3%, and the 2030 total from +41.3% to +8.0%.
  4. Legacy has a 2023 dip of about -20% in every sector, which matches its note "sets baseline 2022 and 2023 emissions to equal estimates" (legacy Mitigation row 475). The MVP has no such calibration.
  5. Remaining price differences: legacy residential gas 5.5-6.2 vs 2.3-4.6 $/GJ; coal; legacy other oil products 15 $/bbl from 2025.
- **ETS cap:** `ETS_Cap_Design_Options_v0_1.md`. The legacy method is semi-elasticity based with iteration. In the MVP it would create a circular reference through scenario 1. Four options; recommended A: baseline covered emissions as a refreshed data row, with the current rule kept as override.
- **ETS algorithm (legacy):** `ETS_Legacy_Algorithm_v0_1.md`. Legacy has a fast closed-form price estimate (reduction left / (covered semi-elasticity × effectiveness), effectiveness auctioned 1.0, OBA 0.5) plus a VBA damped log-space goal seek on an override row. The auctioned part acts as a price wedge; the OBA part acts as a shadow price on efficiency only. The cached legacy run is not converged (max error 0.33). The MVP still treats the ETS as fully effective. The proposed benchmark design (OBR share from benchmarks, same split afterwards) is not built yet.
- **Decisions for the user:**
  1. Which price vintage and last historical year to use. The corrected data imply the jump is a real consequence of the devaluation under CPAT's method.
  2. Whether to add legacy's 2023 calibration to emission estimates.
  3. Which ETS option to build.

## 2026-10-09 - Check: legacy treatment of regulated prices (no model change)
- **Question (user):** Does legacy have functionality that keeps regulated prices flat for an economy like Egypt's, based on the pass-through numbers, which would explain the difference?
- **Answer:** No extra mechanism. Legacy uses the same rule as the MVP: chosen = bucketed pass-through, with Egypt values coal 1, gas 0, gasoline/diesel/LPG/kerosene 0, other oil products 1 (MVP 0 by user decision), biomass 1. The floating subsidy is set at its 2024 level, so with pass-through 0 the retail price stays at its 2024 value from 2025 (legacy gasoline 0.6036 $/liter flat). Price-control phase-out is off; there is no fuel subsidy override.
- **The difference is the historical price data.** Legacy's block has no devaluation fall: gasoline 0.612 / 0.612 / 0.604 $/liter real in 2022-2024, against the corrected 0.600 / 0.411 / 0.335. Residential gas 6.24 / 5.81 / 5.54 against 4.63 / 3.21 / 2.33 $/GJ. The MVP's fuel use responds to that 2022-2024 fall; legacy's prices barely move (and legacy calibrates 2023 emissions to estimates).
- **Details and options:** appended to `cpat_excel_new/mitigation_copypaste/legacy_comparison_v0_1.md`.

## 2026-10-09 - Reality check: Egypt gasoline prices 2022-2024 (no model change)
- **Question (user):** Did Egypt's gasoline prices really fall like that?
- **Answer:** Yes, in US dollars.
  - The official 92-octane price rose from about EGP 8.9/liter (2022 average) to 10.3 (2023) and 13.1 (2024).
  - The pound averaged about 19.2, 30.7 and 45.4 per USD over the same years.
  - So the USD price fell from about 0.47 to 0.34 and 0.29 $/liter nominal (-38%). The corrected block falls 0.528 -> 0.376 -> 0.316 (-40%, average over grades).
  - Legacy's block rises (0.54 -> 0.57). The corrected data are right; legacy's price vintage is stale.
- **Details and sources:** `cpat_excel_new/mitigation_copypaste/legacy_comparison_v0_1.md`. Exchange rates are from secondary compilers; the official IMF/World Bank series is still to be pulled.

## 2026-10-09 - CPAT-AI-Mitigation-MVP v1.02: new ETS with cap, benchmarks, volatility, goal seek
- **Task (user):** "The ETS is split into a shadow price and a carbon tax near equivalent piece. We can update old model with a volatility dependent effectiveness. Anyway, implement ETS then as you wish."
- **Built** (`cpat_excel_new/mitigation_copypaste/build_v1_02.py`, method `ETS_Method_v0_1.md`):
  - **Cap.** Baseline covered emissions use option A of `ETS_Cap_Design_Options_v0_1.md`. The data rows `bco2` hold scenario 1's CO2 by ETS sector; the build fills them from a LibreOffice recalculation, and check row `bco2.chk` = 0. The cap follows MTInputs 85-87. There is no circular reference.
  - **Price.**
    - Fast estimate: LN(cap/baseline) / (effective semi-elasticity × volatility adjustment). This is the log form; legacy's is linear.
    - Or the override row `ets.ovr` (MTInputs `D_ETSPriceOverride` = Yes; a blank cell takes the carbon price path, the v1.01 behaviour).
    - `ets_goalseek_v0_1.py` replaces the legacy VBA. It uses the same damped log-space step on the workbook's `ets.next` proposal and writes a CSV for pasting into `ets.ovr`.
  - **Benchmarks** by sector group (pow, tra, res, ind; typed rows, assumption 1.0 → 0.8) replace the auction share. OBR share = benchmark.
    - The permit price splits into a tax-equivalent wedge in `atp`, × (1 − OBR).
    - The OBR part, × OBR, is a shadow price on the efficiency margin, like feebates.
    - Revenue is the permit price on the auctioned part.
  - **Volatility adjustment** = (1 + policy risk × impact) / (1 + volatility × impact), from MTInputs 95-97 and a Settings table (legacy ETS+LTS). It equals legacy's 1/1.1 at Medium and scales the tax-equivalent price and the fast estimate.
  - Section 13 adds CO2 by ETS sector, covered emissions, cap, gap and next price. MTOutputs adds covered CO2 and the cap. LegacyDiff is updated.
- **Checks** (`check_report_v1.02.md`, PASS):
  - Scenarios 1-2 are unchanged against v1.01 (3,476 codes, diff 0).
  - The cap-based ETS test matches Python to 1e-14.
  - ETS with a blank override, volatility Zero, policy risk 0 and benchmarks 0 equals the carbon tax exactly.
  - Goal seek: the fast estimate misses the cap by at most 2.9%. After 12 iterations the gap is 0.46%. Prices go 2030 26.6 → 24.6 and 2035 54.3 → 52.5 $/tCO2 (cap −2% → −15% on iron & steel and cement).
- **Caveats:**
  - Power is not modelled, so covered power emissions are 0.
  - Benchmarks are relative to current intensity, with output proxied by fuel use (OBR share = benchmark).
  - Semi-elasticities are legacy's Egypt sector values (fast estimate only).
  - VAT on new policies applies to the tax-equivalent ETS cost.
  - The `bco2` data must be refreshed (rebuild, or paste scenario 1's `co2.sec`) when the baseline changes; `bco2.chk` shows it.
  - Goal-seek convergence is slow (about 12% error reduction per iteration), as with legacy's damping.
  - MTInputs rows 90-92 (auction proportion) are no longer read.

## 2026-10-09 - CPAT-AI-Mitigation-MVP v1.03: multiple scenarios (definitions, VBA batch run, stored results, comparison)
- **Task (user):**
  - Store results as values so multiple scenarios can be kept: the baseline, which should not change, and the scenarios.
  - Put a table of scenario definitions in MTInputs, one column per scenario, and add combined packages. No High/Low international price sensitivity, because it deviates from legacy.
  - Write a VBA macro that cycles through the definitions, runs the ETS goal seek when the ETS is on, and stores the results as values.
  - Add a comparison tab for one year.
  - The user chose an embedded .xlsm.
- **Built** (`cpat_excel_new/mitigation_copypaste/build_v1_03.py`, `Scenarios_Method_v0_1.md`):
  - **MTInputs:** J = baseline, K = live policy scenario, L-N = definitions with a Run? flag in row 4:
    - 3: $20 carbon price;
    - 4: Package A, carbon tax $10 → $50 by 2030 + feebates on transport and industry;
    - 5: Package B, ETS on power and industry (cap −5% → −20% by 2035) + carbon tax $25 elsewhere.
  - **VBA module CPATScenarios** (`CPATScenarios_v0_1.bas`): RunAllScenarios, StoreBaseline, StoreLiveScenario, SolveETSLive, ClearStoredScenarios. The ETS goal seek ports `ets_goalseek_v0_2.py`.
  - **StoredResults:** values only, 25 MTOutputs rows per scenario ID.
  - **ScenarioCompare:** one year, baseline + 8 scenario slots; levels, differences and % differences.
  - **`vba_project_v0_1.py`** writes the VBA project without Excel: MS-OVBA compression, dir/PROJECT/PROJECTwm streams, encrypted protection fields (checked by decrypting legacy's), and a compound file writer. It embeds the project as an .xlsm.
  - StoredResults ships filled by a Python emulation of the macro.
- **Checks** (`check_report_v1.03.md`, PASS):
  - Mitigation is unchanged against v1.02.
  - Stored baseline and definition 3 equal the recalculated MTOutputs blocks.
  - Definitions 4 and 5 equal the independent Python recomputation (5 at its stored ETS prices). Package B meets its cap within 0.49%: ETS price $41 (2030), $77 (2035), $147 (2040).
  - ScenarioCompare equals StoredResults.
  - **The embedded macro runs in LibreOffice:** RunAll stored the same four blocks as the emulation (max relative difference 8e-15) in 27 s, and restored column K and the override row.
- **Caveats:**
  - **The VBA project has not been opened in Excel.** It is tested only with oletools and LibreOffice. If Excel rejects it or shows an empty project, import `CPATScenarios_v0_1.bas` and save as .xlsm.
  - The macro needs macros enabled. Excel may warn about files from the internet (unblock in file Properties).
  - The batch runs one policy scenario at a time through Mitigation group 2. Stored indicators are the MTOutputs rows.
  - The Package A result is driven by assumptions: the carbon price continues linearly after 2030 (template switch) and feebates continue linearly. CO2 is −31% vs baseline in 2030.
  - LibreOffice saves booleans as `=TRUE()` in .xlsx, which affects only the test's comparison.

## 2026-10-09 - CPAT-AI-Mitigation-MVP: documentation completed, advanced-features deck, hand-over folder
- **Task (user):**
  - Complete the documentation.
  - Adjust the team deck on AI and the MVP if visible; otherwise write a short 4-slide deck on the MVP's advanced features.
  - Update the reference .md files, then commit and push everything to main.
- **Deck:**
  - The only deck in the repo is the generated v1.0 overview; `cpat_excel_new/working_version/` holds an identical copy. No separate team deck is visible.
  - New 4-slide deck written: `presentation/CPAT-AI-Mitigation-MVP_AdvancedFeatures_v1.0.pptx`. Slides:
    1. what is new in v1.01-v1.03;
    2. the cap-based ETS and its two channels, with the Package B price path;
    3. the scenario table and the RunAllScenarios macro;
    4. ScenarioCompare, CO2 by scenario in 2030, quality numbers and next steps.
  - Screenshots and chart data come from the v1.03 workbook (`make_screenshots_advanced_v1_0.py`); the deck is built by `build_deck_advanced_v1_0.js` (pptxgenjs).
  - Validation passed; visual check in LibreOffice.
- **Documentation:**
  - `CPAT-AI-Mitigation-MVP_Documentation_v1.03.docx` (25 pages) now has appendices:
    - A: differences with legacy CPAT, read from sheet LegacyDiff;
    - B: version log, from Settings;
    - C: verification summary, from `check_report_v1.03.md` (all sections, Overall PASS);
    - D: presentations.
  - Reference .md files updated:
    - folder README: deck and appendices;
    - root README: key documents, the mitigation line and the new `working_version/` line;
    - NORMS key documents: MVP row;
    - AGENTS: the MVP builds on Linux with openpyxl + LibreOffice, an exception to the Excel COM rule.
  - `make_screenshots_v1_0.py` now finds the archived v1.00 builder in `Old/`.
  - The check report's goal-seek heading now names `ets_goalseek_v0_2.py` (label only).
- **Hand-over folder** `cpat_excel_new/working_version/`:
  - added the v1.03 .xlsm, the Word documentation, the advanced-features deck and the .bas module;
  - kept the v1.0 overview deck;
  - moved v1.00 to `working_version/Old/`.
- **Caveats:**
  - The decks' screenshots are LibreOffice renders, so fonts may differ slightly from Excel.
  - If a separate team deck on AI exists outside the repo, it was not updated. Share it to have the new slides merged in.
  - The VBA project still needs its first run in Excel (see the v1.03 entry).

## 2026-10-10 - CPAT-AI-Mitigation-MVP v1.04: first-run test for Excel (CheckBatchRun, sheet MacroCheck)
- **Task (user):** "Go for it", the open item being the macro's first run in Excel. Excel is not available in this Linux environment, so the run itself is left to the user. This version makes that run a one-click, self-reporting test.
- **Built** (`build_v1_04.py`, module `CPATScenarios_v0_2.bas`):
  - Macro **CheckBatchRun**:
    1. checks the 12 lookups the macros rely on (MTInputs scenario columns and parameter rows, Mitigation rows by code and the scenario-2 columns, MTOutputs blocks);
    2. compares the host's own recalculation of the baseline with the stored baseline. Excel evaluates the LAMBDA column natively; LibreOffice needed it expanded;
    3. reruns RunAll, including the ETS goal seek, and compares every stored value with the value stored before.
  - The report goes to the new sheet **MacroCheck** (Result PASS / CHECK / FAIL, 1e-6 relative tolerance). Its first-run checklist (unblock, enable macros, run, fallback import) is in `Scenarios_Method_v0_1.md`.
  - The v0.1 macros are unchanged; Mitigation is unchanged.
  - Deck updated to `CPAT-AI-Mitigation-MVP_AdvancedFeatures_v1.1.pptx` (the next step names the test).
  - Word documentation `..._v1.04.docx`; `working_version/` refreshed (v1.03 files to `working_version/Old/`).
- **Checks** (`check_report_v1.04.md`, PASS):
  - Regression vs v1.03: 3,642 codes, difference 0.
  - CheckBatch run in LibreOffice: Result PASS; 12 lookups found; baseline difference 4.4e-15; 1,900 stored cells after the batch rerun, largest relative difference 7.6e-15; 27 s.
  - Negative test with the row code ets.ovr broken: Result FAIL naming exactly that lookup, returned −1, no run.
- **Caveats:**
  - Still not run in Excel. The MacroCheck report from that first run is the evidence to look for.
  - LibreOffice's VBA layer reports itself as "Microsoft Excel 11.0" in the report's Started line, so that line names the host but does not prove Excel.
  - In Excel, small differences against the LibreOffice-computed stored values are possible but expected below 1e-6. Anything larger shows as CHECK, with a count.

## 2026-10-10 - CPAT-AI-Mitigation-MVP v1.05: power generation costs and electricity prices (section 3)
- **Task (user):** plan the prices part of power (generation costs and downstream electricity prices), with simplifications that match the existing model; lay out the interactions (generation mix, systemic cost of renewables); write a method note with pseudocode; then build it following the legacy outline and sub-table letters.
- **Method note:** `PowerPrices_Method_v0_1.md`:
  - interactions I1-I8 and simplifications S1-S7;
  - method by legacy sub-table (A, B, C, D, E, F, G, J) with pseudocode;
  - workbook layout, interim data, validation and open decisions.
- **Data:** `extract_power_v0_1.py` reads legacy Mitigation rows 2950-3935, locating each row by table id and label. It writes `data/power_tech.csv`, `power_paths.csv`, `power_params.csv` and `legacy_power_reference.csv`.
  - Correction found while validating: the legacy B2 capital cost is the base-year cost. It already includes the capex time factor of 2022 (1.05 for wind and solar), so `cax0` is divided by that factor.
- **Built** (`build_v1_05.py`):
  - **Mitigation section 3**, with legacy sub-tables in legacy order:
    - A: inputs;
    - B: amortised fixed cost of the stock, with vintage averages;
    - C: levelised fixed cost of new plants;
    - D: variable cost and new policy cost per kWh;
    - E: generation cost by type and the averages;
    - G: residential and industrial electricity prices, subsidy gap and revenues;
    - J: storage cost of variable renewables.
  - Sheet `Inputs_power` (one row per type, with derived constants; one row per end-user group) and data sheets `PowerTech`, `PowerPaths`, `PowerParams`.
  - 14 new LAMBDAs, 44 in all.
  - 3 new MTOutputs indicators: residential and industrial price, and the average generation cost.
  - 7 new LegacyDiff rows; the version log is updated.
- **Checks** (`check_report_v1.05.md`, PASS):
  - **Method = legacy:** the Python reference, fed with legacy inputs, reproduces 80 cached legacy rows (A13, B2, B6, B8, C3, C7, D4, E3, J, G1/G2 supply cost and price), 2022-2040. Largest relative difference: 5.6e-17.
  - **Workbook = Python** for all 28 power variables in every scenario test (carbon tax, ETS, fuel price reform, price source, price controls), and in the stored packages (ID 5 is the ETS on power).
  - One formula per power block; the right column calls a LAMBDA (1,669 rows).
  - Regression vs v1.04: 3,642 codes, difference 0. The macro tests still pass.
- **Results (Egypt, real 2026 $/kWh):**
  - The average generation cost is about 0.069.
  - Residential price before policy: 0.061 (2022) and 0.042 (2030). Industrial: 0.064 and 0.049.
  - A $20 carbon price adds about 0.009 $/kWh in 2027 and raises about USD 1.8 bn a year.
  - The VRE share rises from 5% to 28% by 2040. Storage then adds about 0.011 $/kWh to wind and solar.
- **Caveats:**
  - Generation shares, investment shares and consumption are **interim data** from the legacy baseline run. A policy scenario therefore changes power prices only through fuel and carbon costs, not through its own generation mix, until the engineer model exists. Power emissions are still 0, so a cap that covers power cannot be met by power.
  - Fuel prices are spot prices from section 2. Legacy uses a 5-year moving average.
  - Historical electricity prices come from the corrected Egypt block: 0.054 $/kWh nominal residential in 2022, VAT 0. Legacy used 0.112 including 15.6% VAT.
  - Power revenues value the ETS at the tax-equivalent price. They are not yet in the section 12 totals.
  - The coal implicit cost and the extra shadow price are not built (both are 0 for Egypt).
  - Open decisions are in section 8 of the method note.

## 2026-10-10 - CPAT-AI-Mitigation-MVP v1.06: power engineer model; legacy bugs report
- **Task (user):** "Build the power model. Align except for bugs, and report both bugs and improvements in new human readable word doc."
- **Legacy reading:**
  - The legacy .xlsb was converted to .xlsx with LibreOffice so its formulas could be read (named ranges come through as `#NAME?`; their values were taken from the cached cells).
  - For each row, the formulas were compared across all year columns.
  - `extract_power_v0_2.py` adds the engineer-model inputs (A4 plant data, retirement, investment permission, scale-up rates, demand by subsector, balance shares, dispatch and investment parameters). It also writes `legacy_engineer_reference.csv`: cached inputs and results of the baseline and the carbon-tax run (rows +6331).
- **Python reference** `power_model_ref_v0_1.py`:
  - covers H demand, I dispatch, J storage, K retirement and capacity, L investment, plus the price chain;
  - with the bugs switched on (and the legacy demand shock), it reproduces 208 legacy rows of both runs, 2022-2040, to 2e-13;
  - each of the 10 bugs is a switch (`FIXES`).
- **Bugs** (details and effects in `PowerModel_BugsAndImprovements_v1_0.docx`):
  1. services demand grows with the industrial multiple from 2024 (row 3731);
  2. the autonomous-efficiency exponent uses eps_F instead of eps_U;
  3. the income factor on the GDP elasticity is applied twice, anchored to 2023 ($M3686);
  4. the long-term storage allocation is applied twice;
  5. investment USD uses the capital cost at factor 1;
  6. the investment logit is normalised by non-permitted types;
  7. early coal retirement mixes nameplate and effective MW;
  8. the effective coal retirement is uncapped;
  9. the feebate rebate is x 1.22;
  10. the capacity check uses last year's plan.
  - Plus 2 minor defects: a stray `$G4291` and a total-row multiplier on an empty cell.
  - With all fixes, on the legacy Egypt inputs: 2040 generation +12.6% in the baseline and +9.2% in the carbon-tax run; 2040 power CO2 +10.8%.
- **MVP corrections of v1.05:**
  - VAT on the power policy term (legacy row 3614 applies it);
  - LT storage allocation applied once;
  - VRE share includes other renewables.
  - None of the three affects Egypt numbers.
- **Built** (`build_v1_06.py`):
  - **Section 3 now has legacy H, I, K and L**:
    - D6 shadow price and rebate, D9 levelised variable cost, E1 investment cost;
    - the interim rows of v1.05 (generation shares, investment shares, consumption) are now modelled;
    - 33 named LAMBDAs added;
    - the look-ahead rows `dsum` and `lvc` are plain formulas in every column.
  - **Power CO2 enters section 13** (total, by fuel, ETS sector power).
  - Inputs_power gains plant, demand-group and subsector tables; new data sheet PowerDemand.
  - Settings C19 switches the 2023-2024 demand shock (off by default).
  - 5 MTOutputs indicators; LegacyDiff rows; version log.
  - Method note `PowerModel_Method_v0_1.md`. Price note corrected (VAT, LT allocation, VRE share).
- **Checks** (`check_report_v1.06.md`):
  - legacy test: 208 rows x 2 runs;
  - the workbook equals the Python reference (all fixes) for every power row in every scenario test (carbon tax, fuel price reform, feebates, price source, ETS with power, price controls);
  - power properties hold;
  - regression vs v1.05: 3,484 codes unchanged (388 intended changes: power block and CO2 including power);
  - macro tests pass.
  - The shipped-file check now allows LibreOffice errors in the power block and the section-13 CO2 rows, because D9 looks ahead to the LAMBDA column. The expanded variant checks their values.
- **Caveats:**
  - **2023 jump:** MVP electricity demand jumps +22% in 2023 (generation 220 to 269 TWh), and investment surges in 2023-2025 (scale-up limits leave some need unplaced; gas then runs harder). The cause is the corrected Egypt electricity prices, which fall in USD (0.075 to 0.054 $/kWh, devaluation), and demand responds to that fall. The fuel blocks show the same jump (+25% fuel use). This is the open price-vintage decision in TODO. The legacy demand shock (Settings C19) largely offsets it.
  - **Earlier power response:** investment looks ahead (D9 includes future carbon prices), so power changes before a carbon price starts, as in legacy.
  - **ETS with power:** an ETS covering power now has power emissions in its cap. The Package B ETS price is about $124/t in 2027 (the goal seek converges within 17 iterations).
  - **Demand not linked to sectors:** electricity demand (H) is its own block, not linked to the sector fuel blocks. Electricity is not yet a fuel there.
  - **Features not built:** PPAs, manual capacity and investment, planned coal retirement data, elasticity-model power demand, and a carbon price on power use (all off or 0 for Egypt).
  - **Not run in Excel:** not yet opened in Excel. The MINIFS function (`_xlfn.MINIFS`) requires Excel 2019 or later, which the LAMBDAs already require.

## 2026-10-11 - CPAT-AI-Mitigation-MVP v1.07: short formulas, planned retirement, demand memo lines, power indicators

- **Task (user):**
  - "Accept the dip. Keep sectoral policies separate. Add planned retirement. Keep all formulae short - lambdas can have more complexity. Add memo lines for sectoral demand. Add power indicators."
- **Decisions recorded:**
  - The 2023 power-demand jump is accepted. It comes from the USD electricity prices after the devaluation; Settings C19 stays 0.
  - Sector policies stay separate from power demand. Electricity is shown as memo lines only, not as a fuel in the sector blocks.
  - Both decisions are in LegacyDiff and TODO.
- **Built** (`build_v1_07.py`):
  - **Short formulas:** every cell formula has at most 100 characters, LAMBDA calls included. The longest is 99, across 110,473 formulas. The rule is in NORMS section 3. Tools used:
    - 53 defined names for data tables and fixed ranges;
    - helper rows: carbon-price path years and nominal factor; ETS on, volatility, impact, cap change, log ratio; policy rates by sector `ctr`, `etr`, `shpx`; `gp.src`, `gp.adj`, `gpn`; `txf`; `ecmin`, `ecshr`;
    - helper columns in Inputs_prices (data by year), MTOutputs and ScenarioCompare; Settings rows 20-25;
    - new fuel-use parameters: AEEI factor and efficiency exponent (Inputs);
    - the LAMBDA `ROWLABEL` for column H;
    - history lookups moved into the SUPPLYCOST, OTHERTAX and POWERPRICE LAMBDAs.
    - There are now 85 LAMBDAs.
  - **Planned retirement for every type:**
    - PowerTech gains a `planned_retirement` switch and PowerPaths gains `pret|<type>` rows (`extract_power_v0_3.py`).
    - Retirement = MIN(planned retirement or the rate-based retirement, + early coal retirement, capacity).
    - Python reference: `power_model_ref_v0_2.py`.
  - **Memo lines `eldem`:**
    - electricity demand by subsector in each sector block, and by sector in section 11;
    - demand by demand group in H.
  - **Power indicators:** 18 more MTOutputs indicators (51 in all) and 6 power charts (10 in all).
  - Method note `PowerModel_Method_v0_2.md`; Documentation v1.07; `ets_goalseek_v0_3.py`, `legacy_compare_v0_2.py`.
- **Checks** (`check_report_v1.07.md`, Overall PASS):
  - formula length;
  - regression vs v1.06: 4,404 shared codes, max difference 0;
  - planned retirement switched on for gas and oil = Python reference, and retirement = the data;
  - memo lines add up;
  - legacy test 208 rows × 2 runs;
  - macro tests;
  - Package B goal seek converges within 17 iterations.
- **Caveats:**
  - **No planned-retirement data for Egypt.** The legacy file has none, so every switch is No and the results are unchanged. Enter data in PowerPaths and set the switch in PowerTech.
  - **Name clash found during the build.** A defined name must not equal a LAMBDA name, ignoring case: `ETSSemi` collided with `ETSSEMI` and gave #NAME?. It was renamed to `SemiEl`, and the builder now asserts against such clashes.
  - **Column H uses a LAMBDA (ROWLABEL).** Without LAMBDA support (LibreOffice, Excel before 365) the descriptions show errors. The check accepts this, as it does for the right column. Output codes stay plain formulas.
  - **The shipped-file check now also allows LibreOffice errors in the memo lines.** They read power demand, which depends on the D9 look-ahead.
  - **Memo grouping.** In the memo by sector, food & forestry counts under buildings, as in the sector blocks. The legacy demand groups put it with industry.
  - **Not run in Excel yet.**

## 2026-10-11 - CPAT-AI-Mitigation-MVP v1.08: works without LAMBDA; combined deck v2.0; documentation refresh

- **Task (user):**
  - "Main functionality should work without lambdas. If formulae become too long break up. Do and test."
  - Then: finalise, redo the PowerPoint and the documentation for the whole project. The integrated AI-tools deck is not in the repository, so the deck is one simple combined deck of the MVP, based on the existing MVP decks.
- **Built** (`build_v1_08.py`):
  - **No LAMBDA in cells.** Every cell holds a plain formula of at most 100 characters (longest 99), the right column (2040) included.
  - **Column H** reads the member labels from a table on Mapping (sector|fuel key) instead of the LAMBDA ROWLABEL.
  - **The 84 named LAMBDAs are an optional library.** New sheet `Lambdas` gives the LAMBDA formula of each calculated row (as text). The right column is no longer light beige.
  - Method notes renewed so they describe the current state: PriceProjection v0.5, Scenarios v0_2, PowerPrices v0_2, PowerModel v0_3.
  - Documentation v1.08.
- **Deck:** `presentation/CPAT-AI-Mitigation-MVP_v2.0.pptx`.
  - 11 slides, built by `build_deck_v2_0.js` from `make_screenshots_v2_0.py` (screenshots and stored-scenario data of v1.08).
  - It replaces the Overview v1.0 and Advanced features v1.1 decks, which moved to `presentation/Old/`.
- **Checks** (`check_report_v1.08.md`, Overall PASS):
  - no cell calls a LAMBDA;
  - the shipped file recalculated by LibreOffice has zero error values in all 29 sheets;
  - values equal the Python recomputation as shipped, and also with each row's LAMBDA from sheet `Lambdas` copied over the whole row;
  - regression vs v1.07: 4,630 shared codes, max difference 0;
  - macro tests and the ETS goal seek pass.
- **Caveats:**
  - **Excel 2019 or later still needed** for MINIFS (power dispatch and investment). LAMBDA is no longer required.
  - **New sector|fuel pairs need a label.** A new pair on Mitigation (e.g. a new fuel) needs a row in the member-label table on Mapping, or its description shows #N/A. The table has 20 spare rows in its defined name.
  - **Deck charts render slightly differently in LibreOffice.** The deck was checked in LibreOffice renders; in the CO2 chart there the year labels are dense.
  - **The integrated AI-tools deck (user guide for Claude and GitHub Copilot) was not found in this repository** and was not redone.
  - **Not run in Excel yet.**

## 2026-10-11 - CPAT-AI-Mitigation-MVP v1.09: existing carbon pricing; power in the main thread; full consistency check

- **Task (user):** "Do a full check for consistency. One could now bring in existing carbon pricing data; ensure power is fully linked up with revenues and energy use and emissions built into the main thread."
- **Built** (`build_v1_09.py`, data step `extract_ecp_v0_1.py`):
  - **Existing carbon pricing.**
    - New sheet `ECP`: the legacy ECP data (carbon tax rate, ETS permit price by year 2021-2025; sector and fuel coverage; 52 countries) and the values of the country in Settings C4 (real prices, coverage of the last data year). Egypt has no entries, so all values are 0.
    - Section 1 reads MTInputs 208-214: `ExistingCTApply`, `ExistingCTGrowth`, `ExistingETSApply`, `ExistingETSGrowth`, `CTaxComplimentaryToETS`. Prices used = data x switch, growing after the last data year; the level embedded in the price data is held after the last historical price year.
    - Per subsector and fuel: `xcp` (level) and `xcpd` (change against the price data). `xcpd` enters the after-tax price and existing-tax revenue; `rxcp` is a revenue memo. Power: `xck` in the variable cost.
    - `CTaxComplimentaryToETS` = Yes: the new carbon tax covers 1 - existing ETS sector coverage (legacy rows 2270-2288).
  - **Power in the main thread.**
    - Section 11: generation fuel use of coal, gas, oil and biomass (row `ener.pow`) is in the total and in the fuel rows.
    - Section 12: power rows in every total: fuel taxes and subsidies on generation, change in existing carbon pricing, electricity VAT (also on the passed-through policy cost) and subsidy gap, new policies with the ETS at the permit price (new sub-table M by generation type). This fixes the power ETS revenue (it was valued at the tax-equivalent price, about 10% high).
    - Section 13 already had power CO2 (v1.06); new check row `co2.chk`.
  - **Consistency fixes found by the review:**
    - stale texts in the workbook (section 12 and 13 notes said "no power", PowerPaths title said interim data, data-sheet sources named old extractors, method-note names);
    - method notes renewed to the current state: PriceProjection v0.6 (with existing carbon pricing, 5b), PowerPrices v0_3 (interim data text removed, revenues and section 12), PowerModel v0_4, Scenarios v0_3 (58 indicators);
    - `ets_goalseek_v0_4.py` and `legacy_compare_v0_3.py` import v1.09;
    - folder and root README (sections, variables, indicator count, files), TODO.
  - MTOutputs: 58 indicators (adds fuel use of power, revenues of power, existing carbon pricing revenue and prices). Deck `presentation/CPAT-AI-Mitigation-MVP_v2.1.pptx` (from v1.09) and documentation v1.09.
- **Checks** (`check_report_v1.09.md`): Overall PASS.
  - 133,331 formulas in 30 sheets, none over 100 characters, no LAMBDA calls, zero LibreOffice errors.
  - Values equal the independent Python recomputation (fuel and revenue totals with power, power revenues by type) in every scenario test.
  - Existing carbon pricing with non-zero test data (shipped scenarios, growth, switched off, complementary carbon tax) equals Python; `xcpd` = 0 up to 2024.
  - Regression vs v1.08: 4,592 shared codes, max difference 0 outside the intended changes (fuel and revenue totals, their reference and change rows).
  - New consistency section: check rows and identities in every column, file references in the workbook texts, versions, no stale texts, one current method note each.
- **Caveats:**
  - **Fuel and revenue totals now include power.** The totals of v1.09 are not comparable with v1.08. Fuel use 2022 is 86.4 Mtoe, of which 36.7 Mtoe is power generation. Revenue totals add power: in 2022 USD 14.4 bn of subsidies on the gas and oil used for generation and a USD 6.2 bn electricity subsidy gap.
  - **The electricity subsidy gap is signed** (negative when the tariff is above the modelled supply cost). Fuel subsidies are positive parts only. The two do not overlap: the supply cost already uses the subsidised fuel price.
  - **Existing carbon pricing simplifications** (listed in LegacyDiff):
    - one ETS per country (national, else the regional EU ETS; DEU and AUT keep only their national ETS);
    - a single carbon tax rate (legacy fuel-specific rates `ecr.<fuel>` are not used);
    - coverage of the last data year (coverage carried forward over blank years);
    - existing ETS revenue at the full permit price (free allocation not used);
    - existing and new ETS add up where both cover a sector.
  - **Only the change against the price data enters prices.** The IMF retail prices already contain the existing carbon pricing of their year, so the level is a memo (`xcp`, `rxcp`). Coverage differences within a price fuel cannot be in the price data, so `xcpd` uses the embedded level of the same subsector and fuel.
  - **Not run in Excel yet.**

## 2026-10-11 - CPAT-AI-Mitigation-MVP v2.2: one version number; register of bugs, issues and upgrades; quick start

- **Task (user):** "Make sure we have consistent human readable bug list and potential issues or upgrades list in word format. Update all versions to 2.2. Check for consistency. Add quick start guide / intro at beginning of word document. Include guide to structure of repo."
- **Built:**
  - **One version number.** Every current file of `cpat_excel_new/mitigation_copypaste/` has version 2.2: the workbook `CPAT-AI-Mitigation-MVP-v2.2.xlsm`, `build_v2_2.py`, `check_v2_2.py`, the data-step and helper scripts, `CPATScenarios_v2_2.bas`, the method and background notes, the Word documents and the deck `presentation/CPAT-AI-Mitigation-MVP_v2.2.pptx`. The previous versions are in `Old/` (and `presentation/Old/`). Results are unchanged.
  - **Register of bugs, issues and upgrades.** `issues_register_v2_2.py` is the single source. It feeds `CPAT-AI-Mitigation-MVP_BugsIssuesUpgrades_v2.2.docx` (`make_issues_doc_v2_2.py`) and the appendices of the documentation. IDs:
    - L1-L16: legacy CPAT bugs, with the effect of each power fix computed when the document is built;
    - M1-M9: bugs of earlier MVP versions, all fixed;
    - I1-I14: known issues;
    - LD<#>: the LegacyDiff rows of type Not yet, Simplification or To check;
    - U1-U16: upgrades by priority.
    - It replaces `PowerModel_BugsAndImprovements_v1_0.docx` (now in `Old/`).
  - **Word documentation v2.2** opens with `QuickStart_v2_2.md`: what the MVP is, a quick start in seven steps, how to rebuild, a guide to the repository (top level, `cpat_excel_new/`, the MVP folder) and to the workbook sheets. Appendix E lists the register by ID, appendix F the upgrades by priority.
- **Consistency fixes:**
  - LegacyDiff: the row "Existing ETS and existing carbon taxes: Not modelled" removed (modelled since v1.09); the formula-structure row described the v1.07 LAMBDA right column. The sheet now points to the register.
  - Notes and scripts naming old versions (ETS note: ets_goalseek_v0_2; VBA module comment; script headers).
  - `legacy_comparison` was the v1.01 note and its script wrote another file name. It is re-run on v2.2: mean absolute difference to legacy 45.6% with prices to 2024, 17.3% with prices to 2022 (unchanged, cause I4).
  - **Repository guide:** the root README described `egypt/`, `egypt-final/` and `cpat_excel_new/standalone_working_version/`. These were removed on 2026-10-09 (commit f070c20); the kernel working version is in `cpat_excel_new/old/standalone_working_version/`. The root README layout is corrected; `NORMS.md`, `TODO.md` and `AGENTS.md` got a note at the top (their Egypt paths are kept for reference).
- **Checks** (`check_report_v2.2.md`, Overall PASS):
  - regression vs v1.09: 5,556 shared codes, max difference 0;
  - every current versioned file of the folder carries 2.2 (27 files);
  - register IDs unique, all 17 LegacyDiff rows of the listed types included;
  - no stale texts, file references valid, all earlier checks as v1.09.
- **Caveats:**
  - **Tool scripts were renamed, not rewritten.** Their headers say so. Old cross-references inside `Old/` still use the old names.
  - **The Egypt material is no longer in the repository.** NORMS, TODO and AGENTS still describe the Egypt workflow; decide whether to delete those parts.
  - **Not run in Excel yet** (I1).

## 2026-10-12 - CPAT-AI-Mitigation-MVP v2.3: last year 2041 as the LAMBDA year; legacy moving-average fuel prices for power; Egypt parts removed from NORMS, TODO, AGENTS

- **Task (user):** "Delete the Egypt parts from NORMS, TODO and AGENTS. For moving average prices, we go what is in legacy. Please note that the LAMBDA functions should be a full year as part of the flow. Let's go with 2041, which will be our last year."
- **Built:**
  - **Repository documents.** The Egypt parts are deleted from `NORMS.md` (key documents, reference implementation, EGY flag rule, working-version folder, task-inventory step, superseded copies), `TODO.md` (Egypt checks and modelling sections, kernel conventions) and `AGENTS.md`.
  - **Horizon 2041, the LAMBDA year** (`build_v2_3.py`):
    - the last year is 2041 (legacy horizon of the power block); power data re-extracted to 2041 (`extract_power_v2_3.py`; 2022-2040 unchanged);
    - every calculated cell of 2041 calls its named LAMBDA (light beige, dark red text) and feeds the results like any year;
    - 2022-2040 keep plain formulas of at most 100 characters.
    - NORMS section 3 records the rule.
  - **Fuel prices for power as legacy (A18):** new row `fpp` = centred 5-year moving average (years t-2 .. t+2 within 2022-2041, blanks ignored as legacy) of the price before new policies. Settings C21 `SpotPow` = Yes gives spot prices (legacy option). `vbc` uses `fpp`.
  - **Levelised variable cost window to 2041** as legacy (closes I9/U12 of v2.2).
  - **Hosts without LAMBDA support** (Excel before 365, LibreOffice) still calculate 2022-2040:
    - `tvcx` = IFERROR(tvc, last year's tvc) feeds the D9 window;
    - `fpp` averages with AGGREGATE (errors skipped), so the 2041 prices are left out of 2039-2040 there.
  - **Consistency:** the total rows were labelled "All subsectors" although they include power since v1.09; the label is now "All sectors" (M10). Register updated: I2, I9, I10 rewritten; moving average and window upgrades closed.
  - **VBA year count (M11).** The macros had 19 years as a constant; found by the check and set to 20, and the builder now asserts it.
  - Every current file renumbered to 2.3 (previous versions in `Old/`); documentation, register and deck v2.3; legacy comparison re-run (unchanged: 45.6% / 17.3%).
- **Checks** (`check_report_v2.3.md`): Overall PASS.
  - 139,801 formulas; 4,888 cells call a named LAMBDA, all in 2041; plain formulas at most 99 characters.
  - In LibreOffice (no LAMBDA support) errors only in 2041 (5,184 cells); price and subsector rows 2022-2040 identical to the LAMBDA host; power within 0.01%.
  - Values (LAMBDA year expanded) equal the Python recomputation with moving-average fuel prices and the window to 2041, in every scenario test.
  - Regression vs v2.2: 3,854 shared codes, 0 outside the intended power changes.
  - Macro batch run and CheckBatchRun in LibreOffice match the stored results (1.7e-14).
- **Caveats:**
  - **2041 needs Excel 365.** Elsewhere 2041 shows errors (Mitigation and the 2041 column of MTOutputs). The power look-ahead there uses last year's values for 2041, so power results 2022-2040 differ slightly from Excel 365 (size in the check report). Stored results are the Excel 365 values.
  - **Results changed against v2.2** for power and everything that includes it (fuel use, CO2 and revenue totals, electricity prices, the ETS baseline). Price and subsector rows are unchanged.
  - **Not run in Excel yet** (I1).
