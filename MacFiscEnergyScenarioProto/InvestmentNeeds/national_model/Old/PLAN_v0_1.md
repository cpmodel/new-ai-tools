# National Mitigation Investment Model (NMIM) v0.1 — build plan

A simplified national-level model of net zero mitigation investment for every country
with IMF CPAT data (~186), annual 2024-2050, built to CPAT norms so that its codes,
layout, data and equations line up with CPAT-AI-Mitigation-MVP and the CBAM Policy
Tool. It replaces the v1.1 workbook's allocation of group totals (Tier 3) with a
country stock-and-flow calculation (Tier 2).

## Deliverables

| File | What it is |
|---|---|
| `extract_data_v0_1.py` | Reads repo sources, writes `data/*.csv` in CPAT code conventions |
| `nmim_v0_1.py` | Python engine: all countries, baseline and net zero, 2024-2050; results in long format with CPAT-style codes |
| `build_workbook_v0_1.py` | Writes `NMIM_v0.1.xlsx` in CPAT layout: one country at a time in live formulas (country selector), all-country results as values, Check sheet Excel = Python |
| `check_v0_1.py` | LibreOffice recalculation; Excel vs Python for several countries; identities |
| `NMIM_Method_v0_1.md` | Methodology (current state only, NORMS s7) |
| `VersionNotes_NMIM.md`, `CAVEATS.md` | Version notes and the task log (NORMS s6) |

## CPAT norms adopted

- **Codes** `mit.<indicator>.<fuel>.<sector>`; fuels `coa, nga, gso, die, lpg, ker, oop, ecy`; sectors `pow, rod, res, ind`; results `<iso>.mit.<indicator>.<item>.<scenario>` with scenario 1 = baseline, 2 = net zero.
- **Mitigation layout**: A section, B Fuel, C Sector, D Description, E Unit, F Source, G Input Code, H Output Code, I Note, J Helper, K = 2021 ... N = 2024 (base data year, tan), O = 2025 ... AN = 2050 (extends the norm's 2045 horizon to 2050 for net zero).
- **Colours**: title band 00B050, section band 92D050, inputs EBF1DE, codes DCE6F1, base year DDD9C4, results to review FFF2CC; Arial.
- **Inputs** in MTInputs column structure (B Parameters, C unit, D/E hidden placeholders, F Used for calculation, G Defaults, H NameOfParameter), read by name with INDEX/MATCH. (Row structure is new: these parameters are not in legacy MTInputs.)
- **Data**: CPAT's own data (IMF fossil fuel subsidies / CPAT 2024 panel) for energy use, prices, emission and pollution factors; CPAT income-group elasticities (inc, aei) from the mitigation MVP; CPAT PowerTech parameters as an option.
- Short formulas (<= 100 characters where practical), version log in ReadMe, `Old/` for superseded versions, CAVEATS entry, separate method note and version notes.

## Consistency with the other models

- **CPAT activity projection**: baseline end-use energy follows CPAT's equation with prices held at 2024 real levels: `E_t = E_{t-1} (1 + g_t)^inc / (1 + aei)`, income-group elasticities from CPAT.
- **CBAM Policy Tool**: industrial output of the CBAM goods (steel, clinker, ammonia, aluminium), their growth and emission factors come from the CBAM tool's Country data sheet; process emissions of those goods are abated in the industry sector.
- **v1.1 workbook**: same country WACC (observed, Damodaran, IHLEG default), same three lenses, same co-benefit basis (CPAT air pollution costs), same groups.

## Equations (both Python and Excel)

1. GDP growth: WEO 2025-26, then linear convergence to a long-run rate by income group by 2040.
2. Baseline end use (PJ), rows rod.gso, rod.die, res.coa, res.nga, res.lpg, res.ker, res.ecy, ind.coa, ind.nga, ind.oil, ind.ecy: CPAT equation above.
3. Net zero end use: road EV sales share to 100% by group target year, stock turnover over 12 years; building fossil heat electrified linearly to 100% by target year; industrial fossil use abated linearly to a group share by 2050 (part electrified). Electricity added with efficiency ratios (EV 0.25, heat pump 0.33, industry 0.8).
4. Power: generation = end-use electricity x (generation / consumption, base year); clean firm generation held at 2024; fossil share falls on trend (baseline) or linearly to 0 by the group target year (net zero); VRE fills the rest. Capacity = generation / capacity factor; stock turnover with lives; additions x unit cost (learning) x local price level^phi.
5. Grids and storage: grid investment = gamma x generation investment; storage hours per kW of VRE rising with the VRE share.
6. Transport, buildings, industry investment: EV premium (falling to parity) + charger per EV sold; capital per GJ a year of fossil heat or industrial fuel displaced; CBAM process emissions at a CCS capital cost per tonne.
7. Fuels supply and AFOLU: abatement to targets (methane -75% by 2030; net land-use emissions 0 by 2030; agricultural non-CO2 -30%) at cost per tonne (annual spending).
8. Fuel bills, CO2 and air pollution: fossil use x CPAT supply cost, IPCC/CPAT emission factors, CPAT air-pollution cost per unit.
9. Incremental cost by year and lens: sum over vintages of CRF(r, life) x (net zero - baseline investment) + change in O&M - fuel savings (- air pollution in lens c); lenses market WACC, 6%, 6% + co-benefits; PV 2025-2050.

## Validation

- 2030 world and EMDC clean-energy investment vs IHLEG (USD 4.8 tn / 1.6 tn) and v1.1.
- Baseline 2030 emissions vs CPAT; net zero energy CO2 in 2050 near 0.
- Excel (selected country, live formulas) = Python to rounding, for 5+ countries.
- Identities: generation balance, capacity balance, investment >= 0.

## Out of scope for v0.1

Price responses (no carbon pricing; prices fixed real), optimisation of the power mix, aviation and shipping, hydrogen trade, macro feedback, Damodaran (still not loaded), uncertainty runs.
