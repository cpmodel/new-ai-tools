# Power engineer model: method (CPAT-AI-Mitigation-MVP v1.09)

**Scope.** This note covers the quantity side of the power sector:
- power demand (legacy table H);
- dispatch (I);
- capacity and retirement (K);
- investment (L);
- the cost rows that investment needs: the shadow price and rebate (D6), the levelised variable cost (D9) and the levelised investment cost (E1).

It completes the price side (`PowerPrices_Method_v0_3.md`). The legacy reference is Mitigation section 4 of `CPAT 1.0pre_456_NoPropData.xlsb`: rows 2954-4297 for the baseline, rows 9285-10628 for the carbon-tax run.

**Alignment.** The MVP follows legacy formula by formula, except for ten bugs, which are fixed. They are listed with their effects in `PowerModel_BugsAndImprovements_v1_0.docx`. `power_model_ref_v0_2.py` implements the method in Python. It has one switch per bug (`FIXES`), and with all switches off it reproduces 208 cached legacy rows of both runs to 1e-13. Planned retirement is available for every generation type; legacy offers it for coal only.

**Power demand and sector policies.** Power demand follows its own elasticities (H). Sector policies (carbon tax, ETS, feebates, fuel price reform on fuels) act on the fuel blocks only, not on power demand. The sector blocks and section 11 show electricity demand as memo lines; these lines are not in the fuel totals.

Units: generation and capacity in MW / MWy; demand in ktoe; costs in real $/kWh of 2026.

## 1. Order within a year and across years

```
year t:  H  demand(t)        <- demand(t-1), GDP(t), price(t-1)              (no circularity: price lagged)
         K1 capacity(t)      <- capacity(t-1) - retirement(t-1) + investment(t-1)
         I  dispatch(t)      <- capacity(t), required generation(t), variable costs(t)
            shares(t)        -> J VRE share -> storage cost(t) -> B (vintages), C (levelised)
            shares(t)        -> E average generation cost(t) -> G supply cost(t+1) -> price(t+1) -> demand(t+2)
         K  retirement(t)    <- capacity(t), early coal retirement (coal variable cost vs wind and solar costs)
         L  investment(t)    <- required generation(t) - capacity after retirement(t); logit on E1(t)
                             -> capacity(t+1); share in the stock phi(t) -> B vintages(t+1)
```

The levelised variable cost D9(t) averages the variable cost from t to the end of the horizon. Variable costs depend only on prices and policies, not on the generation mix, so D9 is computed before the year loop. In the workbook it runs backwards: next year's value enters this year's formula.

## 2. Equations (with fixes)

**H Demand.** There are four demand groups g (transport, residential, industry, services), with electricity elasticities from sheet Elasticities: eps_Y, eps_U, eps_F, alpha.
- p_g(t) = w_g × p_res(t-1) + (1 - w_g) × p_ind(t-1), with w = 0.5 for transport, 1 for residential, 0 for industry and services.
- mult_g(t) = (1 / (1 + alpha + extra))^(1 + eps_U) × (1 + GDP growth)^(eps_Y × inc(t)) × (p_g(t) / p_g(t-1))^(eps_U + eps_F + eps_U eps_F) × (1 + shock × switch).
- dem_s(t) = dem_s(t-1) × mult_group(s)(t), for the 15 subsectors (services use the services multiple in every year).
- req(t) = Σ dem × (1 + net exports + transmission + own use + stat. diff. shares) × 11.63 × 1000 / 8760 (MWy). In the base year, req is actual generation.
- Memo lines:
  - demand by demand group = Σ dem over the subsectors of the group;
  - electricity demand by subsector (`eldem`) in each sector block of sections 5-8, and by sector (transport, buildings, industry, other energy use) and in total in section 11.

**I Dispatch.**
- Fixed-output types (wind, solar, hydro, other renewables, oil, nuclear, biomass): gen = capacity × capacity factor.
- Coal and gas:
  - reqcg = MAX(req - fixed output, 0);
  - each type's minimum = MIN(MAX(reqcg - maximum of the other, 0), own maximum), where maximum = capacity × 0.9;
  - the remainder is shared by a logit EXP(-2 × tvc / lowest tvc);
  - gen = MIN(maximum, minimum + remainder × share).
- Generation share = gen / total.
- Fuel use = gen × 8.76 / 11.63 / efficiency (ktoe); CO2 = fuel use × PJ/ktoe × EF (power).

**K Capacity and retirement.**
- Planned retirement: each type has a switch (PowerTech `planned_retirement`, Yes/No) and a data row of planned retirements in nameplate MW per year (PowerPaths `pret|<type>`).
- Retirement = MIN(planned retirement if the switch is Yes, else capacity × 1/lifetime; + early coal retirement for coal; capacity). Hydro has rate 0.
- For Egypt, the legacy file has no planned-retirement data, so every switch is No.
- Effective retirement = retirement × capacity factor.
- Early coal retirement:
  - room = MAX(0, maximum wind and solar additions - last year's increase in generation - last year's non-coal effective retirements);
  - early = MAX(0, MIN(0.8 × coal effective capacity, room) × (1 - coal share of a logit over coal variable cost, wind and solar investment cost)) / coal capacity factor × 0.8.
- Capacity needed = MAX(req(t) - (effective capacity - effective retirement), 0). The capacity bought in year t is available in year t+1.

**L Investment.**
- First round:
  - permitted types only;
  - logit share = EXP(-2 × E1 / lowest permitted E1) / sum;
  - first round = MIN(scale-up limit, need × share), with scale-up limit = rate × total effective capacity.
- Second round: deficit = need - Σ first round; second = MAX(0, MIN(own remaining space, deficit - space of cheaper types)).
- Nameplate investment = effective / capacity factor.
- USD = nameplate × capital cost of the year / 1000.
- phi = nameplate / (capacity + nameplate - retirement), which enters the vintage averages B2 and B6.

**D5, D6, D9, E1.**
- Section 1 holds three policy rates by sector, in $/tCO2:
  - carbon tax = carbon price × sector coverage × (1 - ETS coverage);
  - ETS cost = ETS tax-equivalent price × ETS coverage × (1 - OBR share);
  - shadow price = feebate (and regulation) rate × share impacting efficiency + ETS tax-equivalent price × ETS coverage × OBR share.
- The sector blocks and the power block use the same rates.
- New policy cost per kWh (D5) = tCO2/kWh × (carbon tax × fuel coverage + ETS cost) + fuel price reform × its kWh factor.
- Shadow price per kWh = shadow price of power × tCO2/kWh.
- Rebate = -Σ shadow × base-year share, per kWh generated.
- tvc = vbc + ccp + shadow + rebate.
- D9 = discounted average of tvc to 2040.
- E1 = MAX(lfx + D9, 0.01).

## 3. Workbook

- Mitigation section 3 keeps the legacy letters in the legacy order: A, B, C, D, E, G, H, I, J, K, L, plus M (v1.09: revenues by generation type, `PowerPrices_Method_v0_3.md` section 3.9).
- Each block has one plain formula across years, types and scenarios, the right column included; no LAMBDA support is needed. The named LAMBDAs are an optional library (sheet `Lambdas`); data rows, sums and the look-ahead rows `dsum` and `lvc` have none.
- No cell formula is longer than 100 characters. A formula that would be longer is split into visible steps, for example `ecmin` and `ecshr` before `early`; the optional LAMBDAs may hold longer logic.
- Constants by type and the demand tables are in Inputs_power; the data sheets are PowerTech, PowerPaths, PowerParams and PowerDemand (`extract_power_v0_3.py`).
- Power is part of the main thread (v1.09):
  - CO2 enters section 13: the total, by fuel and as ETS sector power. A new ETS that covers power therefore sees power emissions, and the baseline CO2 data rows include them.
  - Generation fuel use of coal, gas, oil and biomass (I, `pe`) enters section 11: row `ener.pow` (fuel flag x `pe`, flag from Inputs_power) is part of the total, and each fuel row adds the power use of that fuel. Electricity final demand stays a memo (`eldem`), so nothing is counted twice. Nuclear, wind, solar, hydro and other renewables have no fuel row.
  - Revenues enter section 12 (rows `.pow`): fuel taxes and subsidies on generation, the change in existing carbon pricing, electricity VAT and subsidy gap, new policies with the ETS at the permit price.
  - Existing carbon pricing changes the variable cost before new policies (`xck` in D, `PriceProjection_Method_v0.6.md` section 5b).
  - Check rows `ener.chk`, `rev.chk` and `co2.chk` are 0 when the totals equal their sector, power and fuel rows.
- Settings C19 switches the legacy 2023-2024 demand shock on (default 0, as the fuel blocks).

## 4. Validation (`check_report_v1.09.md`)

1. **Method = legacy.** The Python reference, fed with the cached legacy inputs and with the bugs switched on, reproduces 208 legacy rows of the baseline and of the carbon-tax run.
2. **Workbook = Python.** With all bugs fixed and the MVP inputs, the workbook matches the Python reference for every power row, in every scenario test.
3. **Properties:**
   - generation = required generation;
   - the capacity balance holds;
   - investment is at most the capacity needed (scale-up limits can leave part of it unplaced);
   - the supply cost lag holds;
   - a carbon price raises the VRE share and lowers power CO2.
4. **Planned retirement.** With the switch on for gas and oil and test data, the workbook matches the Python reference, and retirement equals the data.
5. **Memo lines.** The memo lines equal power demand by subsector and add up by sector and in total.
6. **Main thread (v1.09).** Section 11, 12 and 13 totals equal the Python reference (fuel use and revenues with power); with non-zero existing carbon pricing test data, power `xck`, the revenues by type and the totals match too.
