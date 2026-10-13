# Power generation costs and power prices: method and pseudocode (CPAT-AI-Mitigation-MVP)

**Note (v1.06).** The quantities are now modelled (`PowerModel_Method_v0_1.md`): the interim data rows of sections 5-6 below are replaced by the engineer model, and three corrections are marked below (VAT on the policy term, long-term storage allocation, VRE share).

**Scope.** This note covers the price side of the power sector:
- what each generation type costs per kWh;
- what the generation mix costs on average;
- the end-user electricity prices for residential and industrial users;
- the power-sector carbon cost and its revenue.

It follows legacy CPAT section 4 ("Technoeconomic ('engineer') power model", Mitigation rows 2954-3672 and 3867-3927) and documentation chapter 3, sections 3.4.1 and 3.4.2.5, keeping legacy's sub-table letters (A Inputs, B Amortised costs, C Levelised costs, D Variable costs, E Total costs, F Memos, G End-user prices, J Storage). Quantities (dispatch, capacity, investment: legacy tables H, I, K, L) are the next bucket. Until it exists, the two quantities this block needs from it are **interim data rows** (section 6).

Units: real US dollars of the results year (2026) per kWh, unless stated. Generation types *f*: coal, natural gas, oil, nuclear, wind, solar, hydro, other renewables, biomass (legacy order).

## 1. Interactions

The price block is mostly a one-way chain. Six links connect it with the rest of the model; the table lists each one and how legacy keeps it free of circular references.

| # | From | To | What flows | How the loop is broken (legacy) | MVP now |
|---|---|---|---|---|---|
| I1 | Section 2 prices (coal and gas for power, oil, biomass) | D Variable costs | Fuel price before new policies ($/GJ, $/bbl) | One way | Live link |
| I2 | Section 1 policies (carbon price, ETS power coverage, fuel price reform) | D Carbon cost; G end-user prices; revenues | Carbon cost per kWh by generation type | One way | Live link |
| I3 | Generation mix (dispatch and investment, engineer model) | E average generation cost; J storage (via the VRE share) | Generation shares by type | Dispatch uses current variable costs, which do not depend on the average cost; prices use last year's average cost (I5) | Interim data (legacy baseline run) |
| I4 | Investment (engineer model) | B weighted average capex and storage cost of vintages | New investment as a share of remaining capacity Φ | B uses last year's Φ (verified: exact match with last year's Φ, see 3.2) | Interim data (legacy baseline run) |
| I5 | E average generation cost | G end-user prices | Supply cost = average generation cost + transmission | **Deliberately lagged one year**: the price in *t* uses the cost of *t*−1 | Implemented |
| I6 | G end-user prices | Electricity demand (sector fuel use) → generation → I3 | Residential and industrial electricity prices | Through I5's lag | Electricity is not yet a fuel in the sector blocks: prices are computed and reported, demand is interim data |
| I7 | J storage (systems cost of renewables) | B (vintage storage cost) and C (levelised, for investment) | Marginal storage cost of wind, solar, other renewables | Uses the current VRE share from dispatch | Live link from the interim shares |
| I8 | C levelised costs, D9 levelised variable costs | Investment logit (engineer model); F7 switching carbon prices | Levelised total investment cost by type | Investment only: they do not enter end-user prices | Levelised fixed cost built; levelised variable cost and the investment cost wait for the engineer model |

Two consequences shape the design:
1. **End-user prices depend on generation costs only through the weighted average** (E) and the transmission add-on. Variable, amortised and storage costs matter for prices only as far as they enter that average.
2. **The renewables systems cost is an interaction inside the price block.** The VRE share (from the generation mix) sets the marginal storage cost (J). This enters the cost of wind and solar in two ways:
   - **amortised**, through vintages: the cost of the existing stock (B), which goes into end-user prices;
   - **levelised:** the cost of new plants (C), which goes into investment decisions.

   So a scenario that adds renewables raises their storage cost per kWh, which feeds back into the average generation cost a year later.

## 2. Simplifications against legacy

Each simplification gives the same numbers as legacy (checked in the Python reference, section 7) unless the table says otherwise.

| # | Legacy | MVP | Same numbers? |
|---|---|---|---|
| S1 | Five amortised-cost tables (B3 capital, B4 interest, B5 decommissioning, B6 storage, B7 fixed O&M) summed in B8 | One row per type: fix = (cax_av × (1/life + wacc/2) + dtc/life) / (cf × 8,760) + sto_av + opf | Yes (algebraically identical) |
| S2 | Six levelised tables (C1-C6) summed in C7; discount-factor tables A5, A6 | One row per type: lfx = (cax + dtc + tcx) / (cf × 8,760 × dlf) + msc + opf + rns, with dlf in closed form (1 − (1+w)^−N) / (1 − (1+w)^−1) as a parameter | Yes |
| S3 | Duplicates: A16 = B7 = C4 (fixed O&M), D7 = D8 (variable cost), E1 = E4 (investment cost), F2/F3 (fuel prices and fuel cost), E2 (cost before storage), E5 (breakdown memo) | Dropped; each quantity once | n/a |
| S4 | Residential and industrial price: supply cost + producer subsidy + other tax + "over/under-estimated costs" residual, the last three held at their last-historical-year values, then VAT | Telescoped: rp_pre(t) = rp_pre(t−1) + pass × (sc(t) − sc(t−1)) × (1 + VAT) after the last historical year; data before. The three held components cancel out of the change | Yes (pass-through 1, VAT constant after the last historical year, as legacy) |
| S5 | Fuel prices for power: 5-year moving average (A18; spot prices are a dashboard option) | Spot prices from section 2 (the legacy option "use spot fuel prices") | **No**: coal and oil cost changes reach power one to four years earlier. Listed in LegacyDiff; a moving average can come later |
| S6 | Coal implicit cost (D3) and the additional shadow price (D6) | Not built (both 0 in the Egypt run, no coal in Egypt's mix) | Yes for Egypt |
| S7 | Producer-side subsidies and the over/under-estimate residual as rows | Not needed for prices (S4). Reported as one row, the subsidy gap = supply cost − price before VAT, for revenue analysis | Yes for prices |

Row count of the price chain for one scenario: legacy about 330 rows (A5-G4, J1-J3); MVP about 140 (section 4).

## 3. Method by legacy sub-table

### 3.1 A. Inputs (data step: sheet Inputs_power, one row per generation type, and a few year paths)

| Symbol | Meaning | Legacy |
|---|---|---|
| cax0, N | Capital cost at capex time factor 1 ($/kW; legacy base-year cost ÷ base-year factor), lifetime (years) | A2, B2 |
| tcf(t) | Capex time factor (renewables fall; others 1) | A7 |
| cax(t) = cax0 × tcf(t) | Capital cost path | A13 |
| wacc | Country WACC + technology premium | A3, B4 "Final WACC" |
| cf | Capacity factor (constant) | A8 |
| ν | Thermal efficiency, NCV (constant) | A9 |
| dtc, tcx | Decommissioning cost; levelised transmission capex ($/kW) | A14, A15 |
| opf | Fixed O&M ($/kWh) | A16 |
| vop | Variable O&M ($/kWh); nuclear fuel 0.004 $/kWh included as fuel | D1, D2 |
| rns | Renewable subsidy ($/kWh, negative = subsidy) | A17 (0) |
| dlf | Discounted lifetime (1 − (1+wacc)^−N) / (1 − (1+wacc)^−1) | A6 |

### 3.2 B. Amortised (cost-recovery) fixed costs, for the cost of the existing stock

- Weighted average capex of the stock (verified against legacy, exact):
  cax_av(t) = cax_av(t−1) × (1 − Φ(t−1)) + cax(t−1) × Φ(t−1); cax_av(base) = cax(base).
- Weighted average storage cost of vintages (exact; note the current-year marginal cost):
  sto_av(t) = sto_av(t−1) × (1 − Φ(t−1)) + msc(t) × Φ(t−1); sto_av(base) = msc(base).
- Amortised fixed cost (S1): fix(t) = (cax_av × (1/N + wacc/2) + dtc/N) / (cf × 8,760) + sto_av + opf.

### 3.3 C. Levelised (forward-looking) fixed costs, for new plants (investment)

lfx(t) = (cax(t) + dtc + tcx) / (cf × 8,760 × dlf) + msc(t) + opf + rns (S2).

### 3.4 D. Current variable costs

- Fuel cost per kWh: fc(t) = pf(t) × 0.0036 / ν, with pf the section-2 price before new policies in $/GJ (oil: $/bbl ÷ GJ per bbl). Nuclear: 0.004 $/kWh. Renewables 0.
- Variable cost before carbon: vbc(t) = vop + fc(t).
- Carbon cost per kWh generated: ccp(t) = nce_pow(f, t) × 0.0036 / ν. nce_pow is the new policy wedge on the fuel in the power sector ($/GJ):
  - carbon tax: carbon price × EF × fuel coverage × power coverage × (1 − ETS coverage of power);
  - ETS: ETS tax-equivalent price × EF × ETS coverage of power × (1 − OBR share of power);
  - fuel price reform for coal and gas in power.
- Total variable cost: tvc(t) = vbc(t) + ccp(t). This is used by dispatch (engineer model).

### 3.5 E. Total costs

- Generation cost before carbon (cost recovery): gnc(t) = fix(t) + vbc(t) (legacy E3).
- Average generation cost: gnc_av(t) = Σ_f gns(f, t) × gnc(f, t).
- Average carbon cost per kWh consumed: ccp_av(t) = Σ_f gns(f, t) × ccp(f, t) × gen/cons. Gen/cons is the ratio of generation to final consumption: losses, own use and net exports, from the base-year balance.

Carbon is kept out of gnc_av and added separately (legacy, documentation 3.4.1.7), so that a rebate can be computed against it.

### 3.6 F. Memos

F7 switching carbon prices, F1 system cost and F4 carbon intensity are memos. Only F4 (tCO2 per kWh = EF × 0.0036/ν) is built, as a parameter.

### 3.7 J. Systems integration cost of renewables (storage)

- VRE share: v(t) = gns(wind) + gns(solar) + gns(other renewables) (legacy row 3851; v1.05 omitted other renewables, 0 for Egypt).
- Short-term storage (batteries):
  - marginal hours: h(t) = ast × 18 × v(t), where ast = 1 is the share allocated to VRE and 18 = d(9 v²)/dv;
  - cost per hour per kWh: cph(t) = (cbat(t) + cint(t) / r) / (dlf_ren × 8,760) + obat(t) / 8,760, where r = 2 hours kWh/kW, and cbat, cint, obat are the battery capex, interface capex and opex paths.
- Long-term storage (electrolysers): marginal units m(t) = 2 × max(0, v − x) / (1 − x)², with x = 0.75. Cost per unit: cel(t) = (electrolyser capex + LT storage capex) / (dlf_ren × 8,760) + fixed opex / 8,760 + variable opex.
- Marginal storage cost of wind, solar and other renewables: msc(t) = h(t) × cph(t) + m(t) × cel(t). It is 0 for other types.

Legacy has a 33% "allocation of LT storage to VRE" parameter and applies it twice (rows 3911 and 3912, a bug). From v1.06 the MVP applies it once (v1.05 left it out).

### 3.8 G. End-user prices (residential g = res, industrial g = ind)

- Supply cost (I5, lagged): sc(g, t) = gnc_av(t−1) + tmc(g), with tmc = 0.040 (res) and 0.015 (ind) $/kWh. In the base year, gnc_av of the base year.
- Price before new policies (S4):
  - up to the last historical year L: rp_pre(g, t) = hrp(g, t), the historical retail price including VAT from the price dataset (nominal → real with the CPI index);
  - after L: rp_pre(g, t) = rp_pre(g, t−1) + pass × (sc(g, t) − sc(g, t−1)) × (1 + VAT(g, L)), with pass = 1 (legacy hardcode).
- End-user price: rp(g, t) = rp_pre(g, t) + (ccp_av(t) − reb(g, t) + pex(g, t)) × pass × (1 + VAT). The rebate (output-based rebating) and the power excise are 0 for now. As in legacy (row 3614), VAT is charged on the new policy terms (corrected in v1.06; v1.05 left VAT out, no effect for Egypt with VAT 0).
- Subsidy gap (S7): sgap(g, t) = sc(g, t) − rp_pre(g, t) / (1 + VAT). Positive = price below modelled supply cost.

### 3.9 G4. Revenues (USD million real)

- VAT: rp_pre × VAT / (1 + VAT) × cons(g).
- Carbon revenue from power: ccp_av × cons(g). This equals the carbon cost of generation, because ccp_av is per kWh consumed.
- Subsidy gap: sgap × cons(g).

Here cons(g) is consumption in GWh, interim data until electricity demand is modelled (I6). Power revenues are reported separately, not yet in the section-12 totals.

## 4. Pseudocode (one scenario, years base..2040, in order)

```
inputs: tech table T[f] (cax0, N, wacc, cf, nu, dtc, tcx, opf, vop, rns, dlf, k_fuel, price_pos, ef_pow)
        year paths: tcf[f,t], cbat[t], cint[t], obat[t], cel_parts[t]
        interim (engineer model later): gns[f,t], phi[f,t], cons[g,t], gen_per_cons
        section 2: rpb[price_pos,t]  (coal/gas for power, oil, biomass; before new policies)
        section 1: cptraj[t], ets_pe[t], cover/ETS/OBR switches for power, fpr[coa.pow|nga.pow,t]
        data: hrp[g,t], vat[g,t] (t <= L)

for t in years:
    # J storage (I3 -> I7)
    v      = gns[wnd,t] + gns[sol,t] + gns[ore,t]
    cph    = (cbat[t] + cint[t]/2) / (dlf_ren*8760) + obat[t]/8760
    m_lt   = 2*max(0, v-0.75)/(0.25**2)
    msc    = 18*v*cph + m_lt*cel[t]*0.33                      # wind, solar, other renewables (v1.06: LT allocation)
    for f in types:
        cax[f]      = cax0[f]*tcf[f,t]
        # B amortised (vintages lagged one year, I4)
        if t == base: cax_av[f] = cax[f];  sto_av[f] = msc_f
        else:         cax_av[f] = cax_av_prev[f]*(1-phi[f,t-1]) + cax_prev[f]*phi[f,t-1]
                      sto_av[f] = sto_av_prev[f]*(1-phi[f,t-1]) + msc_f*phi[f,t-1]
        fix[f]  = (cax_av[f]*(1/N[f] + wacc[f]/2) + dtc[f]/N[f]) / (cf[f]*8760) + sto_av[f] + opf[f]
        # C levelised (investment only, I8)
        lfx[f]  = (cax[f] + dtc[f] + tcx[f]) / (cf[f]*8760*dlf[f]) + msc_f + opf[f] + rns[f]
        # D variable (I1, I2)
        vbc[f]  = vop[f] + rpb[price_pos[f],t]*k_fuel[f]          # k_fuel = 0.0036/nu (oil: /GJ per bbl)
        ccp[f]  = nce_pow[f,t]*0.0036/nu[f]
        # E
        gnc[f]  = fix[f] + vbc[f]
    gnc_av = sum_f gns[f,t]*gnc[f]
    ccp_av = sum_f gns[f,t]*ccp[f] * gen_per_cons
    for g in (res, ind):                                          # G (I5: lag)
        sc[g]     = (gnc_av if t == base else gnc_av_prev) + tmc[g]
        rp_pre[g] = hrp[g,t] if t <= L else rp_pre_prev[g] + pass*(sc[g]-sc_prev[g])*(1+vat[g,L])   # hrp incl. VAT
        rp[g]     = rp_pre[g] + (ccp_av - reb[g] + pex[g]) * pass * (1 + vat[g])
        sgap[g]   = sc[g] - rp_pre[g]/(1+vat_used[g])
        rev_vat[g], rev_carbon[g], rev_gap[g] = ... * cons[g,t]
```

## 5. Workbook layout (Mitigation section 3, as legacy section 4; built in v1.05)

The power block sits in section 3, "Power sector: generation costs and electricity prices", per scenario column group, with one row per generation type (or end-user group) for each variable. Each block has one formula. Type-specific constants sit in the hidden parameter columns D:G, looked up in the data step sheet Inputs_power, which derives them from the data sheets PowerTech, PowerPaths and PowerParams. The right column has the same plain formula as the other years (no LAMBDA needed). Sub-headings carry the legacy letters, in legacy order (J last, as in legacy, although B and C use it):

- A: tcf (capex factor), gns (generation shares; interim), phi (investment share; interim), storage cost paths cbat, cint, obat, cel.
- B: cax, caxav, stoav, fix.
- C: lfx.
- D: vbc, ccp. The new policy wedge is not a separate row: ccp is computed per kWh in one step with the constants tCO2/kWh (EF × 0.0036/ν) and kWh factor of the fuel price reform (0.0036/ν ÷ GJ per price unit); both are 0 for types without a fuel.
- E: gnc, gncav, ccpav.
- G: cons (interim), sc, rppre, rp, sgap by user group; revenues rvat, rcarb, rgap.
- J: vre, cph, msc.

Interim rows (gns, phi, cons) are shown in orange italic, with a note naming the engineer model as their future source. When rolled up, section 3 shows the average generation cost and the two end-user prices. MTOutputs adds these three indicators.

Optional named LAMBDAs of section 3 (sheet `Lambdas` gives the formula for each row): CAPEX, VINTAGE (B2 and B6), AMORTISED, LEVELISED, VARCOST, GENCOST, WEIGHTED (averages and the VRE share), STORAGEHOUR, STORAGECOST, POWERSUPPLY, POWERPRICE, ENDUSERPOWER, SUBSIDYGAP, POWERREV; ccp reuses NEWPOLICY, CARBONTAX and ETSCOST of the sector sections.

## 6. Interim data until the engineer model (I3, I4, I6)

| Rows | Source now | Source later |
|---|---|---|
| gns: generation shares by type, 2022-2040 | Legacy cached baseline (Mitigation I6 / E0) | Engineer model dispatch and investment |
| phi: investment as a share of remaining capacity | Legacy cached baseline (B1) | Engineer model investment |
| cons: residential and industrial consumption (GWh); generation / consumption | Legacy cached baseline (G1, G2, H2) | Electricity demand in the sector blocks plus the balance shares |

Because these are baseline values, a policy scenario's power prices respond to fuel and carbon costs but not to its own generation mix until the engineer model is built. This matters for high carbon prices; it is flagged in CAVEATS and LegacyDiff.

## 7. Validation

1. **Method = legacy.** A Python reference feeds the legacy inputs (its fuel prices, shares, Φ, prices and VAT) into this method and reproduces the cached legacy rows: B2, B6, B8, C7, D4, E3, the weighted average, J storage, and G1/G2 supply cost and prices. The pass mark is a relative error below 1e-6.
2. **Workbook = Python** for the MVP's own inputs (corrected Egypt prices, spot fuel prices), in the same style as the other checks.
3. **Formula uniformity**, the optional LAMBDA library, and copy of scenario groups, as for the rest of the sheet.
4. Regression: every existing output code is unchanged. The power block does not yet feed any other section.

Result in v1.05 (`check_report_v1.05.md`, PASS): (1) 80 legacy rows reproduced, largest relative difference 5.6e-17; (2) all 28 power variables equal the Python reference in every scenario test, largest relative difference about 5e-15; (3) one formula per block, LAMBDA in every right-column cell; (4) 3,642 codes unchanged vs v1.04.

## 8. Open decisions

1. Spot vs moving-average fuel prices (S5).
2. VAT on power policy terms: legacy charges VAT (row 3614), as the MVP does for fuels; applied from v1.06.
3. Historical electricity prices: the corrected Egypt block (residential 0.054, 0.048, 0.043 $/kWh nominal, 2022-2024) differs from legacy's vintage (0.097 in 2022). The MVP uses the corrected block.
4. Producer-side subsidy data: the corrected block has none for electricity (`mit.pros.ecy.all` = 0); legacy used −9.2 bn USD in 2022. The MVP reports the modelled subsidy gap instead (S7).
