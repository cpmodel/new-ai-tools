# InvestmentNeeds — Mitigation (Net Zero) Investment Needs Model

`Mitigation_Investment_Needs_v1.0.xlsx` estimates 2030 **total** and **incremental**
mitigation investment needs, and their **incremental cost**, for 195 economies
(every economy with IMF WEO GDP) and seven sectors, under three cost-of-capital
lenses:

| Lens | Rate | Benefits counted |
|---|---|---|
| (a) Market | Country WACC (high case also reported) | Fuel and operating savings |
| (b) World Bank | 6% (Fay et al. 2016); IBRD ~5.2% by switch | Fuel and operating savings |
| (c) World Bank + local co-benefits | 6% | Savings + local air-pollution co-benefits (memo: + carbon at the WB shadow price) |

The method is the one set out in *Global Investment Needs for Climate Action: A Method
Based on SISEPUEDE* (October 2026), equations 1–7. The workbook implements its Tier 3
(group totals allocated to countries by drivers), with the cost of capital, savings and
co-benefits differentiated by country.

## Sectors

Power generation · grids and storage · transport · buildings · industry · low-emission
fuels and fuel supply · AFOLU and nature.

## How it works

1. **Pools** (USD bn/yr, 2030): clean energy USD 4.8 tn world, 1.6 tn EMDCs ex-China
   (IHLEG 2024), China 0.85 tn (derived from IEA-IFC 2023), advanced economies the
   residual; AFOLU from IHLEG (EMDCs) and UNEP SFN (world NbS). Clean energy is split by
   sector with the IEA-IFC (2023) shares.
2. **Allocation** to countries: `w = [β·driver share + (1−β)·PPP GDP share] × price level^φ`,
   drivers from IMF CPAT 2024 (electricity demand, road fuel, building and industrial
   energy, fossil supply, non-CO2 + LULUCF emissions).
3. **Incremental** = total × incremental share (AE/China 0.38, McKinsey 2022; EMDC 0.24,
   CCDR vs IHLEG).
4. **Capital charge** = CRF(r, asset life) × investment.
5. **Country WACC** (lens a): observed RE WACC from the dashboard (IRENA 2021, Steffen et
   al. 2025) → Damodaran-based WACC (eq. 5, calibrated to the observed level) → IHLEG
   group rate.
6. **Savings and co-benefits** per dollar (b = 0.133, c = 0.129, SISEPUEDE LAC
   calibration) scaled by country indices from CPAT: fossil fuel bill per tCO2 and local
   air-pollution cost per tCO2.
7. **Incremental cost** = capital charge on incremental investment − savings
   (− co-benefits in lens c). Positive = net cost.
8. **PIMA**: PIMA-adjusted investment uses the public-investment efficiency gap (PEFA
   PI-11 where available, else IMF 2015 income-group gaps), with a reform switch.

## Indicators taken from TaxMacroFiscalDashboard

Cost of capital (Finance 1 block): IRENA WACC, Steffen WACC, evidence strength, IEA
Observatory flag, IDS official-private spread, private creditor rate, maturity.
PIMA block: BOOST capex execution, PEFA PI-11 and PI-12.2, UNDP climate-PFM flag.
Also gross debt, interest-to-revenue, mitigation finance received, PPI.

## Damodaran

Damodaran's country risk data extend WACC coverage beyond the 84 economies with an
observed WACC. Either:

- put `ctryprem.xlsx` (from pages.stern.nyu.edu/~adamodar) in `InvestmentNeeds/raw_data/`
  and re-run the build script; or
- paste it into the `Damodaran` sheet (instructions on the sheet).

**v1.0 ships without Damodaran values**: the source host could not be reached when the
workbook was built, so the 111 economies without an observed WACC use the IHLEG group
rate.

## Build

```bash
pip install pandas numpy openpyxl
python InvestmentNeeds/build/build_investment_model.py
```

It reads only repository files: `TaxMacroFiscalDashboard/build/output/global_data_wide.csv`,
`.../output/imffossilfuelsubsidiesdata_sheets.xlsx`, and the WEO and Steffen files in
`.../raw_data/`. Excel recalculates on open. To recalculate headless, use LibreOffice.

## Caveats

See the README sheet in the workbook. The key ones:

- This is a Tier 3 allocation. Group totals are matched exactly. Country values depend
  on the drivers and on β and φ.
- b, c and a are calibrated on one region. The country indices adjust them but are no
  substitute for a country model run.
- Yellow cells are author's assumptions to review: AE/China AFOLU pools, the end-use
  split, the China WACC, public shares, and the PEFA-to-gap mapping.
- Market rates (nominal) are compared with the 6% WB rate (real), as in the method note.
  Set `p_basis = 1` to compare in real terms.
