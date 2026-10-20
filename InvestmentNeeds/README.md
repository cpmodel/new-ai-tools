# InvestmentNeeds — Mitigation (Net Zero) Investment Needs Model

Two models live here:

- **`Mitigation_Investment_Needs_v1.1.xlsx`** — 2030 snapshot for 195 economies: group totals
  (IHLEG, IEA-IFC) allocated to countries, three cost-of-capital lenses (this README).
- **`national_model/NMIM_v0.1.xlsx`** — National Mitigation Investment Model: a simplified
  national stock-and-flow model for 186 countries, 2024-2050, built to CPAT norms. See
  [`national_model/NMIM_Method_v0_1.md`](national_model/NMIM_Method_v0_1.md).

Task log: [`CAVEATS.md`](CAVEATS.md). Superseded versions: `Old/`.

## The v1.1 workbook

`Mitigation_Investment_Needs_v1.1.xlsx` (current; v1.0 in `Old/`) estimates 2030 **total** and **incremental**
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
6. **Savings, co-benefits and abatement** (v1.1) by sector method:
   - *Engineering* (power generation, transport): renewable output displaces the country's
     fossil generation mix (Ember) at CPAT 2024 fuel and air-pollution costs, with heat
     rates and IPCC 2006 CO2 factors; EVs replace gasoline cars per vehicle.
   - *Enabling* (grids): no benefits of their own (switch `p_grid_credit`).
   - *Calibrated* (buildings, industry): b = 0.133, c = 0.129 (SISEPUEDE LAC) scaled by
     CPAT country indices.
   - *Cost per tonne* (low-emission fuels USD 100/t, AFOLU USD 50/t).
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

**v1.0 and v1.1 ship without Damodaran values**: the source host could not be reached when the
workbook was built, so the 111 economies without an observed WACC use the IHLEG group
rate.

## Build

```bash
pip install pandas numpy openpyxl matplotlib
python InvestmentNeeds/build/build_investment_model.py     # the workbook
python InvestmentNeeds/build/run_sensitivity.py            # Sensitivity sheet (needs LibreOffice)
python InvestmentNeeds/build/make_summary_image.py         # one-page PNG summary
```

The methodology document (sector by sector, sensitivity, and what a national-level
model would need) is a shared doc linked from the session; its content mirrors this
README and the workbook's README sheet.

It reads only repository files: `TaxMacroFiscalDashboard/build/output/global_data_wide.csv`,
`.../output/imffossilfuelsubsidiesdata_sheets.xlsx`, and the WEO and Steffen files in
`.../raw_data/`. Excel recalculates on open. To recalculate headless, use LibreOffice.

## Caveats

See the README sheet in the workbook. The key ones:

- This is a Tier 3 allocation. Group totals are matched exactly. Country values depend
  on the drivers and on β and φ.
- Buildings and industry still use the one-region calibration for savings and
  co-benefits; power is conservative (no credit for fossil plants not built; 2024 fuel
  prices).
- Yellow cells are author's assumptions to review: AE/China AFOLU pools, the end-use
  split, the China WACC, public shares, and the PEFA-to-gap mapping.
- Market rates (nominal) are compared with the 6% WB rate (real), as in the method note.
  Set `p_basis = 1` to compare in real terms.

## National model (NMIM v0.1)

```bash
python InvestmentNeeds/national_model/extract_data_v0_1.py   # data/*.csv from repo sources
python InvestmentNeeds/national_model/nmim_v0_1.py           # engine: results/ for all countries
python InvestmentNeeds/national_model/build_workbook_v0_1.py # NMIM_v0.1.xlsx (CPAT layout)
python InvestmentNeeds/national_model/check_v0_1.py          # LibreOffice checks -> check_report_v0_1.md
```

One equation specification drives both the Python engine (all countries) and the
workbook's live formulas (the country chosen on Settings!C4), and the check confirms they
agree. Method, data, validation and open issues: `national_model/NMIM_Method_v0_1.md`.
