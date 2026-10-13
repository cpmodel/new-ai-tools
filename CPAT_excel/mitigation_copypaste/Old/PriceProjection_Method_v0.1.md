# Projecting pre-tax and retail energy prices (legacy CPAT method, basics)

Mini method note for the next build step of the copy-pasteable mitigation module. It describes how legacy CPAT projects future energy prices, reduced to the basics: no subsidy phase-outs, no price-liberalisation, no existing carbon prices or ETS, no electricity prices (power module).

Sources: CPAT documentation, Mitigation chapter 3.2 (github.com/cpmodel/cpat_public); legacy workbook `CPAT 1.0pre_456_NoPropData.xlsb`, sheet `Mitigation` rows 629-760 (international prices, selection, forecasting coefficients) and sheet `Prices_int` (raw source data, regional assumptions); `cpat_coded/cpat_model/components/prices/` (`international_prices.py`, `domestic_prices.py`, `prices.py`), which ports the legacy formulas line by line.

## 1. Dimensions and units

- Prices are kept by **fuel x price sector**, as in the IMF price dataset: coal and natural gas by power / residential / industry; gasoline, diesel, LPG, kerosene, other oil products and biomass for `all` sectors (14 combinations with electricity, which is left to the power module). Subsectors take the price of their price sector (road and residential use residential coal and gas prices).
- All prices are in **real US dollars of a reference year** (legacy: 2026). Data in nominal dollars are converted with an inflation index `I_t = P_ref / P_t` (legacy `Mitigation` row 725, "Inflation index, 2026 = 100", and row 630, "1/denominator").
- Units: domestic prices in $/GJ ($/liter for gasoline, diesel, LPG and kerosene; $/bbl for other oil products); international prices in $/bbl (oil), $/ton (coal) and $/MMBtu (gas). Only ratios of international prices are used, so their units drop out.
- **Historical price years** come from data (IMF price dataset: 2021-2024); **projected years** start the year after the last historical year (2025).

## 2. International prices `gp`

1. **Source** (MTInputs `IntEnerPricForeSource`, row 220): IMF, WB, IEA, EIA, AVG (average of the four), IMF-IEA, **IMF-WB*** (default, average of IMF and WB) or Manual. Each source gives crude oil, coal and four gas series (LNG, North America, Europe, global), nominal, by year (legacy `Mitigation` rows 633-695).
2. **Beyond a source's last data year** (IMF 2030, WB 2026, IEA 2040) the price is held constant in real terms (it grows with inflation in nominal terms).
3. **Gas market** per country (`Prices_int` regional assumptions; Egypt: *Global* = average of LNG, North America and Europe).
4. **Adjustment** (MTInputs `IntEnerPricForecastAdjustment`, row 258): *Base* = 1; *High* / *Low* multiply oil by 1.5 / 0.5 and gas and coal by 1.25 / 0.75, from the year after the first projection year (2024).
5. **Real terms**: `gp_t = gp_nom,t x I_t`.

## 3. Supply cost (pre-tax price) `sp`

The supply cost has a fixed part and a part that floats with the international price:

`sp_t = fixsp + fltsp_t (+ ps_t)`

- `fixsp` (constant in real terms) = the **margin over international prices** (IMF dataset `mit.mar`); for coal and gas plus the price-control coefficient x the **domestic production cost** (`mit.coa.prod.cost`, `mit.nga.prod.cost`).
- Historical years: `fltsp_t = sp_t - fixsp (- ps_t)`.
- Projected years: `fltsp_t = fltsp_{t-1} x gp_t / gp_{t-1}`, with oil prices for gasoline, diesel, LPG, kerosene and other oil products, the coal price for coal, the gas price for natural gas; biomass is held constant.
- `ps` (producer-side subsidy) is zero for Egypt and is left out here.

This is the documentation's `sp_t = sp_t0 / gp_t0 x gp_t` applied to the floating part only.

## 4. Taxes and the pass-through of price changes

The **pass-through (price-control) coefficient** `pcc` says how much of a change in the supply cost reaches consumers. It comes from the IMF dataset (`mit.ps`, "Passthrough"; Egypt: coal 1, natural gas 0, oil products 0) and is used as chosen in MTInputs `GovPriceControls` (row 194):
- *Bucketed** (default): raw < 0.25 -> 0; 0.25-0.5 -> 0.5; > 0.5 -> 1;
- *Manual*: 0.8 for all (legacy `Manual inputs`);
- *None*: 1 (full pass-through).
Other oil products and biomass are always 1.

Excise and other taxes `txo` are split in the **last historical year** `L`:
- floating part (an implicit subsidy or tax that absorbs supply-cost changes): `cs_L = txo_L x (1 - pcc)`;
- fixed part: `fx = txo_L - cs_L` (kept constant in real terms, "Fixed" taxes; `Prices_int` gives Fixed for Egypt).

Projected years:

`cs_t = (sp_L - sp_t + cs_L) x (1 - pcc)`; if `cs_L >= 0` then `cs_t = max(cs_L, cs_t)`

`txo_t = fx + cs_t + nce_t` (nce = new policies: new carbon tax, new excise)

- `pcc = 1`: the retail price follows the supply cost one for one.
- `pcc = 0`: the retail price before new policies stays at its last historical level; the floating subsidy (or tax) absorbs the change in the supply cost. A positive floating tax does not fall below its last value (prices are not passed on downwards beyond it).
- New policies (`nce`) are always passed through in full.

## 5. VAT and retail price

`vat_t = (sp_t + txo_t) x vatrate_L` (VAT rate held at its last historical value; VAT is paid on top of other taxes, including new policies)

`rp_t = max(sp_t + vat_t + txo_t, 0.01)`

In the fuel-use equation the after-tax price is `rp_t` (legacy `atp`), by subsector through its price sector.

## 6. Data the step needs

| Item | Source | Coverage |
|---|---|---|
| sp, txo, rp, vatrate, mar, ps (pass-through), production costs, 2021-2024 | `Prices_dom` (already in the workbook) | 221 countries |
| International prices by source, 2021-2045, nominal | legacy `Mitigation` rows 633-695 (built from `Prices_int`) | global |
| Inflation index (reference year = 100) | legacy `Mitigation` rows 630 / 725 | global |
| Gas market and fixed / ad-valorem taxes by country | `Prices_int` regional assumptions | all countries |
| Selectors: source, adjustment, price controls | MTInputs rows 220, 258, 194 | per scenario |

## 7. Left out (later steps)

Producer and consumer subsidy phase-outs and price-control phase-out (phase-out factors = 1); existing carbon taxes and ETS (none for Egypt); the Europe shift to LNG; ad-valorem baseline taxes (Egypt is Fixed); the oil-product share in the VAT of transport fuels; electricity prices.
