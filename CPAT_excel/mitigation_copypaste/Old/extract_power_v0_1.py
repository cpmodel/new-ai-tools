"""Extract the power generation-cost and power-price inputs from the legacy workbook (Egypt, sheet Mitigation,
section 4 'Technoeconomic (engineer) power model', rows 2954-3927) to CSVs for CPAT-AI-Mitigation-MVP.

Writes (data/):
  power_tech.csv     one row per generation type: capital cost, lifetime, efficiency, capacity factor, WACC,
                     decommissioning and transmission capex, fixed and variable O&M, fixed fuel cost (nuclear),
                     price fuel of section 2, VRE flag, EF of the power sector
  power_paths.csv    year paths 2022-2040: capex time factors, storage cost components, and the interim rows that
                     the engineer model will compute later (generation shares, investment shares, consumption)
  power_params.csv   scalars: transmission and distribution add-ons, pass-through, generation/consumption,
                     storage parameters, discounted lifetime of renewables
  legacy_power_reference.csv   cached legacy results and legacy inputs used only by the checks (method = legacy)

Every row is located by its table id and label (asserted), so a changed legacy file fails loudly.

    python extract_power_v0_1.py
"""
import csv
import os

from pyxlsb import open_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
LEGACY = os.path.join(HERE, '..', '..', 'cpat_excel_original', 'CPAT 1.0pre_456_NoPropData.xlsb')
DATA = os.path.join(HERE, 'data')
YEARS = list(range(2022, 2041))
COL = {y: y - 2011 for y in YEARS}          # legacy Mitigation: 2022 in column L (index 11)
TECH = [('coa', 'Coal'), ('nga', 'Natural gas'), ('oop', 'Oil'), ('nuc', 'Nuclear'), ('wnd', 'Wind'),
        ('sol', 'Solar'), ('hyd', 'Hydro'), ('ore', 'Other renewables'), ('bio', 'Biomass')]
PRICE_CODE = {'coa': 'coa.pow', 'nga': 'nga.pow', 'oop': 'oop.all', 'bio': 'bio.all'}   # section 2 price fuels
VRE = {'wnd', 'sol', 'ore'}                 # carry the marginal storage cost (legacy C3)
VRE_SHARE = {'wnd', 'sol'}                  # count in the variable-renewable share (legacy row 3907)
SRC = 'CPAT 1.0pre_456_NoPropData.xlsb, Mitigation row {}'


def read_rows():
    rows = {}
    with open_workbook(LEGACY) as wb, wb.get_sheet('Mitigation') as sh:
        for i, row in enumerate(sh.rows()):
            if 2950 <= i + 1 <= 3935:
                rows[i + 1] = [c.v for c in row]
            if i + 1 > 3935:
                break
    return rows


def main():
    R = read_rows()

    def label(r, text, col=3):
        assert str(R[r][col]).strip().startswith(text), (r, R[r][col], text)

    def table(r0, tid):
        assert R[r0][2] == tid, (r0, R[r0][2], tid)

    def path(r):
        return [R[r][COL[y]] for y in YEARS]

    for r0, tid in ((2976, 'A2'), (3048, 'A8'), (3060, 'A9'), (3109, 'A14'), (3120, 'A15'), (3131, 'A16'),
                    (3168, 'B1'), (3180, 'B2'), (3202, 'B4'), (3338, 'D1'), (3349, 'D2'), (3451, 'E0'), (3036, 'A7')):
        table(r0, tid)
    # technology table (row offsets inside each legacy table follow TECH order). cax0 = capital cost at capex time
    # factor 1: the base-year capital cost (B2 column H) divided by the base-year factor (A7), so cax(t) = cax0 x tcf(t)
    tech = []
    for i, (f, name) in enumerate(TECH):
        for r in (2977, 3049, 3061, 3110, 3121, 3132, 3181, 3203, 3339, 3350, 3452, 3038, 3169):
            label(r + i, name.split()[0] if f != 'oop' else 'Oil')
        tech.append({
            'code': f, 'name': name,
            'cax0_usd_per_kw': R[3181 + i][7] / R[3038 + i][COL[2022]], 'life_years': R[2977 + i][19],
            'efficiency': R[3061 + i][COL[2022]], 'capacity_factor': R[3049 + i][COL[2022]],
            'wacc': R[3203 + i][7], 'dtc_usd_per_kw': R[3110 + i][COL[2022]],
            'tcx_usd_per_kw': R[3121 + i][COL[2022]], 'opf_usd_per_kwh': R[3132 + i][COL[2022]],
            'vop_usd_per_kwh': R[3339 + i][COL[2022]],
            'fuel_fixed_usd_per_kwh': R[3350 + i][COL[2022]] if f == 'nuc' else 0.0,
            'rns_usd_per_kwh': 0.0, 'price_code': PRICE_CODE.get(f, ''), 'vre': 1 if f in VRE else 0,
            'vre_share': 1 if f in VRE_SHARE else 0,
            'source': SRC.format('2977-3350 (A2, A8, A9, A14-A16, B2, B4, D1, D2)')})
    with open(os.path.join(DATA, 'power_tech.csv'), 'w', newline='', encoding='utf-8') as fh:
        w = csv.DictWriter(fh, fieldnames=list(tech[0]))
        w.writeheader()
        w.writerows(tech)
    # year paths
    label(3887, 'Capex battery'); label(3888, 'Capex interface'); label(3891, 'Opex per kwh storage')
    label(3900, 'Capex per kwh power produced'); label(3901, 'Opex Fix'); label(3902, 'Opex Var')
    label(3621, 'Residential Power Consumption'); label(3650, 'Industrial Power Consumption')
    paths = []
    for i, (f, _n) in enumerate(TECH):
        paths.append([f'tcf|{f}', 'factor', 'capex time factor (A7)', 3038 + i] + path(3038 + i))
    for i, (f, _n) in enumerate(TECH):
        paths.append([f'gns|{f}', 'share', 'INTERIM generation share (E0; engineer model later)', 3452 + i]
                     + path(3452 + i))
    for i, (f, _n) in enumerate(TECH):
        paths.append([f'phi|{f}', 'share', 'INTERIM new investment / remaining capacity (B1; engineer model later)',
                      3169 + i] + path(3169 + i))
    paths += [['cbat', '$/kWh storage', 'battery capex (J2)', 3887] + path(3887),
              ['cint', '$/kW interface', 'battery interface capex (J2)', 3888] + path(3888),
              ['obat', '$/kWh storage/y', 'battery opex (J2)', 3891] + path(3891),
              ['cel', '$/kWh produced', 'electrolyser cost per marginal unit: capex + fixed + variable opex (J2)',
               '3900-3902'] + [R[3900][COL[y]] + R[3901][COL[y]] + R[3902][COL[y]] for y in YEARS],
              ['cons|res', 'GWh', 'INTERIM residential electricity consumption (G1; demand later)', 3621] + path(3621),
              ['cons|ind', 'GWh', 'INTERIM industrial electricity consumption (G2; demand later)', 3650] + path(3650)]
    with open(os.path.join(DATA, 'power_paths.csv'), 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        w.writerow(['key', 'unit', 'description', 'legacy_row'] + YEARS)
        w.writerows(paths)
    # scalars
    label(3599, 'Transmission'); label(3628, 'Transmission'); label(3596, 'Passthrough')
    label(3723, 'Generation/Consumption'); label(3883, 'Discounted Lifetime of renewables')
    for r, t in ((3871, 'Percent allocation of ST'), (3873, 'Marginal hours ST'), (3874, 'kwh storage to kw'),
                 (3877, 'Starting point of long term')):
        label(r, t)
    params = [('tmc|res', R[3599][COL[2022]], '$/kWh', 'transmission and distribution, residential (G1)', 3599),
              ('tmc|ind', R[3628][COL[2022]], '$/kWh', 'transmission and distribution, industrial (G2)', 3628),
              ('pass', R[3596][COL[2022]], 'share', 'pass-through of supply-cost changes to end users (G1)', 3596),
              ('gen_per_cons', R[3723][8], 'ratio', 'generation / final consumption (H2, base year)', 3723),
              ('dlf_ren', R[3883][COL[2022]], 'years', 'discounted lifetime of renewables (J2)', 3883),
              ('st_alloc', R[3871][4], 'share', 'allocation of short-term storage cost to VRE (J1)', 3871),
              ('st_marg_hours', R[3873][4], 'hours', 'marginal hours of short-term storage per unit VRE (J1)', 3873),
              ('st_ratio', R[3874][4], 'hours', 'kWh storage per kW interface (J1)', 3874),
              ('lt_start', R[3877][4], 'share', 'VRE share where long-term storage starts (J1)', 3877)]
    with open(os.path.join(DATA, 'power_params.csv'), 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        w.writerow(['key', 'value', 'unit', 'description', 'legacy_row'])
        w.writerows(params)
    # legacy results and inputs for the method test
    ref = []
    blocks = {'B2 caxav': 3181, 'B6 stoav': 3225, 'B8 fix': 3247, 'C3 msc': 3282, 'C7 lfx': 3326,
              'D4 vbc': 3374, 'E3 gnc': 3486, 'A13 cax': 3099}
    for name, r0 in blocks.items():
        for i, (f, _n) in enumerate(TECH):
            ref.append([f'{name}|{f}', r0 + i] + path(r0 + i))
    for name, r in (('E3 gncav', 3495), ('J msc_st', 3910), ('J msc_lt', 3911), ('J vre', 3907),
                    ('J cph', 3905), ('G1 sc', 3600), ('G2 sc', 3629), ('G1 rppre', 3613), ('G2 rppre', 3642),
                    ('G1 vat', 3611), ('G2 vat', 3640), ('A18 coa', 3155), ('A18 nga', 3156), ('A18 oop', 3157),
                    ('A18 bio', 3163)):
        ref.append([name, r] + path(r))
    with open(os.path.join(DATA, 'legacy_power_reference.csv'), 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        w.writerow(['key', 'legacy_row'] + YEARS)
        w.writerows(ref)
    print(f'wrote power_tech.csv ({len(tech)}), power_paths.csv ({len(paths)}), power_params.csv ({len(params)}), '
          f'legacy_power_reference.csv ({len(ref)})')


if __name__ == '__main__':
    main()
