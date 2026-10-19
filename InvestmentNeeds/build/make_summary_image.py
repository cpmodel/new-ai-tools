import openpyxl, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
wb = openpyxl.load_workbook('/home/user/new-ai-tools/InvestmentNeeds/Mitigation_Investment_Needs_v1.0.xlsx', data_only=True)
ws = wb['Results_Summary']
rows = {r[0]: r[1:5] for r in ws.iter_rows(min_row=5, max_row=60, max_col=5, values_only=True) if r[0]}
sect = []
for r in ws.iter_rows(min_row=40, max_row=70, max_col=6, values_only=True):
    if r[0] and isinstance(r[4], (int, float)) and r[0] != 'Total' and not r[0].startswith('Net'): sect.append((r[0], r[4]))
I = rows['Total investment need, I']; dI = rows['Incremental investment need, dI']; gdp = rows['GDP 2030 (USD bn, 2024 prices)']
lens = [('(a) Market cost of capital', 'INCREMENTAL COST (a) market cost of capital'),
        ('(a) Market, high case', 'INCREMENTAL COST (a) market, high case'),
        ('(b) World Bank rate, 6%', 'INCREMENTAL COST (b) World Bank rate'),
        ('(c) WB rate + local co-benefits', 'INCREMENTAL COST (c) WB rate + local co-benefits')]
SURF, T1, T2, T3, GRID = '#fcfcfb', '#0b0b0b', '#52514e', '#8a8984', '#e4e3df'
COL = {'Advanced': '#2a78d6', 'China': '#eb6834', 'Other EMDCs': '#1baf7a'}
plt.rcParams['font.family'] = 'DejaVu Sans'
fig = plt.figure(figsize=(10.8, 19.2), dpi=100, facecolor=SURF)
def txt(x, y, s, size, color=T1, weight='normal', **kw): fig.text(x, y, s, fontsize=size, color=color, weight=weight, **kw)
txt(0.06, 0.955, 'Mitigation investment needs, 2030', 30, weight='bold')
txt(0.06, 0.933, 'Net zero pathway · 195 economies · 7 sectors · USD a year, 2023 prices', 15, T2)
# hero tiles
tiles = [('Total investment', f'${I[3]/1000:.1f} tn', f'{I[3]/gdp[3]:.1%} of world GDP'),
         ('Incremental investment', f'${dI[3]/1000:.1f} tn', f'{dI[3]/gdp[3]:.1%} of world GDP'),
         ('Incremental cost, lens (c)', f'−${-rows[lens[3][1]][3]:.0f} bn', 'net saving with co-benefits')]
for k, (lab, big, sub) in enumerate(tiles):
    x0 = 0.06 + k * 0.30
    fig.patches.append(FancyBboxPatch((x0, 0.825), 0.28, 0.085, boxstyle='round,pad=0,rounding_size=0.012',
                                      transform=fig.transFigure, fc='#f1f0ec', ec='none'))
    txt(x0 + 0.015, 0.89, lab, 13, T2); txt(x0 + 0.015, 0.852, big, 30, weight='bold'); txt(x0 + 0.015, 0.833, sub, 11.5, T2)
# chart 1: incremental cost by lens
ax = fig.add_axes([0.33, 0.50, 0.61, 0.25], facecolor=SURF)
txt(0.06, 0.785, 'Incremental cost by cost-of-capital lens', 19, weight='bold')
txt(0.06, 0.767, 'USD bn a year · below zero = net saving, above zero = net cost', 13, T2)
groups = [('Advanced', 0), ('China', 1), ('Other EMDCs', 2)]
h = 0.24
for li, (lab, key) in enumerate(lens):
    for gname, gi in groups:
        v = rows[key][gi]; y = li + (gi - 1) * (h + 0.03)
        ax.barh(y, v, height=h, color=COL[gname], edgecolor=SURF, linewidth=2)
        ax.text(v + (3 if v >= 0 else -3), y, (f'{v:+.0f}' if round(v) != 0 else '0'), va='center', ha='left' if v >= 0 else 'right', fontsize=11, color=T1)
ax.set_yticks(range(4)); ax.set_yticklabels([l for l, _ in lens], fontsize=13, color=T1)
ax.invert_yaxis(); ax.axvline(0, color=T2, lw=1.2)
ax.set_xlim(-175, 45); ax.tick_params(axis='x', colors=T3, labelsize=11); ax.tick_params(axis='y', length=0)
ax.grid(axis='x', color=GRID, lw=0.8); ax.set_axisbelow(True)
for s in ax.spines.values(): s.set_visible(False)
for k, (gname, _) in enumerate(groups):
    x0 = 0.33 + k * 0.2
    fig.patches.append(FancyBboxPatch((x0, 0.467), 0.018, 0.009, boxstyle='round,pad=0,rounding_size=0.002', transform=fig.transFigure, fc=COL[gname], ec='none'))
    txt(x0 + 0.025, 0.466, gname, 12.5, T1)
# chart 2: total investment by sector
txt(0.06, 0.435, 'Total investment by sector', 19, weight='bold')
txt(0.06, 0.417, 'USD bn a year, world', 13, T2)
ax2 = fig.add_axes([0.33, 0.215, 0.61, 0.19], facecolor=SURF)
names = [n.split(' (')[0] for n, _ in sect]; vals = [v for _, v in sect]
ax2.barh(range(len(vals)), vals, height=0.62, color='#2a78d6')
for i, v in enumerate(vals): ax2.text(v + 20, i, f'{v:,.0f}', va='center', fontsize=11.5, color=T1)
ax2.set_yticks(range(len(vals))); ax2.set_yticklabels(names, fontsize=13, color=T1); ax2.invert_yaxis()
ax2.set_xlim(0, 1950); ax2.tick_params(axis='x', colors=T3, labelsize=11); ax2.tick_params(axis='y', length=0)
ax2.grid(axis='x', color=GRID, lw=0.8); ax2.set_axisbelow(True)
for s in ax2.spines.values(): s.set_visible(False)
# key messages
txt(0.06, 0.175, 'What it shows', 19, weight='bold')
msgs = [f'Other EMDCs need {I[2]/gdp[2]:.1%} of GDP in total, against {I[0]/gdp[0]:.1%} in advanced economies.',
        'At market rates the EMDC case is thin: −\\$5 bn, turning to a +\\$14 bn net cost in the high case.',
        'At 6% EMDCs save more; China shows a small cost (its observed WACC is 3.9%).',
        'Local co-benefits turn every group to a large net saving: the case for public finance.']
for k, m in enumerate(msgs): txt(0.06, 0.148 - k * 0.024, '•  ' + m, 13.5, T1)
txt(0.06, 0.035, 'Sources: IHLEG (2024), IEA-IFC (2023), IMF CPAT and WEO, IRENA, Steffen et al. (2025), SISEPUEDE LAC calibration.\n'
    'Tier 3 allocation; savings and co-benefit ratios calibrated on one region. Model: Mitigation_Investment_Needs_v1.0.xlsx', 10.5, T3)
out = '/home/user/new-ai-tools/InvestmentNeeds/Mitigation_Investment_Needs_v1.0_summary.png'
fig.savefig(out, facecolor=SURF); print(out, sect)
