"""Word register of bugs, known issues and potential upgrades of CPAT-AI-Mitigation-MVP v2.2
(CPAT-AI-Mitigation-MVP_BugsIssuesUpgrades_v2.2.docx). Content from issues_register_v2_2.py (also used by the Word
documentation) and the workbook sheet LegacyDiff (via build_v2_2). The effect of each legacy power bug fix is computed
with power_model_ref_v2_2 on the cached legacy inputs of Egypt (baseline and carbon-tax run), so every number is
reproducible. Replaces PowerModel_BugsAndImprovements_v1_0.docx (power bugs only).

    python make_issues_doc_v2_2.py      (needs node with the docx package: NODE_PATH)
"""
import json
import os
import subprocess
import tempfile

import build_v2_2 as B
import issues_register_v2_2 as IR
import power_model_ref_v2_2 as PM

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, f'{B.NAME}_BugsIssuesUpgrades_v{B.VERSION}.docx')
YEARS = list(range(2022, 2041))
K30, K40 = YEARS.index(2030), YEARS.index(2040)
EGYPT = {'services': 'changes results', 'aee': 'changes results', 'income': 'changes results',
         'invcost': 'investment spending (reporting)', 'cmin': 'carbon-tax run'}


def effects():
    def kpi(o):
        gen = lambda k: sum(o['gen'][i][k] for i in range(9)) * 8.76 / 1000
        return {'gen30': gen(K30), 'gen40': gen(K40), 'co2_40': sum(o['co2'][i][K40] for i in range(9)),
                'gas40': o['shr'][1][K40], 'vre40': o['vre'][K40],
                'usd': sum(o['usd'][i][k] for i in range(9) for k in range(len(YEARS))) / 1000}
    base = {run: kpi(PM.engine(PM.legacy_inputs(run), YEARS, fix=frozenset(), shock=True)) for run in ('base', 'ctax')}
    rows = []
    for name in list(PM.FIXES) + ['all']:
        fx = PM.ALL_FIXES if name == 'all' else frozenset([name])
        b = kpi(PM.engine(PM.legacy_inputs('base'), YEARS, fix=fx, shock=True))
        c = kpi(PM.engine(PM.legacy_inputs('ctax'), YEARS, fix=fx, shock=True))
        pct = lambda a, z: '0' if abs(a / z - 1) < 5e-5 else f'{(a / z - 1) * 100:+.1f}%'
        rows.append([name if name != 'all' else 'All fixes', pct(b['gen30'], base['base']['gen30']),
                     pct(b['gen40'], base['base']['gen40']), pct(b['co2_40'], base['base']['co2_40']),
                     pct(c['gas40'], base['ctax']['gas40']), pct(c['vre40'], base['ctax']['vre40']),
                     pct(b['usd'], base['base']['usd'])])
    return base, rows


def blocks():
    base, eff = effects()
    n = {c: sum(1 for *_x, k in IR.items() if k == c) for c in
         ('Legacy bug', 'MVP bug (fixed)', 'Known issue', 'Simplification / not yet', 'Upgrade')}
    up = {p: [u for u in IR.UPGRADES if u[3] == p] for p in ('High', 'Medium', 'Low')}
    bl = [{'t': 'title', 'text': 'CPAT-AI-Mitigation-MVP: bugs, known issues and potential upgrades'},
          {'t': 'subtitle', 'text': f'Register for version {B.VERSION} (workbook {B.NAME}-v{B.VERSION}.xlsm), Egypt data. '
                                    '11 October 2026.'},
          {'t': 'h1', 'text': '1. Summary'},
          {'t': 'p', 'text': 'One list of everything that is wrong, simplified or missing, so that a reviewer can see '
                             'the state of the model in one place. Every item has an ID that the other documents '
                             'use. The register is kept in issues_register_v2_2.py and the workbook sheet LegacyDiff; '
                             'the check script tests that both agree.'},
          {'t': 'table', 'head': ['Part', 'IDs', 'Items', 'What it lists'], 'widths': [2600, 1100, 900, 5000],
           'rows': [['2. Bugs in legacy CPAT', 'L', str(n['Legacy bug']), 'Defects of the legacy workbook found while '
                     'rebuilding it; fixed in the MVP unless stated'],
                    ['3. Bugs found and fixed in the MVP', 'M', str(n['MVP bug (fixed)']), 'Defects of earlier MVP '
                     'versions, all fixed'],
                    ['4. Known issues', 'I', str(n['Known issue']), 'What a user of v2.2 should know'],
                    ['5. Simplifications and not yet modelled', 'LD', str(n['Simplification / not yet']),
                     'LegacyDiff rows of type Not yet, Simplification or To check (row number = ID)'],
                    ['6. Potential upgrades', 'U', str(n['Upgrade']), f'{len(up["High"])} high, {len(up["Medium"])} '
                     f'medium, {len(up["Low"])} low priority']]},
          {'t': 'bullets', 'items': [
              f'**Legacy power model:** 10 bugs (L1-L10). 5 change the Egypt results today; with all fixes Egypt\'s '
              f'generation in 2040 is **{eff[-1][2]}** higher than legacy and power CO2 **{eff[-1][3]}** higher.',
              '**Nothing open is a known error in v2.2.** The known issues are decisions to take (I4), checks still to '
              'run (I1, U1, U2) or simplifications that do not matter for Egypt (I7).',
              '**Priorities:** review in Excel (U1), compare with legacy for Egypt (U2), decide the price basis of '
              'demand (U3), report the legacy bugs (U4).']},
          {'t': 'h1', 'text': '2. Bugs in legacy CPAT'},
          {'t': 'p', 'text': 'Found by reading every formula of the legacy workbook (CPAT 1.0pre_456_NoPropData.xlsb) '
                             'and replicating it in Python. L1-L10 are in the engineer power model: a Python reference '
                             '(power_model_ref_v2_2.py) reproduces 208 legacy rows of the baseline and the carbon-tax '
                             'run to 1e-13 with these bugs switched on; each bug is a switch (FIXES).'},
          {'t': 'table', 'head': ['ID', 'Bug', 'Where (legacy)', 'Egypt effect / status'],
           'widths': [600, 4400, 1900, 2700],
           'rows': [[f'L{i + 1}', t, rows, EGYPT.get(k, 'latent')] for i, (k, t, rows, *_r) in enumerate(IR.BUGS)]
           + [[i, t, w, st] for i, t, w, _d, st in IR.LEGACY]}]
    for i, (k, t, rows, does, should, why) in enumerate(IR.BUGS):
        bl += [{'t': 'h2', 'text': f'L{i + 1}. {t}'},
               {'t': 'p', 'text': f'**Where:** {rows}.'},
               {'t': 'p', 'text': f'**Legacy:** {does}'},
               {'t': 'p', 'text': f'**Correct:** {should}'},
               {'t': 'p', 'text': f'**Why it matters:** {why}'}]
    bl += [{'t': 'h2', 'text': 'Effect of each power fix (legacy Egypt inputs)'},
           {'t': 'p', 'text': 'Change against legacy. Generation and CO2 from the baseline; gas and variable-renewable '
                              '(VRE) shares from the carbon-tax run; investment is the 2022-2040 total of the baseline. '
                              f'Legacy levels: generation {base["base"]["gen30"]:,.0f} TWh (2030) and '
                              f'{base["base"]["gen40"]:,.0f} TWh (2040); power CO2 {base["base"]["co2_40"]:.1f} MtCO2 '
                              f'(2040); carbon-tax run gas share {base["ctax"]["gas40"]:.1%} and VRE share '
                              f'{base["ctax"]["vre40"]:.1%} (2040); investment USD {base["base"]["usd"]:,.0f} bn.'},
           {'t': 'table', 'head': ['Fix', 'Generation 2030', 'Generation 2040', 'Power CO2 2040',
                                   'Gas share 2040 (carbon tax)', 'VRE share 2040 (carbon tax)', 'Investment USD'],
            'widths': [1500, 1300, 1300, 1300, 1500, 1500, 1300],
            'rows': [[(f'L{[k for k, *_x in IR.BUGS].index(r[0]) + 1} ' + IR.SHORT[r[0]]) if r[0] in IR.SHORT
                      else '**All fixes together**'] + r[1:] for r in eff]},
           {'t': 'note', 'text': 'Effects do not add up exactly, because the fixes interact (more demand also means '
                                 'more investment).'},
           {'t': 'h2', 'text': 'Other legacy defects (L11-L16)'},
           {'t': 'table', 'head': ['ID', 'Defect', 'What legacy does', 'Status in the MVP'],
            'widths': [600, 2300, 3700, 3000], 'rows': [[i, t, d, st] for i, t, _w, d, st in IR.LEGACY]},
           {'t': 'h1', 'text': '3. Bugs found and fixed in the MVP'},
           {'t': 'p', 'text': 'Defects of the MVP itself, found by the checks or by review, with the version that '
                              'introduced and fixed them. None is open.'},
           {'t': 'table', 'head': ['ID', 'Bug', 'Found / fixed', 'What happened', 'Fix'],
            'widths': [500, 2000, 1200, 3200, 2700], 'rows': [list(m) for m in IR.MVP_BUGS]},
           {'t': 'h1', 'text': '4. Known issues'},
           {'t': 'table', 'head': ['ID', 'Issue', 'What it means', 'Documented in'], 'widths': [500, 2200, 5000, 1900],
            'rows': [list(x) for x in IR.ISSUES]},
           {'t': 'h1', 'text': '5. Simplifications and not yet modelled (from LegacyDiff)'},
           {'t': 'p', 'text': 'The workbook sheet LegacyDiff lists every difference with legacy CPAT. These are the rows '
                              'of type Not yet, Simplification and To check; the ID is the LegacyDiff row number. The '
                              'other rows are decisions, data choices or equal to legacy.'},
           {'t': 'table', 'head': ['ID', 'Area', 'Legacy CPAT', 'This model', 'Effect', 'Type'],
            'widths': [750, 1500, 2100, 2500, 1500, 1250], 'rows': [list(x) for x in IR.legacy_diff_items()]},
           {'t': 'h1', 'text': '6. Potential upgrades'}]
    for p in ('High', 'Medium', 'Low'):
        bl += [{'t': 'h2', 'text': f'{p} priority'},
               {'t': 'table', 'head': ['ID', 'Upgrade', 'What and why'], 'widths': [600, 2800, 6200],
                'rows': [[i, t, d] for i, t, d, _p in up[p]]}]
    bl += [{'t': 'h1', 'text': '7. Improvements already made in the MVP'},
           {'t': 'bullets', 'items': IR.IMPROVEMENTS},
           {'t': 'h1', 'text': '8. Sources and how this register is kept consistent'},
           {'t': 'bullets', 'items': [
               'issues_register_v2_2.py: the register (L, M, I, U items); LegacyDiff rows are read from build_v2_2.py.',
               f'check_v2_2.py (report check_report_v{B.VERSION}.md): IDs unique, every LegacyDiff row of these types '
               'included, workbook texts free of stale statements, file references valid.',
               'CAVEATS.md (repository root): the caveats of each task; TODO.md: open decisions.',
               f'{B.NAME}_Documentation_v{B.VERSION}.docx: method and quick start; its appendix summarises this '
               'register by ID.',
               'make_issues_doc_v2_2.py builds this document (power-fix effects computed when it runs).']}]
    return bl


def main():
    data = {'title': 'CPAT-AI-Mitigation-MVP: bugs, known issues and potential upgrades',
            'footer': f'CPAT-AI-Mitigation-MVP v{B.VERSION} - bugs, issues and upgrades', 'blocks': blocks()}
    src = os.path.join(tempfile.mkdtemp(), 'blocks.json')
    with open(src, 'w', encoding='utf-8') as f:
        json.dump(data, f)
    subprocess.run(['node', os.path.join(HERE, 'render_docx_v2_2.js'), src, OUT], check=True)
    print('wrote', OUT)


if __name__ == '__main__':
    main()
