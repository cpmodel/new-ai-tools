"""Screenshots of CPAT-AI-Mitigation-MVP-v2.2.xlsm for the combined deck v2.2 (presentation/img/v22_*.png) and chart
data (img/v22_data.json). v2.2: workbook v2.2. Every view is rendered from the real workbook: recalculated in LibreOffice (no LAMBDA support
needed since v1.08), one sheet turned into a values copy (chosen cells show their formula text), rows / columns / print area
set, printed to PDF by LibreOffice and cut to a PNG.

    python presentation/make_screenshots_v2_2.py
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from PIL import Image, ImageChops

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import build_v2_2 as B                      # noqa: E402
import check_v2_2 as C                      # noqa: E402

IMG = os.path.join(HERE, 'img')
LO_PROFILE = 'file://' + os.path.join(tempfile.gettempdir(), 'lo_shots_profile')


def recalculated():
    """Values workbook: the shipped file recalculated in LibreOffice (no LAMBDA in cells since v1.08)."""
    return C.recalc_wb(B.OUT)


def soffice(args, outdir):
    subprocess.run(['soffice', f'-env:UserInstallation={LO_PROFILE}', '--headless'] + args + ['--outdir', outdir],
                   check=True, capture_output=True, timeout=900)


def _unused_recalculated():
    """(formula workbook path, values workbook): LAMBDAs expanded, recalculated in LibreOffice."""
    path = C.variant(openpyxl.load_workbook(B.OUT), 'expand')
    out = tempfile.mkdtemp()
    soffice(['--calc', '--convert-to', 'xlsx', path], out)
    return openpyxl.load_workbook(os.path.join(out, os.path.basename(path)), data_only=True)


def crop(png):
    im = Image.open(png).convert('RGB')
    bg = Image.new('RGB', im.size, (255, 255, 255))
    box = ImageChops.difference(im, bg).getbbox()
    if box:
        im = im.crop((max(box[0] - 8, 0), max(box[1] - 8, 0), box[2] + 8, box[3] + 8))
    im.save(png)


def render(name, sheet, values, area, rows=None, cols=None, show_formula=(), widths=None, hide_cols=(),
           show_cols=(), wrap_cols=(), heights=None):
    """rows: set of rows to show (others hidden) or None (keep the shipped outline state); cols likewise."""
    wb = openpyxl.load_workbook(B.OUT)
    ws = wb[sheet]
    vs = values[sheet]
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                if c.coordinate in show_formula:
                    c._value, c.data_type = c.value, 's'              # show the formula text
                    c.alignment = Alignment(wrap_text=True, vertical='top')
                    c.font = Font(name='Courier New', size=11, color=c.font.color.rgb if c.font and c.font.color
                                  and isinstance(c.font.color.rgb, str) else '000000')
                else:
                    v = vs[c.coordinate].value
                    c.value = v
    for other in [w for w in wb.worksheets if w.title != sheet]:
        wb.remove(other)
    wb.defined_names.clear()
    if rows is not None:                     # explicit rows: drop the outline so collapsed groups do not hide them
        ws.sheet_format.outlineLevelRow = 0
        for r in range(1, ws.max_row + 2):
            d = ws.row_dimensions[r]
            d.outlineLevel, d.collapsed, d.hidden = 0, False, r not in rows
    if show_cols:
        ws.sheet_format.outlineLevelCol = 0
        for d in list(ws.column_dimensions.values()):
            d.outline_level, d.collapsed = 0, False
    for col in hide_cols:
        ws.column_dimensions[col].hidden = True
    for col in show_cols:
        ws.column_dimensions[col].hidden = False
    for col, w in (widths or {}).items():
        ws.column_dimensions[col].width = w
    for col in wrap_cols:
        for r in range(1, ws.max_row + 1):
            ws[f'{col}{r}'].alignment = Alignment(wrap_text=True, vertical='top')
    for r, h in (heights or {}).items():
        ws.row_dimensions[r].height = h
    ws.freeze_panes = None
    ws.print_area = area
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.headings = True
    ws.print_options.gridLines = True
    ws.page_margins.left = ws.page_margins.right = ws.page_margins.top = ws.page_margins.bottom = 0.2
    ws.oddHeader.center.text = ws.oddFooter.center.text = ''
    tmp = tempfile.mkdtemp()
    x = os.path.join(tmp, f'{name}.xlsx')
    wb.save(x)
    soffice(['--convert-to', 'pdf', x], tmp)
    subprocess.run(['pdftoppm', '-png', '-r', '170', '-f', '1', '-l', '1', os.path.join(tmp, f'{name}.pdf'),
                    os.path.join(tmp, name)], check=True)
    png = [f for f in os.listdir(tmp) if f.startswith(name) and f.endswith('.png')][0]
    dst = os.path.join(IMG, f'{name}.png')
    shutil.move(os.path.join(tmp, png), dst)
    crop(dst)
    print('wrote', dst)


def main():
    os.makedirs(IMG, exist_ok=True)
    values = recalculated()
    (k1, b1, y1), (k2, b2, y2) = B.group_cols(1), B.group_cols(2)
    yr = lambda ys, y: L(ys[y - B.BASE_YEAR - 1])
    h = B.SUB_HEAD['rod']
    last = L(y2[-1] + 1)

    # 1. Overview: the sheet as shipped (rolled up), both scenarios
    render('v22_01_overview', 'Mitigation', values, f'A1:{last}{B.R_EMI["co2.pct"] + 1}')

    # 1b. Scenarios left to right: top rows and carbon price, code columns shown
    top = set(range(1, B.R_DEFL + 1)) | {B.B_POL, B.R_CP}
    render('v22_01b_scenarios', 'Mitigation', values, f'A1:{last}{B.R_CP}', rows=top, show_cols=['K'],
           widths={'K': 22, L(k2): 22})

    # 1c. Left section: road subsector expanded, parameter columns shown, first years
    end = h + B.VOFF['nce'] - 1
    rows = {1, 4, 5, 6, B.SEC_BAND['tra'], h, h + 1} | set(range(h + 2, end))
    render('v22_01c_left', 'Mitigation', values, f'A1:{yr(y1, 2026)}{end}', rows=rows,
           show_cols=list('BCDEFG'), hide_cols=['I', 'J', 'K'])

    # 2. Short plain formulas: road gasoline rows, formula text in 2040 (same formula as every year)
    vars_ = ['ctxnew', 'nce', 'atp', 'shp', 'ener']
    rws = [h + B.VOFF[v] + 2 for v in vars_]
    c40 = L(y1[-1])
    cells = {f'{c40}{r}' for r in rws}
    hide = [L(c) for c in range(CI('I'), y1[-1]) if c not in (y1[-2],)]
    render('v22_02_formulas', 'Mitigation', values, f'A1:{c40}{max(rws)}', rows={4, 5} | set(rws),
           show_formula=cells, hide_cols=hide + ['B', 'D', 'E', 'F', 'G'], show_cols=['C', L(y1[-2]), c40],
           widths={c40: 95, 'H': 34, L(y1[-2]): 11}, heights={r: 36 for r in rws})

    # 3. Power: generation by type, shares, VRE, capacity, investment, electricity prices (scenario 1)
    keep = [('gen', m) for m, _n in B.PW_TECH] + [('gtot', ''), ('vre', ''), ('captot', ''), ('nimtot', ''),
                                                    ('usdtot', ''), ('rp', 'res'), ('rp', 'ind'), ('demtot', '')]
    rows = {4, 5, 6, B.B_POW} | {B.R_PW[k] for k in keep}
    render('v22_03_power', 'Mitigation', values, f'A1:{L(y1[-1])}{max(rows)}', rows=rows, hide_cols=['I', 'J', 'K'])

    # 4. Results: fuel use, electricity memo, revenue and CO2 totals for both scenarios
    res = {4, 5, 6, B.B_RES, B.R_TOTAL, B.R_EPOW, B.R_PCT} | {B.R_ELD['all']} | {B.B_REV} \
        | {B.R_REV[k] for k in ('rtx.all', 'rtx.pow', 'rsub.all', 'rsub.pow', 'rnew.all', 'rnew.pow', 'rtot.all',
                                'rtot.chg')} \
        | {B.B_EMI, B.R_EMI['co2.all'], B.R_EMI['co2.pow'], B.R_EMI['co2.chg'], B.R_EMI['co2.pct']}
    render('v22_04_results', 'Mitigation', values, f'A1:{last}{B.R_EMI["co2.pct"]}', rows=res)

    # 5. ScenarioCompare: header and the level block (selected indicators)
    keep = ['mit.ener.all.all.e', 'mit.co2.all.all.e', 'mit.co2.pow.all.e', 'mit.rtot.all.all', 'mit.cptraj',
            'mit.ets.p', 'mit.gtot', 'mit.vre', 'mit.rp.res.ecy', 'mit.captot']
    stems = [s for s, _f in B.CMP_STEMS]
    rows = {4, 6, 7} | {9 + stems.index(s) for s in keep}
    render('v22_05_compare', 'ScenarioCompare', values, f'A1:G{9 + len(stems)}', rows=rows, hide_cols=['A'],
           widths={'B': 44, 'C': 11, 'D': 15, 'E': 15, 'F': 15, 'G': 15}, wrap_cols=[], heights={7: 52})

    # 6. MTInputs: scenario definitions J:N
    n = B.MT_NAMES
    keys = ['CPIntro', 'CPLevelStart', 'CPLevelTarget', 'CPOutro', 'D_Feb_Level_Start_Trans',
            'D_Feb_Level_Target_Trans', 'D_NewETS', 'D_ETSChangeRelStart', 'D_ETSChangeRelTarget']
    rows = {4, 5, 6, 55, 60, 83} | {n[k] for k in keys}
    render('v22_06_definitions', 'MTInputs', values, 'B1:N95', rows=rows, hide_cols=list('CDEFGHI'),
           widths={'B': 40, 'J': 17, 'K': 17, 'L': 17, 'M': 17, 'N': 17}, heights={6: 64, 4: 26})

    # 7. Lambdas sheet: the optional LAMBDA library
    lam = values['Lambdas']
    rows = {1, 3, 4} | set(range(5, 15))
    render('v22_07_lambdas', 'Lambdas', values, 'A1:F14', rows=rows, widths={'F': 110})

    # chart data from StoredResults
    st, names = C.stored_blocks(values['StoredResults'])
    years = list(range(B.BASE_YEAR, B.LAST_YEAR + 1))
    data = {'names': {str(k): v for k, v in names.items()}, 'years': years,
            'co2': {str(k): st[k]['mit.co2.all.all.e'] for k in st},
            'co2_pow': {str(k): st[k]['mit.co2.pow.all.e'] for k in st},
            'vre': {str(k): st[k]['mit.vre'] for k in st},
            'gtot': {str(k): st[k]['mit.gtot'] for k in st},
            'ets_p': st[5]['mit.ets.p'], 'co2_ets': st[5]['mit.co2.ets.all.all'], 'co2_cap': st[5]['mit.co2.cap.all.all']}
    with open(os.path.join(IMG, 'v22_data.json'), 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=1)
    print('wrote v22_data.json')


if __name__ == '__main__':
    main()
