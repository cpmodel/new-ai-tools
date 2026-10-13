"""Screenshots of CPAT-AI-Mitigation-MVP-v1.00.xlsx for the overview deck (presentation/img/*.png).

Each view is rendered from the real workbook: the file is recalculated in LibreOffice (LAMBDAs expanded so that
the right column shows values), one sheet is turned into a values copy (chosen cells show their formula text
instead), rows / columns / print area are set, and LibreOffice prints it to PDF, which is cut to a PNG.

    python presentation/make_screenshots_v1_0.py
"""
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
sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'Old'))   # v1.00 builder and checker (archived)
import build_v1_00 as B                      # noqa: E402
import check_v1_00 as C                      # noqa: E402

IMG = os.path.join(HERE, 'img')
LO_PROFILE = 'file://' + os.path.join(tempfile.gettempdir(), 'lo_shots_profile')


def soffice(args, outdir):
    subprocess.run(['soffice', f'-env:UserInstallation={LO_PROFILE}', '--headless'] + args + ['--outdir', outdir],
                   check=True, capture_output=True, timeout=900)


def recalculated():
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
    render('01_overview', 'Mitigation', values, f'A1:{last}{B.R_EMI["co2.pct"] + 1}')

    # 2. Scenarios left to right: top rows and carbon price, code columns shown
    top = set(range(1, B.R_DEFL + 1)) | {B.B_POL, B.R_CP}
    render('02_scenarios', 'Mitigation', values, f'A1:{last}{B.R_CP}', rows=top, show_cols=['K'],
           widths={'K': 22, L(k2): 22})

    # 3. Left section: road subsector expanded, parameter columns shown, first years
    end = h + B.VOFF['nce'] - 1                              # ctxnew, ets, ntx blocks of road
    rows = {1, 4, 5, 6, B.SEC_BAND['tra'], h} | set(range(h + 1, end))
    render('03_left_section', 'Mitigation', values, f'A1:{yr(y1, 2026)}{end}', rows=rows,
           show_cols=list('BCDEFG'), hide_cols=['I', 'J', 'K'])

    # 4. Two formula styles: road gasoline rows, plain formula in 2029 (left) and LAMBDA in 2040 (right)
    vars_ = ['ctxnew', 'nce', 'atp', 'ener']
    rws = [h + B.VOFF[v] + 2 for v in vars_]
    c29, c40 = yr(y1, 2029), L(y1[-1])
    cells = {f'{c29}{r}' for r in rws} | {f'{c40}{r}' for r in rws}
    hide = [L(c) for c in range(CI('I'), y1[-1]) if L(c) != c29]
    render('04_formulas', 'Mitigation', values, f'A1:{c40}{max(rws)}', rows={4, 5} | set(rws),
           show_formula=cells, hide_cols=hide + ['B', 'D', 'E', 'F', 'G'], show_cols=['C', c29, c40],
           widths={c29: 62, c40: 62, 'H': 30}, heights={r: 62 for r in rws})

    # 5. MTInputs: legacy rows, template columns and one Used-for-calculation column per scenario
    render('05_mtinputs', 'MTInputs', values, 'A1:K30', rows=set(range(1, 31)) - {7, 8, 9, 11, 13, 14, 15, 16},
           hide_cols=['D', 'E', 'I'], widths={'J': 30, 'K': 30})

    # 6. Results: fuel use, revenue and CO2 totals for both scenarios (2030-2039 rolled up)
    res = {4, 5, 6, B.B_RES, B.R_TOTAL, B.R_PCT, B.B_REV} | {B.R_REV[k] for k in ('rtx.all', 'rsub.all', 'rnew.all',
                                                                                 'rtot.all', 'rtot.chg')} \
        | {B.B_EMI, B.R_EMI['co2.all'], B.R_EMI['co2.chg'], B.R_EMI['co2.pct']}
    render('06_results', 'Mitigation', values, f'A1:{last}{B.R_EMI["co2.pct"]}', rows=res)

    # 7. Section 2: retail prices before new policies (12 price fuels), history block in light beige
    pr = {4, 5, 6, B.B_PRI} | {B.R_PV0 + B.PVOFF['sp'] + i for i in range(B.NP)} \
        | {B.R_PV0 + B.PVOFF['rpb'] + i for i in range(B.NP)}
    render('07_prices', 'Mitigation', values, f'A1:{L(y1[-1])}{B.R_PV0 + B.PVOFF["rpb"] + B.NP}', rows=pr)


if __name__ == '__main__':
    main()
