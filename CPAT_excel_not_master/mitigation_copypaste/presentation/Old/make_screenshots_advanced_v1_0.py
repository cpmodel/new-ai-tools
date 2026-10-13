"""Screenshots of CPAT-AI-Mitigation-MVP-v1.03.xlsm for the advanced-features deck (presentation/img/1x_*.png).

As make_screenshots_v1_0.py: the workbook is recalculated in LibreOffice (LAMBDAs expanded), one sheet is turned into
a values copy, rows / columns / print area are set, and LibreOffice prints it to PDF, which is cut to a PNG. Also
writes presentation/img/advanced_data.json with the numbers the deck charts (from StoredResults).

    python presentation/make_screenshots_advanced_v1_0.py
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import column_index_from_string as CI

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(1, os.path.join(os.path.dirname(HERE), "Old"))     # v1.04 builder archived in Old/
import build_v1_04 as B                      # noqa: E402
import check_v1_04 as C                      # noqa: E402

IMG = os.path.join(HERE, 'img')
LO_PROFILE = 'file://' + os.path.join(tempfile.gettempdir(), 'lo_shots_profile')


def soffice(args, outdir):
    subprocess.run(['soffice', f'-env:UserInstallation={LO_PROFILE}', '--headless'] + args + ['--outdir', outdir],
                   check=True, capture_output=True, timeout=900)


def crop(png):
    from PIL import Image, ImageChops
    im = Image.open(png).convert('RGB')
    box = ImageChops.difference(im, Image.new('RGB', im.size, (255, 255, 255))).getbbox()
    if box:
        im = im.crop((max(box[0] - 8, 0), max(box[1] - 8, 0), box[2] + 8, box[3] + 8))
    im.save(png)


def render(name, sheet, values, area, rows, hide_cols=(), widths=None, wrap_rows=(), heights=None):
    wb = openpyxl.load_workbook(B.OUT)
    ws = wb[sheet]
    vs = values[sheet]
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                c.value = vs[c.coordinate].value
    for other in [w for w in wb.worksheets if w.title != sheet]:
        wb.remove(other)
    wb.defined_names.clear()
    ws.sheet_format.outlineLevelRow = 0
    for r in range(1, ws.max_row + 2):
        d = ws.row_dimensions[r]
        d.outlineLevel, d.collapsed, d.hidden = 0, False, r not in rows
    for col in hide_cols:
        ws.column_dimensions[col].hidden = True
    for col, w in (widths or {}).items():
        ws.column_dimensions[col].width = w
    for r in wrap_rows:
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical='top')
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
    values = C.recalc_wb(C.variant(openpyxl.load_workbook(B.OUT), 'expand'))

    # 1. ScenarioCompare: header and the level block (selected indicators), hiding the code-stem column
    keep = ['mit.ener.all.all.e', 'mit.co2.all.all.e', 'mit.co2.tra.all.e', 'mit.co2.ind.all.e', 'mit.rtot.all.all',
            'mit.cptraj', 'mit.ets.p', 'mit.co2.ets.all.all', 'mit.co2.cap.all.all', 'mit.atp.rod.gso.e']
    stems = [s for s, _f in B.CMP_STEMS]
    rows = {4, 6, 7} | {9 + stems.index(s) for s in keep}
    last = openpyxl.utils.get_column_letter(4 + 3)
    render('11_compare', 'ScenarioCompare', values, f'A1:{last}{9 + len(stems)}', rows, hide_cols=['A'],
           widths={'B': 44, 'C': 11, 'D': 15, 'E': 15, 'F': 15, 'G': 15}, wrap_rows=[7], heights={7: 52})

    # 2. MTInputs: scenario columns J:N with Run? flags, names and a few changed inputs
    n = B.MT_NAMES
    keys = ['CPIntro', 'CPLevelStart', 'CPLevelTarget', 'CPOutro', 'D_Feb_Level_Start_Trans',
            'D_Feb_Level_Target_Trans', 'D_NewETS', 'D_ETSChangeRelStart', 'D_ETSChangeRelTarget']
    rows = {4, 5, 6, 55, 60, 83} | {n[k] for k in keys}        # 55, 60, 83: sub-headings
    render('12_definitions', 'MTInputs', values, 'B1:N95', rows, hide_cols=list('CDEFGHI'),
           widths={'B': 40, 'J': 17, 'K': 17, 'L': 17, 'M': 17, 'N': 17}, wrap_rows=[4, 6], heights={6: 64, 4: 26})

    # chart data from StoredResults
    st, names = C.stored_blocks(values['StoredResults'])
    years = list(range(B.BASE_YEAR, B.LAST_YEAR + 1))
    data = {'names': {str(k): v for k, v in names.items()},
            'co2_2030': {str(k): st[k]['mit.co2.all.all.e'][2030 - B.BASE_YEAR] for k in st},
            'years': years, 'ets_p': st[5]['mit.ets.p'], 'co2_ets': st[5]['mit.co2.ets.all.all'],
            'co2_cap': st[5]['mit.co2.cap.all.all']}
    with open(os.path.join(IMG, 'advanced_data.json'), 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=1)
    print('wrote advanced_data.json', data['co2_2030'])


if __name__ == '__main__':
    main()
