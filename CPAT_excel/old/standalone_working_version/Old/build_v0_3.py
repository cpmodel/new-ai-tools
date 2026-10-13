"""Build CPAT_Industry_Kernel_Egypt_v0.3.xlsx from v0.2 (TASK-1: CBAM goods / process emissions block).

Ports the legacy CBAM block (Mitigation rows 12198-12353 carbon tax, 5881-6036 baseline; 156 rows,
columns A..AI) into Mitigation_Industry after each scenario's sector blocks, plus the legacy
'Manual inputs' Process Emissions section (rows 18-47, same row/column positions), carbon price /
ETS trajectories in Data_Prices, MTInputs wiring and Check rows vs stored legacy values.

Run with Excel installed:  python build_v0_3.py
"""
import os
import re
import shutil

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.2.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.3.xlsx")
LEGACY = os.path.join(REPO, "cpat_excel_original", "CPAT 1.0pre_456_NoPropData.xlsb")

L_FIRST, L_LAST = 12198, 12353          # legacy carbon-tax CBAM block (template)
L_BASE_FIRST = 5881                     # legacy baseline CBAM block
NROWS = L_LAST - L_FIRST + 1            # 156
NCOLS = 35                              # A..AI
L_SCEN_ROW = 8504                       # legacy scenario-number cell (C8504)

# v0.2 layout -> v0.3 layout
INSERT_AT = 245                         # baseline CBAM block goes here (after baseline sectors, row 244 blank)
SHIFT = NROWS + 1                       # 156 block rows + 1 blank
BASE_T0 = INSERT_AT                     # 245
CTAX_T0 = 485 + SHIFT + 2               # last carbon-tax sector row 485 -> 642, blank 643, block 644
BASE_SCEN = "R5C10"                     # J5   (baseline iron & steel header, scenario 1)
CTAX_SCEN = "R%dC10" % (247 + SHIFT)    # J404 (carbon tax iron & steel header, scenario 2)

GREEN, BLUE, TAN = 0xDEF1EB, 0xF1E6DC, 0xC4D9DD   # EBF1DE, DCE6F1, DDD9C4 as BGR
TITLE, SECTION = 0x50B000, 0x50D092                # 00B050, 92D050 as BGR

SECTORS = [  # inclusion rows 12206..12210: (MCov name, MTInputs ETS row, MTInputs process row)
    ("MCovMch", 107, 117), ("MCovIrn", 108, 118), ("MCovNfm", 109, 119),
    ("MCovMac", 110, 120), ("MCovCem", 111, 121)]


def col(j):  # 0-based -> letter
    return chr(65 + j) if j < 26 else "A" + chr(65 + j - 26)


def mt(name):
    return 'INDEX(MTInputs!R8C6:R415C6,MATCH("%s",MTInputs!R8C8:R415C8,0))' % name


REF = re.compile(r"(?<![A-Za-z0-9_!'.\]])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z(])")


def port(f, lrow, scen):
    """Translate one legacy R1C1 formula of the template block to the prototype."""
    if not isinstance(f, str) or not f.startswith("="):
        return f
    f = f.replace("LOWER(_Country_code)", "Settings!R3C2")
    f = f.replace('_Country_code="EGY"', 'UPPER(Settings!R3C2)="EGY"')
    if "_" in re.sub(r'"[^"]*"', "", f):
        raise ValueError("unhandled name in %d: %s" % (lrow, f))

    def sub(m):
        rr = m.group(1)
        if rr is None:
            return m.group(0)
        tr = lrow + int(rr[1:-1]) if rr.startswith("[") else int(rr)
        if L_FIRST <= tr <= L_LAST:
            if not rr.startswith("["):
                raise ValueError("absolute in-block ref %d: %s" % (lrow, f))
            return m.group(0)
        if tr == L_SCEN_ROW:
            return scen
        raise ValueError("external ref to row %d in %d: %s" % (tr, lrow, f))
    return REF.sub(sub, f)


def read_legacy(xl):
    wb = xl.Workbooks.Open(LEGACY, 0, True)
    xl.Calculation = -4135
    m = wb.Worksheets("Mitigation")
    blk = m.Range("A%d:AI%d" % (L_FIRST, L_LAST))
    out = {
        "r1c1": [list(r) for r in blk.FormulaR1C1],
        "nf": [[m.Cells(r, c + 1).NumberFormat for c in range(NCOLS)] for r in range(L_FIRST, L_LAST + 1)],
        "bold": [[bool(m.Cells(r, c + 1).Font.Bold) for c in range(1, 4)] for r in range(L_FIRST, L_LAST + 1)],
        "v_ctax": [list(r) for r in m.Range("A%d:AI%d" % (L_FIRST, L_LAST)).Value2],
        "v_base": [list(r) for r in m.Range("A%d:AI%d" % (L_BASE_FIRST, L_BASE_FIRST + NROWS - 1)).Value2],
        "base_r1c1": [list(r) for r in m.Range("A%d:AI%d" % (L_BASE_FIRST, L_BASE_FIRST + NROWS - 1)).FormulaR1C1],
        "cp": {r: list(m.Range("A%d:AI%d" % (r, r)).Value2[0]) for r in (2251, 2252, 8582, 8583)},
    }
    mi = wb.Worksheets("Manual inputs")
    out["mi"] = {r: [mi.Cells(r, c).Formula for c in range(1, 32)] for r in range(18, 48)}
    out["mi_nf"] = {r: [mi.Cells(r, c).NumberFormat for c in range(1, 32)] for r in range(18, 48)}
    out["mi_bold"] = {r: [bool(mi.Cells(r, c).Font.Bold) for c in range(1, 32)] for r in range(18, 48)}
    out["mi_w"] = [mi.Columns(c).ColumnWidth for c in range(1, 32)]
    wb.Close(False)
    return out


def check_template(lg):
    """Baseline and carbon-tax legacy blocks must be identical apart from scenario refs / dead K notes."""
    diffs = []
    for i in range(NROWS):
        for j in range(NCOLS):
            a, b = lg["r1c1"][i][j], lg["base_r1c1"][i][j]
            if a != b and not (j == 7 and "_Country_code" in str(a)) and not (j == 10 and i in (2, 3, 4)):
                if not (i in range(8, 13) and j in (4, 5, 6)):
                    diffs.append((L_FIRST + i, col(j), a, b))
    if diffs:
        raise ValueError("baseline/ctax template mismatch: %s" % diffs[:5])


def block_formulas(lg, t0, scen):
    rows = []
    for i in range(NROWS):
        lrow = L_FIRST + i
        row = []
        for j in range(NCOLS):
            f = lg["r1c1"][i][j]
            if i in range(8, 13) and j in (4, 5, 6):
                f = ""  # inclusion flags, written below
            if i in (2, 3, 4) and j == 10:
                f = ""  # dead legacy notes '=L2727' etc.
            row.append(port(f, lrow, scen) if f != "" else "")
        rows.append(row)

    def put(off, j, f):
        rows[off][j] = f
    years = range(11, 31)  # L..AE
    # title row: scenario helper
    put(0, 8, "Scenario")
    put(0, 9, "=" + scen)
    # carbon price rows (legacy: hard zeros) -> Data_Prices trajectories
    for off, var, unit in ((2, "cptraj", "$/ton CO2 real"), (3, "etstraj", "$/ton CO2 real")):
        put(off, 4, unit)
        put(off, 5, "Data_Prices")
        put(off, 6, '=Settings!R3C2&".mit.%s."&%s' % (var, scen))
        for j in years:
            put(off, j, "=INDEX(Data_Prices!C,MATCH(RC7,Data_Prices!C1,0))")
    put(4, 4, "$/ton CO2")
    for j in years:
        put(4, j, 0)
    for j in [10] + list(years):
        put(5, j, 0)
    # inclusion flags (legacy rows 12206-12210)
    for k, (mcov, ets_row, proc_row) in enumerate(SECTORS):
        off = 8 + k
        put(off, 4, "=IF(AND(%s>1,%s),1,0)" % (scen, mt(mcov)))
        put(off, 5, '=IF(AND(%s>1,%s="Yes"),IF(MTInputs!R%dC6,1,0),0)' % (scen, mt("D_NewETS"), ets_row))
        put(off, 6, "=IF(MTInputs!R%dC6,1,0)" % proc_row)
    # FIX: product-price year header pointed at the carbon tax row (R[-40]C) -> year header R[-13]C
    for j in years:
        put(42, j, "=R[-13]C")
    # FIX: aluminium fuel ER used RC[-7] (shifts column by column) -> $E (RC5)
    for j in years:
        put(114, j, rows[114][j].replace("RC[-7]", "RC5"))
    # FIX: fuel emissions by sector, M..AE were hard-coded -> same formula as L
    for off in range(129, 134):
        for j in years:
            put(off, j, rows[off][11])
    # FIX: missing total of post-policy process emissions
    for j in years:
        put(145, j, "=SUM(R[-5]C:R[-1]C)")
    return rows


def style_block(ws, lg, t0):
    last = t0 + NROWS - 1
    rng = ws.Range("A%d:AI%d" % (t0, last))
    rng.Interior.Pattern = -4142
    rng.Font.Bold = False
    for i in range(NROWS):
        r = t0 + i
        for j in range(NCOLS):
            nf = lg["nf"][i][j]
            if nf != "General":
                ws.Cells(r, j + 1).NumberFormat = nf
        for k in range(3):
            if lg["bold"][i][k]:
                ws.Cells(r, k + 2).Font.Bold = True
    ws.Rows(5).Copy()
    ws.Rows(t0).PasteSpecial(-4122)
    ws.Application.CutCopyMode = False

    def fill(addr, color):
        ws.Range(addr).Interior.Color = color
    fill("L%d:AE%d" % (t0 + 4, t0 + 4), GREEN)          # Total Shadow Price (input, WIP)
    fill("K%d:AE%d" % (t0 + 5, t0 + 5), GREEN)          # Revenues Fund (input, WIP)
    fill("G%d:G%d" % (t0 + 2, t0 + 3), BLUE)
    for off in range(31, 39):                           # production inputs
        fill("F%d:G%d" % (t0 + off, t0 + off), GREEN)
        fill("I%d" % (t0 + off), GREEN)
        fill("N%d" % (t0 + off), TAN)
    for off in range(43, 51):                           # product prices
        fill("E%d" % (t0 + off), BLUE)
        fill("G%d" % (t0 + off), GREEN)
    for off in (126, 129, 130, 131, 132, 133, 134, 140, 141, 142, 143, 144, 145, 155):
        fill("H%d" % (t0 + off), BLUE)


def build_manual_inputs(wb, lg):
    ws = wb.Worksheets.Add(After=wb.Worksheets("Settings"))
    ws.Name = "Manual inputs"
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    for c, w in enumerate(lg["mi_w"], 1):
        ws.Columns(c).ColumnWidth = w
    ws.Range("A1:AE1").Interior.Color = TITLE
    ws.Range("B1").Value = "CPAT Manual inputs"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B1").Font.Color = 0xFFFFFF
    ws.Range("B2").Value = ("Replicates the legacy 'Manual inputs' tab at the same row/column positions. "
                            "Only the Process Emissions section (rows 18-47, used by the CBAM block) is ported; "
                            "other rows are left blank so legacy references port 1:1.")
    ws.Range("B3").Value = "Change only green cells. Row 22 feeds MTInputs F117:F121 (as in legacy)."
    for r in range(18, 48):
        for c in range(1, 32):
            f = lg["mi"][r][c - 1]
            if f == "":
                continue
            if isinstance(f, str) and "Dashboard!$G$20" in f:
                f = f.replace("Dashboard!$G$20", "$E$21")
            ws.Cells(r, c).Formula = f
            nf = lg["mi_nf"][r][c - 1]
            if nf != "General":
                ws.Cells(r, c).NumberFormat = nf
            if lg["mi_bold"][r][c - 1]:
                ws.Cells(r, c).Font.Bold = True
    # legacy Dashboard!G20 switch ("Include non-fuel process emissions in ETS/CT")
    ws.Range("D21").Value = "Include non-fuel process emissions in ETS/CT? (legacy Dashboard!G20)"
    ws.Range("E21").Formula = "=FALSE()"
    ws.Range("E21").Interior.Color = GREEN
    ws.Range("H21:L21").Interior.Color = GREEN
    ws.Range("A18:AE18").Interior.Color = SECTION
    ws.Range("A38:AE38").Font.Bold = True
    # derived EF columns as formulas (legacy stores the same numbers as constants)
    for r in range(30, 38):
        ws.Range("L%d" % r).Formula = "=SUM(H%d:K%d)" % (r, r)
        ws.Range("N%d" % r).Formula = "=L%d+M%d" % (r, r)
        ws.Range("Q%d" % r).Formula = "=N%d+O%d" % (r, r)
        ws.Range("R%d" % r).Formula = "=N%d+O%d*P%d" % (r, r, r)
        ws.Range("Y%d" % r).Formula = "=W%d*(1+X%d)" % (r, r)
        ws.Range("AA%d" % r).Formula = "=R%d-Y%d" % (r, r)
        ws.Range("AB%d" % r).Formula = "=R%d-Z%d" % (r, r)
        for a in ("H%d:K%d", "M%d", "O%d:P%d", "S%d:X%d", "Z%d"):
            ws.Range(a % ((r,) * a.count("%d"))).Interior.Color = GREEN
    ws.Range("E40:F47").Interior.Color = GREEN
    ws.Range("A1").Select()
    return ws


def update_mtinputs(wb):
    ws = wb.Worksheets("MTInputs")
    for k, c in enumerate("HIJKL"):
        ws.Range("F%d" % (117 + k)).Formula = "='Manual inputs'!%s$22" % c
    for r in [39, 40, 41, 42, 43, 84] + list(range(107, 112)) + list(range(117, 122)):
        ws.Range("B%d:H%d" % (r, r)).Interior.Color = GREEN


def update_data_prices(wb, lg):
    ws = wb.Worksheets("Data_Prices")
    first = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
    ws.Rows(first - 1).Copy()
    ws.Rows("%d:%d" % (first, first + 3)).PasteSpecial(-4122)
    wb.Application.CutCopyMode = False
    spec = [(2251, "cptraj", 1, "Carbon price trajectory used"),
            (2252, "etstraj", 1, "Auctioned ETS price (ETS price x auctioned proportion)"),
            (8582, "cptraj", 2, "Carbon price trajectory used"),
            (8583, "etstraj", 2, "Auctioned ETS price (ETS price x auctioned proportion)")]
    for k, (lrow, var, scen, desc) in enumerate(spec):
        r = first + k
        code = "egy.mit.%s.%d" % (var, scen)
        if lg["cp"][lrow][7] != code:
            raise ValueError("legacy code mismatch %s vs %s" % (lg["cp"][lrow][7], code))
        ws.Range("A%d:I%d" % (r, r)).Value = (code, "-", "all", var, scen, "$/ton CO2 real", desc, lrow, code)
        for j in range(10, 35):
            v = lg["cp"][lrow][j]
            ws.Cells(r, j + 1).Value = v if isinstance(v, float) else None
    return first


def update_settings(wb):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype v0.3"
    ws.Rows(17).Insert()
    ws.Range("A17").Value = "Manual inputs"
    ws.Range("C17").Value = ("Legacy 'Manual inputs' Process Emissions section (rows 18-47, same positions): "
                             "CBAM process coverage switches, product emission factors, half-elasticities.")
    ws.Range("C18").Value = ("Kernel sheet in the CPAT Mitigation layout (columns A-AI, years K-AI): 4 CBAM sectors x 2 scenarios, "
                             "each scenario followed by the 156-row CBAM goods block (legacy Mitigation 5881-6036 / 12198-12353). "
                             "All formulas; codes built from components.")
    r = 21
    if ws.Range("A%d" % r).Value != "Data_Prices":
        raise ValueError("Settings layout changed")
    ws.Range("C%d" % r).Value = str(ws.Range("C%d" % r).Value) + " Plus carbon price (cptraj) and auctioned ETS price (etstraj) trajectories for the CBAM block."
    ws.Range("C24").Value = str(ws.Range("C24").Value) + " v0.3: plus CBAM block rows vs legacy (rows from 583)."
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    ws.Rows(last).Copy()
    ws.Rows(last + 1).PasteSpecial(-4122)
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = "v0.3"
    ws.Range("C%d" % (last + 1)).Value = (
        "TASK-1: CBAM goods / process emissions block (legacy Mitigation 12198-12353, 156 rows, cols A-AI) added after each "
        "scenario; 'Manual inputs' rows 18-47; cptraj/etstraj in Data_Prices; MTInputs F117:F121 linked to Manual inputs row 22. "
        "Carbon Tax/ETS rows now read the scenario trajectories (legacy: hard zeros, broken K refs). Legacy WIP bugs fixed: "
        "fuel emissions by sector M..AE hard-coded; aluminium fuel-ER relative ref; missing post-policy process total; "
        "product-price year header. Kernel rows unchanged (regression diff 0).")
    ws.Range("C%d" % (last + 1)).WrapText = False


CHANGED = {  # template offsets whose values differ from legacy by design: (scenarios, reason)
    2: ((2,), "carbon tax wired to cptraj (legacy hard 0)"),
    **{o: ((2,), "uses wired carbon tax") for o in range(54, 62)},
    **{o: ((2,), "uses wired carbon tax") for o in range(75, 83)},
    **{o: ((1, 2), "fix: legacy M..AE hard-coded") for o in range(129, 135)},
    **{o: ((2,), "uses wired carbon tax (revenue base: see notes)") for o in range(147, 156)},
}
CHECK_OFFSETS = ([2, 3] + list(range(31, 39)) + list(range(43, 51)) + list(range(54, 62)) + list(range(65, 73))
                 + list(range(75, 83)) + list(range(85, 94)) + list(range(96, 105)) + list(range(107, 116))
                 + list(range(118, 127)) + list(range(129, 135)) + list(range(140, 146)) + list(range(147, 156)))


def build_check(wb, lg):
    ws = wb.Worksheets("Check")
    r0 = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 3
    ws.Range("A%d" % r0).Value = ("CBAM block vs values stored in CPAT_1_0pre_456 (Egypt): legacy Mitigation rows "
                                  "5881-6036 (baseline) and 12198-12353 (carbon tax); kernel = Mitigation_Industry same block row")
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 3
    ws.Rows(5).Copy()
    ws.Rows(hdr).PasteSpecial(-4122)
    wb.Application.CutCopyMode = False
    ws.Range("A%d:J%d" % (hdr, hdr)).Value = ("Legacy row", "Block section", "Item", "Scenario", "Series",
                                              "max |diff|", "", "", "Expected", "Reason")
    for j in range(10, 35):
        ws.Cells(hdr, j + 1).Value = 2011 + j
    r = hdr + 1
    for scen, t0, lfirst, vals in ((1, BASE_T0, L_BASE_FIRST, lg["v_base"]), (2, CTAX_T0, L_FIRST, lg["v_ctax"])):
        section = ""
        for off in CHECK_OFFSETS:
            # nearest section header label (column C) above
            for o in range(off, -1, -1):
                if lg["r1c1"][o][2] not in ("", None) and not str(lg["r1c1"][o][2]).startswith("="):
                    section = lg["r1c1"][o][2]
                    break
            item = lg["r1c1"][off][3] or lg["r1c1"][off][6] or ""
            exp, why = "same", ""
            if off in CHANGED and scen in CHANGED[off][0]:
                exp, why = "changed", CHANGED[off][1]
            lrow = lfirst + off
            krow = t0 + off
            for k, series in enumerate(("CPAT", "kernel", "diff")):
                rr = r + k
                ws.Range("A%d:E%d" % (rr, rr)).Value = ("Mitigation!%d" % lrow, section, item, scen, series)
            ws.Range("I%d:J%d" % (r + 2, r + 2)).Value = (exp, why)
            for j in range(10, 31):
                v = vals[off][j]
                ws.Cells(r, j + 1).Value = v if isinstance(v, float) else None
            ws.Range("K%d:AE%d" % (r + 1, r + 1)).Formula = "=IF(ISBLANK(Mitigation_Industry!K%d),\"\",Mitigation_Industry!K%d)" % (krow, krow)
            ws.Range("K%d:AE%d" % (r + 2, r + 2)).Formula = "=IF(ISNUMBER(K%d),IF(ISNUMBER(K%d),K%d-K%d,999999),0)" % (r, r + 1, r + 1, r)
            ws.Range("F%d" % (r + 2)).Formula = "=MAX(MAX(L%d:AE%d),-MIN(L%d:AE%d))" % (r + 2, r + 2, r + 2, r + 2)
            r += 4
    last = r - 1
    ws.Range("A%d" % (r0 + 1)).Value = "Max |kernel - CPAT| CBAM block, rows replicated unchanged"
    first = hdr + 1
    ws.Range("D%d" % (r0 + 1)).Formula = '=SUMPRODUCT(MAX((I%d:I%d="same")    *F%d:F%d))' % (first, last, first, last)
    ws.Range("A%d" % (r0 + 2)).Value = "Max |kernel - CPAT| CBAM block, rows changed by design (reason in column J)"
    ws.Range("D%d" % (r0 + 2)).Formula = '=SUMPRODUCT(MAX((I%d:I%d="changed")*F%d:F%d))' % (first, last, first, last)
    ws.Range("D%d:D%d" % (r0 + 1, r0 + 2)).Interior.Color = 0xCCF2FF  # FFF2CC
    return r0, last


def main():
    if os.path.exists(DST):
        raise SystemExit("%s exists; never overwrite a saved version" % DST)
    shutil.copyfile(SRC, DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    try:
        lg = read_legacy(xl)
        check_template(lg)
        wb = xl.Workbooks.Open(DST, 0, False)
        xl.Calculation = -4135
        build_manual_inputs(wb, lg)
        update_mtinputs(wb)
        update_data_prices(wb, lg)
        ws = wb.Worksheets("Mitigation_Industry")
        ws.Rows("%d:%d" % (INSERT_AT, INSERT_AT + SHIFT - 1)).Insert()
        for t0, scen in ((BASE_T0, BASE_SCEN), (CTAX_T0, CTAX_SCEN)):
            if ws.Range("J%s" % scen.split("C")[0][1:]).Value not in (1, 2):
                raise ValueError("scenario helper cell not found")
            style_block(ws, lg, t0)
            arr = block_formulas(lg, t0, scen)
            ws.Range("A%d:AI%d" % (t0, t0 + NROWS - 1)).FormulaR1C1 = tuple(tuple(r) for r in arr)
        if ws.Range("B%d" % (INSERT_AT + SHIFT)).Value is None or "Carbon tax" not in ws.Range("B%d" % (INSERT_AT + SHIFT)).Value:
            raise ValueError("carbon tax section header not where expected")
        build_check(wb, lg)
        update_settings(wb)
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    print("saved", DST)


if __name__ == "__main__":
    main()
