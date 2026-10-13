"""Build CPAT_Industry_Kernel_Egypt_v0.8.xlsx from v0.7 (Task K: CBAM coverage and intensity metrics).

New output rows in both CBAM blocks (spare rows o39-o41, o51-o52; the 156-row block size is unchanged), all with
G = cbam.tot and output codes in H (egy.mit.<var>.cbam.tot.<s>):

  o39 cbcov    CBAM coverage: share of pre-policy CBAM-sector emissions that is priced (covered EF x sector inclusion
               flag, fuel o54 and process o65). EgyptResultsInitial Table 2 "CBAM coverage" (100% / 44%).
  o40 cbint    CBAM-sector emission intensity, post-policy, tCO2e per t product (production-weighted).
  o41 cbintch  Intensity change vs pre-policy at the same output (= emissions change while output is exogenous).
               Table 2 "Emission intensity reduction, CBAM sectors".
  o51 cbintchx Intensity change weighted by 2024 EU exports (o31-o38 I): change in emissions embedded in EU exports
               per unit exported; the CBAM-relevant intensity (Task L obligations).
  o52 cbcovx   Coverage weighted by EU exports (share of export-embedded emissions that is priced; Task L credit).

Pre-policy emissions = o85-o92 (fuel) + o96-o103 (process); post-policy adds o107-o114 and o118-o125.
Coverage is legal scope (flags), not conditional on the price being positive in a given year.
No existing value changes.

Run with Excel installed:  python build_v0_8.py
"""
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLOCKS, BLUE, DATA_COLS, REVIEW, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.7.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.7.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.8.xlsx")
CODE = 'Settings!R3C2&".mit.%s."&RC7&"."&%s'
ROWS = (39, 40, 41, 51, 52)
PCT = "0.0%"


def empty(ws, r, c1=1, c2=31):
    return all(ws.Cells(r, c).Formula in ("", None) for c in range(c1, c2 + 1))


def metrics(t0):
    """[(offset, var, label, unit, R1C1 formula, number format)]"""
    R = lambda o: "R%d" % (t0 + o)
    rng8 = lambda o, c="C": "%s%s:%s%s" % (R(o), c, R(o + 7), c)
    prod, x = rng8(31), rng8(31, "C9")
    fpre, ppre, fer, per = rng8(85), rng8(96), rng8(107), rng8(118)
    fsh, psh = rng8(85, "C9"), rng8(85, "C10")
    fflag = "((%s+%s)>0)" % (rng8(54, "C6"), rng8(54, "C7"))
    pflag = "((%s+%s)>0)" % (rng8(65, "C6"), rng8(65, "C7"))
    pre = "(SUM(%s)+SUM(%s))" % (fpre, ppre)
    cov_num = "(SUMPRODUCT(%s*%s*%s)+SUMPRODUCT(%s*%s*%s))" % (fpre, fsh, fflag, ppre, psh, pflag)
    unit = "(%s+(%s=0))" % (prod, prod)                    # avoids 0/0 for products with no output (BF-BOF)
    xpre = "SUMPRODUCT(%s*(%s+%s)/%s)" % (x, fpre, ppre, unit)
    xpost = "SUMPRODUCT(%s*(%s+%s+%s+%s)/%s)" % (x, fpre, fer, ppre, per, unit)
    xcov = "SUMPRODUCT(%s*(%s*%s*%s+%s*%s*%s)/%s)" % (x, fpre, fsh, fflag, ppre, psh, pflag, unit)
    emis = R(135) + "C"
    return [
        (39, "cbcov", "CBAM coverage: priced share of pre-policy CBAM-sector emissions (covered EF x inclusion flags)",
         "share", "=IF(%s=0,0,%s/%s)" % (pre, cov_num, pre), PCT),
        (40, "cbint", "CBAM-sector emission intensity, post-policy (production-weighted)", "tCO2e/t",
         "=IF(SUM(%s)=0,0,%s*1000/SUM(%s))" % (prod, emis, prod), "0.000"),
        (41, "cbintch", "CBAM-sector intensity change vs pre-policy, same output (Table 2 intensity reduction)",
         "share", "=IF(%s=0,0,%s/%s-1)" % (pre, emis, pre), PCT),
        (51, "cbintchx", "CBAM intensity change weighted by 2024 EU exports (o31-o38 I): embedded emissions per unit "
         "exported", "share", "=IF(%s=0,0,%s/%s-1)" % (xpre, xpost, xpre), PCT),
        (52, "cbcovx", "CBAM coverage weighted by 2024 EU exports: priced share of export-embedded emissions",
         "share", "=IF(%s=0,0,%s/%s)" % (xpre, xcov, xpre), PCT),
    ]


def update_block(ws, t0, scen):
    for off in ROWS:
        if not empty(ws, t0 + off):
            raise ValueError("spare block row o%d (row %d) not empty" % (off, t0 + off))
    checks = ((30, "I", "2024 EUexp kt"), (84, "I", "Covered share, fuel (o54 E / F)"),
              (84, "J", "Covered share, process (o65 E / G)"), (135, "G", "cbam.tot"))
    for off, c, v in checks:
        if ws.Range("%s%d" % (c, t0 + off)).Value != v:
            raise ValueError("block layout changed at o%d %s" % (off, c))
    for off in (85, 96, 107, 118):
        if ws.Range("D%d" % (t0 + off)).Value != "DRI-EAF steel" or ws.Range("D%d" % (t0 + off + 7)).Value is None:
            raise ValueError("product rows moved at o%d" % off)
    for off, var, label, unit, f, fmt in metrics(t0):
        r = t0 + off
        copy_formats(ws.Rows(t0 + 135), ws.Rows(r))
        ws.Application.CutCopyMode = False
        ws.Range("D%d:E%d" % (r, r)).Value = (label, unit)
        ws.Range("G%d" % r).Value = "cbam.tot"
        ws.Range("H%d" % r).FormulaR1C1 = "=" + CODE % (var, scen)
        ws.Range("H%d" % r).Interior.Color = BLUE
        for c in DATA_COLS:
            ws.Cells(r, c).FormulaR1C1 = f
        ws.Range("%s%d:%s%d" % (col(DATA_COLS[0]), r, col(DATA_COLS[-1]), r)).NumberFormat = fmt
    ws.Range("C%d" % (t0 + 39)).Value = "CBAM coverage and intensity metrics (Task K)"
    ws.Range("C%d" % (t0 + 51)).Value = "CBAM metrics, EU-export weighted (Task K)"


def update_scenarios(wb):
    ws = wb.Worksheets("Scenarios")
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task K (stream 1, v0.8): CBAM metric codes (block column H, G = cbam.tot): cbcov (o39, priced share of "
        "CBAM-sector emissions), cbint (o40, tCO2e/t), cbintch (o41, intensity change vs pre-policy), cbintchx (o51, "
        "EU-export-weighted intensity change), cbcovx (o52, EU-export-weighted coverage). Fuel/process split of the "
        "change: emisfc..emisno (o13-o16) vs pre-policy o85-o103.")


def update_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task K (v0.8): CBAM coverage and intensity metrics (expected 0)"
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 2
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:J%d" % (hdr, hdr)).Value = ("Item", "", "", "Scenario", "Series", "max |diff|", "", "", "Expected", "")
    for j in range(10, 35):
        ws.Cells(hdr, j + 1).Value = 2011 + j
    M = "Mitigation_Industry!"
    r = hdr + 1
    for t0, scen in BLOCKS:
        s = 1 if t0 == BLOCKS[0][0] else 2
        rows = []
        ws.Range("A%d" % r).Value = "cbintch - (sum of EF-category rows o13-o16 / pre-policy - 1)"
        ws.Range("D%d:E%d" % (r, r)).Value = (s, "diff")
        for c in DATA_COLS:
            L = col(c)
            pre = "(SUM(%s%s%d:%s%d)+SUM(%s%s%d:%s%d))" % (M, L, t0 + 85, L, t0 + 92, M, L, t0 + 96, L, t0 + 103)
            ws.Cells(r, c).Formula = "=%s%s%d-IF(%s=0,0,SUM(%s%s%d:%s%d)/%s-1)" % (
                M, L, t0 + 41, pre, M, L, t0 + 13, L, t0 + 16, pre)
        ws.Range("F%d" % r).Formula = "=MAX(MAX(L%d:AE%d),-MIN(L%d:AE%d))" % (r, r, r, r)
        ws.Range("I%d" % r).Value = 0
        r += 1
        for item, f in (
                ("Coverage cbcov / cbcovx outside [0,1] (count of years)",
                 "=SUMPRODUCT((%s$L$%d:$AE$%d<-1E-12)+(%s$L$%d:$AE$%d>1+1E-12))+SUMPRODUCT((%s$L$%d:$AE$%d<-1E-12)"
                 "+(%s$L$%d:$AE$%d>1+1E-12))" % (M, t0 + 39, t0 + 39, M, t0 + 39, t0 + 39,
                                                 M, t0 + 52, t0 + 52, M, t0 + 52, t0 + 52)),
                ("Intensity change cbintch / cbintchx above 0 or below -1 (count of years)",
                 "=SUMPRODUCT((%s$L$%d:$AE$%d>1E-12)+(%s$L$%d:$AE$%d<-1))+SUMPRODUCT((%s$L$%d:$AE$%d>1E-12)"
                 "+(%s$L$%d:$AE$%d<-1))" % (M, t0 + 41, t0 + 41, M, t0 + 41, t0 + 41,
                                            M, t0 + 51, t0 + 51, M, t0 + 51, t0 + 51))):
            ws.Range("A%d" % r).Value = item
            ws.Range("D%d" % r).Value = s
            ws.Range("F%d" % r).Formula = f
            ws.Range("I%d" % r).Value = 0
            r += 1
    t0 = BLOCKS[0][0]
    ws.Range("A%d" % r).Value = "Baseline block (scenario 1): cbcov, cbintch and cbintchx not 0 (max abs)"
    ws.Range("D%d" % r).Value = 1
    ws.Range("F%d" % r).Formula = ("=MAX(SUMPRODUCT(ABS(%s$L$%d:$AE$%d)),SUMPRODUCT(ABS(%s$L$%d:$AE$%d)),"
                                   "SUMPRODUCT(ABS(%s$L$%d:$AE$%d)))" % (M, t0 + 39, t0 + 39, M, t0 + 41, t0 + 41,
                                                                         M, t0 + 51, t0 + 51))
    ws.Range("I%d" % r).Value = 0
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (Task K)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (hdr + 1, r)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    return r0 + 1


def update_settings(wb, chk_row):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype v0.8"
    if ws.Range("A25").Value != "Check":
        raise ValueError("Settings layout changed")
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " v0.8: Task K section (row %d)." % chk_row
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = "v0.8"
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task K: CBAM metrics in block spare rows: cbcov (o39), cbint (o40), cbintch (o41), EU-export-weighted "
        "cbintchx (o51) and cbcovx (o52). New rows only; no existing value changes.")
    ws.Range("C%d" % (last + 1)).WrapText = False


def main():
    if os.path.exists(DST):
        raise SystemExit("%s exists; never overwrite a saved version" % DST)
    shutil.copyfile(SRC, DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    try:
        wb = xl.Workbooks.Open(DST, 0, False)
        xl.Calculation = -4135
        if wb.Worksheets("Settings").Range("B10").Value != "LEGACY":
            raise ValueError("source must be saved with Settings!B10 = LEGACY")
        ws = wb.Worksheets("Mitigation_Industry")
        update_scenarios(wb)
        for t0, scen in BLOCKS:
            update_block(ws, t0, scen)
        chk = update_check(wb)
        update_settings(wb, chk)
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
    except Exception:
        for b in list(xl.Workbooks):
            b.Close(False)
        xl.Quit()
        os.remove(DST)
        raise
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
    print("saved", DST)


if __name__ == "__main__":
    main()
