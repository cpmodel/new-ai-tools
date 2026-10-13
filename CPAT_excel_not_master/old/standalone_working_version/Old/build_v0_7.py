"""Build CPAT_Industry_Kernel_Egypt_v0.7.xlsx from v0.6 (Task C: process emissions priced with pptraj).

pptraj (Scenarios row 44, Data_Prices egy.mit.pptraj.<s>) = active bundle's process price path (Scenarios column H)
x process-covered flag (column G). Up to v0.6 the CBAM block priced non-fuel process emissions with the energy
carbon tax o2 (cptraj). From v0.7:

  * Spare block row o63 = Process carbon price, lookup of egy.mit.pptraj.<s> (both CBAM blocks).
  * Every process-price use reads o63 instead of o2 (ETS part o3 unchanged):
      o65-o72 process price increase, o118-o125 process ER, o138 revp, process part of o147-o154,
      o15-o16 post-policy np / no emissions.
  * Fuel CO2 (o54-o61, o137, fuel part of o147-o154, o13-o14) still uses o2.

All current bundles use the same path for energy and process (H = D), and the sector process flags already carry
the process-covered switch, so all values are identical to v0.6. A bundle can now price process emissions on a
different path (e.g. a CBAM-aligned fee) without touching the energy price.

Run with Excel installed:  python build_v0_7.py
"""
import datetime
import os
import re
import shutil

import win32com.client as win32

from build_v0_4 import BLOCKS, DATA_COLS, REVIEW, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.6.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.6.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.7.xlsx")
OPP = 63                                         # block offset of the new process price row
LOOKUP = "=INDEX(Data_Prices!C,MATCH(RC7,Data_Prices!C1,0))"
Y2030 = 20                                       # column T (DATA_COLS L..AE = 2022..2041)


def empty(ws, r, c1=1, c2=31):
    return all(ws.Cells(r, c).Formula in ("", None) for c in range(c1, c2 + 1))


def formulas(t0, pp):
    """{(offset, column): R1C1 formula} for every process-price use; pp = row reference of the process price."""
    R = lambda o: "R%d" % (t0 + o)
    p, e = R(2) + "C", R(3) + "C"
    q = pp + "C"
    rng8 = lambda o, c="C": "%s%s:%s%s" % (R(o), c, R(o + 7), c)
    pproc = "(%s*%s+%s*%s)" % (rng8(65, "C6"), q, rng8(65, "C7"), e)
    out = {}
    for c in DATA_COLS:
        for k in range(8):
            out[(65 + k, c)] = "=(RC6*%s+RC7*%s)*RC5" % (q, e)
            out[(118 + k, c)] = "=R[-22]C*R[-33]C10*(EXP(-(RC5)*(R[-53]C6*%s+R[-53]C7*%s))-1)" % (q, e)
            out[(147 + k, c)] = ("=(R[-62]C+R[-40]C)*R[-62]C9*(R[-93]C6*%s+R[-93]C7*%s)"
                                 "+(R[-51]C*R[-62]C10+R[-29]C)*(R[-82]C6*%s+R[-82]C7*%s)" % (p, e, q, e))
        out[(138, c)] = "=SUMPRODUCT((%s*%s+%s)*%s)" % (rng8(96), rng8(85, "C10"), rng8(118), pproc)
        for off, ec in ((15, 10), (16, 11)):
            ef = rng8(19, "C%d" % ec)
            out[(off, c)] = ("=SUMPRODUCT(%s*(%s>0)*%s*(1+%sC%d*(EXP(-%s*%s)-1)))/1000"
                             % (rng8(31), ef, ef, R(28), ec, rng8(118, "C5"), pproc))
    return out


def norm(f):
    return re.sub(r"\s", "", f)


def update_block(ws, t0, scen):
    if not empty(ws, t0 + OPP) or not empty(ws, t0 + 62):
        raise ValueError("spare block rows o62/o63 (row %d) not empty" % (t0 + OPP))
    if ws.Range("D%d" % (t0 + 2)).Value != "Carbon Tax" or ws.Range("D%d" % (t0 + 3)).Value != "ETS Price (Auctioned)":
        raise ValueError("block price rows o2/o3 moved")
    for c in DATA_COLS:
        if ws.Cells(t0 + 2, c).FormulaR1C1 != LOOKUP:
            raise ValueError("o2 lookup changed at column %s" % col(c))
    old = formulas(t0, "R%d" % (t0 + 2))
    xl = ws.Application
    a1 = lambda f, cell: norm(xl.ConvertFormula(f, -4150, 1, 4, cell)).replace("$", "")
    for (off, c), f in old.items():
        cell = ws.Cells(t0 + off, c)
        got = cell.Formula
        if norm(got).replace("$", "") != a1(f, cell):
            raise ValueError("o%d %s%d formula differs from v0.6:\n  %s\n  %s" % (off, col(c), t0 + off, got, a1(f, cell)))

    r = t0 + OPP
    copy_formats(ws.Rows(t0 + 2), ws.Rows(r))
    ws.Application.CutCopyMode = False
    srow = int(re.match(r"R(\d+)C10$", scen).group(1))
    ws.Range("D%d:F%d" % (r, r)).Value = ("Process carbon price", "$/ton CO2 real", "Data_Prices")
    ws.Range("G%d" % r).Formula = '=Settings!$B$3&".mit.pptraj."&$J$%d' % srow
    ws.Range("I%d" % r).Value = ("pptraj = process path x process flag (Scenarios H, G); prices np + no in "
                                 "o65-o72, o118-o125, o138, o147-o154, o15-o16 (Task C)")
    for c in DATA_COLS:
        ws.Cells(r, c).FormulaR1C1 = LOOKUP
    for (off, c), f in formulas(t0, "R%d" % r).items():
        ws.Cells(t0 + off, c).FormulaR1C1 = f
    ws.Range("D%d" % (t0 + 138)).Value = ("Revenues, covered non-fuel process np + no at the process price o63 "
                                          "(additional to energy carbon revenue)")


def update_scenarios(wb):
    ws = wb.Worksheets("Scenarios")
    if ws.Range("G42").Value != "egy.mit.cptraj.2" or ws.Range("G44").Value != "egy.mit.pptraj.2":
        raise ValueError("Scenarios section 4 layout changed")
    ws.Range("H42").Value = "kernel via nce; CBAM block o2 (fuel CO2 only from v0.7)"
    ws.Range("H44").Value = "CBAM block o63 (Task C, process price); Stream 1 L: CBAM obligations"
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task C (stream 1, v0.7): CBAM process emissions (np + no) are priced at pptraj (block o63), fuel CO2 "
        "(fc + fp) at cptraj (o2); ETS (o3) applies to both via the sector ETS flags. To give process emissions their "
        "own path, set the bundle's 'Process price path' (column H); 'Process emissions covered' (G) switches it off. "
        "Reporting convention: all bundles price 20 $/t in 2030 (checked).")


def update_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task C (v0.7): process emissions priced at pptraj (expected 0)"
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
        ws.Range("A%d" % r).Value = "Block process price o63 - Data_Prices pptraj"
        ws.Range("D%d:E%d" % (r, r)).Value = (s, "diff")
        for c in DATA_COLS:
            L = col(c)
            ws.Cells(r, c).Formula = ('=%s%s%d-INDEX(Data_Prices!%s:%s,MATCH("egy.mit.pptraj.%d",Data_Prices!$A:$A,0))'
                                      % (M, L, t0 + OPP, L, L, s))
        ws.Range("F%d" % r).Formula = "=MAX(MAX(L%d:AE%d),-MIN(L%d:AE%d))" % (r, r, r, r)
        ws.Range("I%d" % r).Value = 0
        r += 1
    t0 = BLOCKS[1][0]
    T = col(Y2030)
    ws.Range("A%d" % r).Value = ("2030 process price - energy price x process flag (non-LEGACY bundles; "
                                 "reporting convention 20 $/t in 2030)")
    ws.Range("D%d" % r).Value = 2
    ws.Range("F%d" % r).Formula = ('=IF(Scenarios!$B$27="LEGACY",0,ABS(%s%s%d-Scenarios!$G$27*%s%s%d))'
                                   % (M, T, t0 + OPP, M, T, t0 + 2))
    ws.Range("I%d" % r).Value = 0
    r += 1
    ws.Range("A%d" % r).Value = "2030 energy price <> 20 $/t (non-LEGACY bundles)"
    ws.Range("D%d" % r).Value = 2
    ws.Range("F%d" % r).Formula = '=IF(Scenarios!$B$27="LEGACY",0,ABS(%s%s%d-20))' % (M, T, t0 + 2)
    ws.Range("I%d" % r).Value = 0
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| (Task C)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (hdr + 1, r)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    return r0 + 1


def update_settings(wb, chk_row):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype v0.7"
    if ws.Range("A25").Value != "Check":
        raise ValueError("Settings layout changed")
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " v0.7: Task C section (row %d)." % chk_row
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = "v0.7"
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task C: CBAM process emissions (np + no) priced at pptraj (new block row o63) in the process price rows, "
        "process ER, revp, product revenues and np/no emissions; fuel CO2 stays on cptraj (o2). All values identical "
        "to v0.6 (every bundle uses the same energy and process path).")
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
