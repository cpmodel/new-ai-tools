"""Build CPAT_Industry_Kernel_Egypt_v0.5.xlsx from v0.4 (Task A: separate fuel/process output codes, fix revenues).

CBAM block (both scenarios), stream-1 rows only; the 156-row block size is unchanged:
  * Output codes (column H) made unique: fuel post-policy emissions cpef.ind.<sec> / emisf.cbam.tot, process
    cpep.ind.<sec> / emisp.cbam.tot, fuel ER emrf.cbam.tot (o115, previously uncoded), process ER emrp.cbam.tot (o126).
  * Spare rows o135-o138: combined totals emis.cbam.tot (fuel + process post-policy), emr.cbam.tot (fuel + process
    additional ER), and revenue split revf.cbam.tot (fuel CO2) / revp.cbam.tot (non-fuel process).
  * Revenue rows o147-o154 fixed: post-policy fuel emissions x fuel inclusion (o54-o61) + post-policy process
    emissions x PROCESS inclusion (o65-o72). Legacy used pre-policy process emissions x fuel inclusion and omitted
    fuel revenue. o155 Total Revenues (rev.cbam.tot, o5 fund base) = revf + revp.

Fuel revenue is a memo item: the same CO2 is priced in the kernel energy wedge (nce), so main-CPAT totals must
add only revp (Task M). Non-revenue values are identical to v0.4.

Run with Excel installed:  python build_v0_5.py
"""
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLOCKS, BLUE, DATA_COLS, REVIEW, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.4.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.4.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.5.xlsx")
L_FIRST = 12198  # legacy Mitigation row of the carbon-tax block (o0)

SECTORS = ("mch", "irn", "nfm", "mac", "cem")
CODE = 'Settings!R3C2&".mit.%s."&RC7&"."&%s'   # % (variable, scenario-cell)


def update_block(ws, t0, scen):
    def lab(off):
        return ws.Range("D%d" % (t0 + off)).Value

    expect = {134: "Total", 145: "Total", 155: "Total Revenues"}
    for off, v in expect.items():
        if lab(off) != v:
            raise ValueError("block o%d label %r at row %d" % (off, lab(off), t0 + off))
    for off in (115, 135, 136, 137, 138):
        if any(ws.Cells(t0 + off, c).Formula not in ("", None) for c in range(1, 4)) or \
                any(ws.Cells(t0 + off, c).Formula not in ("", None) for c in range(4, 9)):
            raise ValueError("block row o%d (row %d) expected empty labels/codes" % (off, t0 + off))
        if off != 115 and any(ws.Cells(t0 + off, c).Formula not in ("", None) for c in range(9, 36)):
            raise ValueError("spare block row o%d (row %d) not empty" % (off, t0 + off))

    def code(off, var):
        ws.Range("H%d" % (t0 + off)).FormulaR1C1 = "=" + CODE % (var, scen)
        ws.Range("H%d" % (t0 + off)).Interior.Color = BLUE

    # 1. unique output codes
    ws.Range("D%d" % (t0 + 115)).Value = "Total"
    ws.Range("G%d" % (t0 + 115)).Value = "cbam.tot"
    code(115, "emrf")
    code(126, "emrp")
    for k, s in enumerate(SECTORS):
        if ws.Range("G%d" % (t0 + 129 + k)).Value != "ind." + s or ws.Range("G%d" % (t0 + 140 + k)).Value != "ind." + s:
            raise ValueError("sector order changed at o%d" % (129 + k))
        code(129 + k, "cpef")
        code(140 + k, "cpep")
    code(134, "emisf")
    code(145, "emisp")

    # 2. spare rows o135-o138: combined totals and revenue split
    copy_formats(ws.Rows(t0 + 134), ws.Rows("%d:%d" % (t0 + 135, t0 + 138)))
    ws.Application.CutCopyMode = False
    ws.Range("D%d:D%d" % (t0 + 135, t0 + 138)).Font.Bold = False
    p, e = "R%dC" % (t0 + 2), "R%dC" % (t0 + 3)

    def sp(em, er, fl):
        a = lambda o: "R%d" % (t0 + o)
        return ("=SUMPRODUCT((%sC:%sC+%sC:%sC)*(%sC6:%sC6*%s+%sC7:%sC7*%s))"
                % (a(em), a(em + 7), a(er), a(er + 7), a(fl), a(fl + 7), p, a(fl), a(fl + 7), e))
    new = {
        135: ("Total, fuel + non-fuel process (post-policy)", "MtCO2e", "emis", "=R%dC+R%dC" % (t0 + 134, t0 + 145)),
        136: ("Total additional ER, fuel + non-fuel process", "MtCO2e", "emr", "=R%dC+R%dC" % (t0 + 115, t0 + 126)),
        137: ("Revenues, fuel CO2 (memo: same CO2 is priced in kernel energy via nce)", "USD million", "revf",
              sp(85, 107, 54)),
        138: ("Revenues, non-fuel process (additional to energy carbon revenue)", "USD million", "revp",
              sp(96, 118, 65)),
    }
    for off, (label, unit, var, f) in new.items():
        r = t0 + off
        ws.Range("D%d:E%d" % (r, r)).Value = (label, unit)
        ws.Range("G%d" % r).Value = "cbam.tot"
        code(off, var)
        for c in DATA_COLS:
            ws.Cells(r, c).FormulaR1C1 = f

    # 3. revenue rows: fuel (fuel flags) + process (process flags), both post-policy
    ws.Range("B%d" % (t0 + 146)).Value = "Revenues (fuel CO2 + non-fuel process, post-policy; USD million)"
    f = ("=(R[-62]C+R[-40]C)*(R[-93]C6*%s+R[-93]C7*%s)+(R[-51]C+R[-29]C)*(R[-82]C6*%s+R[-82]C7*%s)" % (p, e, p, e))
    for off in range(147, 155):
        for c in DATA_COLS:
            ws.Cells(t0 + off, c).FormulaR1C1 = f
    ws.Range("E%d" % (t0 + 155)).Value = "USD million"
    ws.Range("I%d" % (t0 + 155)).Value = "revf (o137) + revp (o138); base of o5 Revenues Fund"


def update_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    # reasons on the legacy-comparison rows for the revenue rows
    a = ws.Range("A1:J%d" % last).Value
    n = 0
    for i, row in enumerate(a):
        if row[3] == 2 and row[4] == "diff" and row[8] == "changed" and \
                str(row[0]) in ["Mitigation!%d" % (L_FIRST + o) for o in range(147, 156)]:
            ws.Range("J%d" % (i + 1)).Value = ("Task A (v0.5): fuel + process revenue, process inclusion flags, "
                                               "post-policy emissions")
            n += 1
    if n != 9:
        raise ValueError("expected 9 revenue rows in the legacy CBAM check, found %d" % n)

    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task A (v0.5): CBAM block identities and output-code uniqueness (expected 0)"
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 2
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:J%d" % (hdr, hdr)).Value = ("Item", "", "", "Scenario", "Series", "max |diff|", "", "", "Expected", "")
    for j in range(10, 35):
        ws.Cells(hdr, j + 1).Value = 2011 + j
    r = hdr + 1
    for t0, scen in BLOCKS:
        s = 1 if t0 == BLOCKS[0][0] else 2
        for item, f in (
                ("Total Revenues o155 - (revf o137 + revp o138)", "=Mitigation_Industry!%%s%d-Mitigation_Industry!%%s%d-Mitigation_Industry!%%s%d"
                 % (t0 + 155, t0 + 137, t0 + 138)),
                ("Revenue o155 - sum of product rows o147-o154", "=Mitigation_Industry!%%s%d-SUM(Mitigation_Industry!%%s%d:%%s%d)"
                 % (t0 + 155, t0 + 147, t0 + 154)),
                ("emis o135 - (sector fuel o129-133 + sector process o140-144)",
                 "=Mitigation_Industry!%%s%d-SUM(Mitigation_Industry!%%s%d:%%s%d)-SUM(Mitigation_Industry!%%s%d:%%s%d)"
                 % (t0 + 135, t0 + 129, t0 + 133, t0 + 140, t0 + 144)),
                ("emr o136 - (fuel ER o107-114 + process ER o118-125)",
                 "=Mitigation_Industry!%%s%d-SUM(Mitigation_Industry!%%s%d:%%s%d)-SUM(Mitigation_Industry!%%s%d:%%s%d)"
                 % (t0 + 136, t0 + 107, t0 + 114, t0 + 118, t0 + 125))):
            ws.Range("A%d" % r).Value = item
            ws.Range("D%d:E%d" % (r, r)).Value = (s, "diff")
            for c in DATA_COLS:
                ws.Cells(r, c).Formula = f.replace("%s", col(c))
            ws.Range("F%d" % r).Formula = "=MAX(MAX(L%d:AE%d),-MIN(L%d:AE%d))" % (r, r, r, r)
            ws.Range("I%d" % r).Value = 0
            r += 1
    ws.Range("A%d" % r).Value = "Duplicate output codes (egy.mit.*) in Mitigation_Industry column H"
    ws.Range("F%d" % r).Formula = ('=SUMPRODUCT((LEFT(Mitigation_Industry!$H$1:$H$1000,LEN(Settings!$B$3)+5)='
                                   'Settings!$B$3&".mit.")*(COUNTIF(Mitigation_Industry!$H$1:$H$1000,'
                                   'Mitigation_Industry!$H$1:$H$1000)>1))')
    ws.Range("I%d" % r).Value = 0
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and duplicate-code count (Task A)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (hdr + 1, r)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    return r0 + 1


def update_scenarios(wb):
    ws = wb.Worksheets("Scenarios")
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Stream 1 output codes (Task A, CBAM block column H, suffix = scenario): cpef / cpep.ind.<sec> (post-policy fuel / "
        "process emissions by sector); emisf / emisp / emis.cbam.tot; emrf / emrp / emr.cbam.tot; revf / revp / rev.cbam.tot "
        "(o137 / o138 / o155). o5 fund base = o155 = fuel + process revenue; main CPAT adds only revp (fuel CO2 already in nce).")


def update_settings(wb, chk_row):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype v0.5"
    if ws.Range("A25").Value != "Check":
        raise ValueError("Settings layout changed")
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " v0.5: Task A identity section (row %d)." % chk_row
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = "v0.5"
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task A: CBAM block output codes unique (cpef/cpep, emisf/emisp, emrf/emrp; combined emis/emr in spare rows "
        "o135/o136). Revenues fixed: o147-o154 = post-policy fuel CO2 x fuel inclusion + post-policy process x process "
        "inclusion (legacy: pre-policy process x fuel inclusion, no fuel revenue); split revf/revp in o137/o138, "
        "o155 = sum. Changed by design: scenario-2 revenue rows (LEGACY o155 2030: 1697.5 -> fuel only, process not "
        "covered); all other values identical to v0.4.")
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
        for t0, scen in BLOCKS:
            update_block(ws, t0, scen)
        update_scenarios(wb)
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
