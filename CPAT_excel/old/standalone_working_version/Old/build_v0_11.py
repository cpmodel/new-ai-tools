"""Build CPAT_Industry_Kernel_Egypt_v0.11.xlsx from v0.10 (Task L: CBAM obligations).

For each CBAM product k and year t (direct embedded emissions only; indirect electricity emissions not modelled):

  EI_k     = post-policy embedded intensity (tCO2e/t) = (o85 + o107 + o96 + o118) / production o31
  d_k      = domestic carbon price effectively paid per t = max(0, revenue o147 / production - OBR rebate per t)
             (OBR rebate = min(obrrb, price) x covered pre-policy EF, as in Task H; 3B -> 0, no credit;
              the 3C abatement fund is not deducted, as in EgyptResultsInitial)
  b_k      = obligation per t exported = max(0, CBF_t x P_EU,t x EI_k - S_t x d_k)
  X_k      = EU exports (kt) = 2024 EU exports (o31 I) x production / 2024 production (output response included)
  cbobl    = sum_k X_k x b_k / 1000                     USD million, policy scenario
  cbobl0   = sum_k X0_k x CBF x P_EU x EI0_k / 1000     no domestic policy: pre-policy EF, baseline output
  cbobch   = cbobl / cbobl0 - 1                         total change (intensity + price credit + output)
  cbobchu  = per tonne exported, base-year EU export mix, same output (comparable to EgyptResultsInitial Table 2)

Manual inputs rows 75-82: EU ETS price path (PLACEHOLDER 100 $/t, as in the doc), CBAM factor CBF = 1 - EU free
allocation share (Regulation (EU) 2023/956 Art. 31 / Directive 2003/87/EC Art. 10a(1a): 2.5% 2026 ... 48.5% 2030 ...
100% 2034) and the selector E76: FULL (default; price paid deducted in full), SCALED (deduction x CBF), NOPHASE
(CBF = 1, the doc's convention). Effective CBF in row 81, deduction scaling S in row 82.

Run with Excel installed:  python build_v0_11.py
"""
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLOCKS, DATA_COLS, REVIEW, TAN, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.10.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.10.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.11.xlsx")

YEARS = list(range(2022, 2042))                       # columns L..AE
CBF = {2026: 0.025, 2027: 0.05, 2028: 0.10, 2029: 0.225, 2030: 0.485, 2031: 0.61, 2032: 0.735, 2033: 0.86}
CBF_OF = lambda y: 0.0 if y < 2026 else CBF.get(y, 1.0)
P_EU = 100.0

R_TITLE, R_SEL, R_NOTE, R_HDR, R_PEU, R_CBF, R_CBFU, R_S = 75, 76, 77, 78, 79, 80, 81, 82   # Manual inputs
MI = "'Manual inputs'!"


def mi_section(wb):
    ws = wb.Worksheets("Manual inputs")
    for r in range(R_TITLE, R_S + 3):
        if any(ws.Cells(r, c).Formula not in ("", None) for c in range(1, 32)):
            raise ValueError("Manual inputs row %d not empty" % r)
    copy_formats(ws.Rows(62), ws.Rows(R_TITLE))
    copy_formats(ws.Rows(65), ws.Rows(R_HDR))
    for r in (R_PEU, R_CBF, R_CBFU, R_S):
        copy_formats(ws.Rows(66), ws.Rows(r))
    wb.Application.CutCopyMode = False
    ws.Range("B%d" % R_TITLE).Value = ("CBAM obligations (Task L, v0.11): EU ETS price and CBAM phase-in; used by CBAM "
                                       "block rows o94, o105, o116, o127")
    ws.Range("D%d" % R_SEL).Value = "CBAM obligation convention (FULL / SCALED / NOPHASE)"
    ws.Range("E%d" % R_SEL).Value = "FULL"
    ws.Range("E%d" % R_SEL).Interior.Color = ws.Range("E21").Interior.Color
    v = ws.Range("E%d" % R_SEL).Validation
    v.Delete()
    v.Add(3, 1, 1, "FULL,SCALED,NOPHASE")
    ws.Range("F%d" % R_SEL).Value = ("FULL: CBF phase-in, domestic price paid deducted in full; SCALED: deduction x CBF; "
                                     "NOPHASE: CBF = 1 (EgyptResultsInitial / ad hoc convention)")
    ws.Range("D%d" % R_NOTE).Value = (
        "Obligation per t exported = max(0, CBF x P_EU x embedded intensity - S x price effectively paid per t, net "
        "of OBR rebates). Direct emissions only (fc + fp + np + no); EU exports scale with production.")
    ws.Range("D%d" % R_HDR).Value = "Series"
    ws.Range("E%d" % R_HDR).Value = "Unit"
    ws.Range("F%d" % R_HDR).Value = "Source"
    rows = ((R_PEU, "EU ETS / CBAM certificate price", "$/tCO2e real",
             "PLACEHOLDER 100 flat (EgyptResultsInitial / AdHocCalculations assumption); replace with an EU ETS outlook"),
            (R_CBF, "CBAM factor = 1 - EU free allocation share (input)", "share",
             "Directive 2003/87/EC Art. 10a(1a) as amended 2023 (CBAM Reg. 2023/956 Art. 31): 2.5% 2026 ... 100% 2034"),
            (R_CBFU, "CBAM factor used", "share", "1 if NOPHASE, else the input row"),
            (R_S, "Deduction scaling used (S)", "share", "CBF used if SCALED, else 1"))
    for r, lab, unit, src in rows:
        ws.Range("D%d" % r).Value = lab
        ws.Range("E%d" % r).Value = unit
        ws.Range("F%d" % r).Value = src
        ws.Range("D%d:K%d" % (r, r)).WrapText = False
        ws.Range("E%d:K%d" % (r, r)).Interior.ColorIndex = -4142
    for j, y in enumerate(YEARS):
        c = 12 + j
        ws.Cells(R_HDR, c).Value = y
        ws.Cells(R_PEU, c).Value = P_EU
        ws.Cells(R_CBF, c).Value = CBF_OF(y)
        a = col(c)
        ws.Cells(R_CBFU, c).Formula = '=IF($E${0}="NOPHASE",1,{1}{2})'.format(R_SEL, a, R_CBF)
        ws.Cells(R_S, c).Formula = '=IF($E${0}="SCALED",{1}{2},1)'.format(R_SEL, a, R_CBFU)
    ws.Range("L%d:AE%d" % (R_PEU, R_PEU)).Interior.Color = REVIEW
    ws.Range("L%d:AE%d" % (R_CBF, R_CBF)).Interior.Color = ws.Range("E21").Interior.Color
    ws.Range("L%d:AE%d" % (R_CBFU, R_S)).Interior.Color = TAN
    ws.Range("L%d:AE%d" % (R_PEU, R_PEU)).NumberFormat = "0"
    ws.Range("L%d:AE%d" % (R_CBF, R_S)).NumberFormat = "0.0%"


def update_block(ws, t0):
    o = lambda n: t0 + n
    if ws.Cells(o(29), 20).Value != 2030 or ws.Cells(o(147), 4).Value != "DRI-EAF steel":
        raise ValueError("block %d layout changed" % t0)
    for n in (94, 105, 116, 127):
        if any(ws.Cells(o(n), c).Formula not in ("", None) for c in range(1, 36)):
            raise ValueError("spare row o%d not empty (block %d)" % (n, t0))
    rr = lambda a, c="": "R%dC%s:R%dC%s" % (o(a), c, o(a + 7), c)
    Q, N, I, G = rr(31), rr(31, 14), rr(31, 9), rr(31, 7)
    sQ = "(%s+(%s=0))" % (Q, Q)
    sN = "(%s+(%s=0))" % (N, N)
    EI = "(%s+%s+%s+%s)*1000/%s" % (rr(85), rr(107), rr(96), rr(118), sQ)
    EI0 = "(%s+%s)" % (rr(85, 6), rr(85, 7))
    reb = ("(MIN(R{rb}C,R{cp}C)*{f6}*{f5}+MIN(R{rb}C,R{pp}C)*{p6}*{p5})"
           .format(rb=o(6), cp=o(2), pp=o(63), f6=rr(54, 6), f5=rr(54, 5), p6=rr(65, 6), p5=rr(65, 5)))
    dn = "(%s*1000/%s-%s)" % (rr(147), sQ, reb)
    d = "((%s+ABS(%s))/2)" % (dn, dn)
    cbf, peu, s = ("%sR%dC" % (MI, r) for r in (R_CBFU, R_PEU, R_S))
    x = "(%s*%s*%s-%s*%s)" % (cbf, peu, EI, s, d)
    b = "((%s+ABS(%s))/2)" % (x, x)
    qb = "%s*(1+%s)^(R%dC-R%dC14)" % (N, G, o(29), o(29))
    base_u = "SUMPRODUCT(%s*%s*%s*%s)" % (I, cbf, peu, EI0)
    rows = (
        (94, 155, "CBAM obligations on EU exports, policy scenario (direct emissions; Task L)", "USD million", "cbobl",
         "=SUMPRODUCT(%s*%s/%s*%s)/1000" % (I, Q, sN, b)),
        (105, 155, "CBAM obligations without domestic carbon pricing (pre-policy intensity, baseline output)",
         "USD million", "cbobl0", "=SUMPRODUCT(%s*%s/%s*%s*%s*%s)/1000" % (I, qb, sN, cbf, peu, EI0)),
        (116, 41, "CBAM obligations change vs no domestic pricing (intensity + price credit + output)", "share",
         "cbobch", "=IF(R%dC=0,0,R%dC/R%dC-1)" % (o(105), o(94), o(105))),
        (127, 41, "CBAM obligations change per t exported, base-year EU export mix, same output (Table 2 metric)",
         "share", "cbobchu", "=IF(%s=0,0,SUMPRODUCT(%s*%s)/%s-1)" % (base_u, I, b, base_u)),
    )
    scen_ref = ws.Range("H%d" % o(126)).Formula.rsplit("&", 1)[1]
    if not scen_ref.startswith("$J$"):
        raise ValueError("unexpected emrp code formula in block %d" % t0)
    for n, fmt_from, label, unit, code, f in rows:
        r = o(n)
        copy_formats(ws.Rows(o(fmt_from)), ws.Rows(r))
        ws.Range("D%d" % r).Value = label
        ws.Range("E%d" % r).Value = unit
        ws.Range("G%d" % r).Value = "cbam.tot"
        ws.Range("H%d" % r).Formula = '=Settings!$B$3&".mit.%s."&$G%d&"."&%s' % (code, r, scen_ref)
        for c in DATA_COLS:
            ws.Cells(r, c).FormulaR1C1 = f
    ws.Application.CutCopyMode = False


def update_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task L (v0.11): CBAM obligations (expected 0)"
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 2
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:I%d" % (hdr, hdr)).Value = ("Item", "", "", "Scenario", "Series", "max |diff| / count", "", "",
                                              "Expected")
    M = "Mitigation_Industry!"
    sel = "%s$E$%d" % (MI, R_SEL)
    yr = lambda r: "%s$L$%d:$AE$%d" % (MI, r, r)
    items = [
        ("Selector %sE%d not FULL, SCALED or NOPHASE (count)" % (MI, R_SEL),
         '=1-OR(%s="FULL",%s="SCALED",%s="NOPHASE")' % (sel, sel, sel)),
        ("CBAM factor input outside [0,1] or decreasing (count)",
         "=SUMPRODUCT((%s<0)+(%s>1))+SUMPRODUCT(--(%s$M$%d:$AE$%d<%s$L$%d:$AD$%d))"
         % (yr(R_CBF), yr(R_CBF), MI, R_CBF, R_CBF, MI, R_CBF, R_CBF)),
        ("CBAM factor used / deduction scaling differ from the selector rule (max abs)",
         '=MAX(SUMPRODUCT(ABS(%s-(%s="NOPHASE")-(%s<>"NOPHASE")*%s)),SUMPRODUCT(ABS(%s-(%s<>"SCALED")-(%s="SCALED")*%s)))'
         % (yr(R_CBFU), sel, sel, yr(R_CBF), yr(R_S), sel, sel, yr(R_CBFU))),
        ("EU price negative (count)", "=SUMPRODUCT(--(%s<0))" % yr(R_PEU)),
    ]
    r = hdr + 1
    for item, f in items:
        ws.Range("A%d" % r).Value = item
        ws.Range("F%d" % r).Formula = f
        ws.Range("I%d" % r).Value = 0
        r += 1
    for t0, scen in BLOCKS:
        s = 1 if t0 == BLOCKS[0][0] else 2
        row = lambda n: "%s$L$%d:$AE$%d" % (M, t0 + n, t0 + n)
        add = [("Obligations negative (count)", "=SUMPRODUCT((%s<0)+(%s<0))" % (row(94), row(105))),
               ("cbobch - (cbobl / cbobl0 - 1) (max abs)",
                "=SUMPRODUCT(ABS(%s-(%s<>0)*(%s/(%s+(%s=0))-1)))" % (row(116), row(105), row(94), row(105), row(105)))]
        if s == 1:
            add.append(("Scenario 1: cbobl - cbobl0 (no domestic policy; sum abs)",
                        "=SUMPRODUCT(ABS(%s-%s))" % (row(94), row(105))))
        for item, f in add:
            ws.Range("A%d" % r).Value = item
            ws.Range("D%d" % r).Value = s
            ws.Range("F%d" % r).Formula = f
            ws.Range("I%d" % r).Value = 0
            r += 1
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (Task L)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (hdr + 1, r - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    return r0 + 1


def update_scenarios(wb):
    ws = wb.Worksheets("Scenarios")
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task L (stream 1, v0.11): CBAM obligation codes cbobl (o94, USD m), cbobl0 (o105, no domestic pricing), "
        "cbobch (o116, total change), cbobchu (o127, per t exported, same output). Domestic price credit = block "
        "revenue per t net of obrrb (o6); the abatement fund (o4/o5) is not deducted. EU price and CBAM phase-in in "
        "Manual inputs rows %d-%d (selector E%d)." % (R_TITLE, R_S, R_SEL))


def update_settings(wb, chk_row):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype v0.11"
    if ws.Range("A25").Value != "Check":
        raise ValueError("Settings layout changed")
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " v0.11: Task L section (row %d)." % chk_row
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = "v0.11"
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task L: CBAM obligations on EU exports (cbobl, cbobl0, cbobch, cbobchu in block o94/o105/o116/o127) with "
        "EU price (placeholder 100), CBAM phase-in factor and deduction convention (Manual inputs rows %d-%d, "
        "selector E%d = FULL)." % (R_TITLE, R_S, R_SEL))
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
        mi_section(wb)
        ws = wb.Worksheets("Mitigation_Industry")
        for t0, scen in BLOCKS:
            update_block(ws, t0)
        chk = update_check(wb)
        update_scenarios(wb)
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
