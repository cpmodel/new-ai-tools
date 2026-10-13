"""Build CPAT_Industry_Kernel_Egypt_v1.5.xlsx from v1.4: the single workbook that confirms the final Table 2.

DRAFTED ON LINUX, NOT RUN. Needs Excel + pywin32 on Windows and kernel v1.4 (build_v1_4.py, task T3) in this folder.
Run from this folder in a normal (non-sandboxed) shell:   python build_v1_5.py

Changes (all deliberate; regression elsewhere must be 0):
  1. 3B rebate decision (2026-10-04): the 3B output-based rebate covers ALL covered industry.
     * 'Manual inputs' rows 111-113: new switch E112 (1 = all covered industry, 0 = CBAM block only).
     * CarveOut_Table2 section G (rows 92-98): D_nb, D_nb_f, rebate to non-block industry; K (D67), P (D73) and Q (D74)
       now include them. Non-3B scenarios are unchanged (D93 = 0). Arithmetic = make_carveout_v0_4.py.
  2. New sheet Table2_Final (after CarveOut_Table2): final Table 2 for the six scenarios (from the stored snapshot),
     the figures as printed in the final documents (typed), the differences (count must be 0), the deduction-based
     obligation memo (2030 phase-in = FULL; no phase-in = NOPHASE) and a LIVE recomputation of Table 2 row O (CBAM-product
     intensity change) for the active scenario.
  3. The stored snapshot row for 3B (CarveOut_Table2 row 84) is refreshed from the live sheet; the other five rows and
     the stored intensity-only O (column M) are verified against the live sheet (|diff| <= 0.02) but not rewritten.
  4. Settings title and version-log row; Check section.
Gates (nothing is saved unless all pass): 3B live = Python mirror (carveout_v1_5_results.json, tol 0.02; deaths 1.0);
live = stored for the other scenarios; live intensity-only O = stored O; every pre-existing cell outside the intended
areas unchanged (max abs diff 0); printed-vs-workbook mismatch count 0.
After a clean run it copies the kernel to egypt/final/ and moves the previous final kernel to egypt/archive/.
Bookkeeping afterwards: NORMS section 6 (CAVEATS entry, yaml, EgyptTaskReference, context-egypt, TODO).
"""
import datetime
import json
import os
import shutil
import sys

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "Old"))
from build_v0_4 import REVIEW, SECTION, TAN, TITLE, copy_formats  # noqa: E402

REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
ADHOC = os.path.join(REPO, "egypt", "supporting", "AdHocRebuild")
FINAL = os.path.join(REPO, "egypt", "final")
ARCHIVE = os.path.join(REPO, "egypt", "archive")
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v1.4.xlsx")
SRC_OLD = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v1.4.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v1.5.xlsx")
VER = "v1.5"
BUNDLES = ["1A", "2A", "2B", "3A", "3B", "3C"]
CO = "CarveOut_Table2"
T2F = "Table2_Final"
MI = "Manual inputs"
MIT = "Mitigation_Industry"

# metric rows of Table2_Final section A / B / C: (label, unit, stored column on CarveOut_Table2, json key, dp)
METRICS = [
    ("J  Coverage of national GHG", "%", "C", "J", 0),
    ("P  Carbon revenue", "USD bn", "J", "P", 1),
    ("K  Emission cut vs baseline", "MtCO2e", "D", "K", 1),
    ("L  Cut, % of national GHG", "%", "E", "L", 1),
    ("M  CBAM coverage", "%", "F", "M", 0),
    ("N  CBAM-sector emission intensity change", "%", "G", "N", 1),
    ("O  CBAM obligations (CBAM-product intensity change, no deduction)", "%", "M", "O_final", 1),
    ("T  CBAM block emissions change", "%", "L", "T", 1),
    ("Q  Premature deaths avoided", "deaths", "K", "Q", 0),
]
M_CBAM = {"1A": 100, "2A": 44, "2B": 44, "3A": 100, "3B": 100, "3C": 100}


def rnd(x, dp):
    """Excel ROUND: half away from zero."""
    q = 10 ** dp
    return (int(abs(x) * q + 0.5 + 1e-9) / q) * (1 if x >= 0 else -1)


def printed_values():
    res = json.load(open(os.path.join(ADHOC, "carveout_v1_5_results.json"), encoding="utf8"))
    out = {}
    for b in BUNDLES:
        r = dict(res[b])
        r["M"] = M_CBAM[b]
        out[b] = {key: rnd(r[key], dp) for (_, _, _, key, dp) in METRICS}
    return res, out


def snapshot(wb):
    snap = {}
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        v = ur.Value
        snap[ws.Name] = (ur.Row, ur.Column, v if isinstance(v, tuple) else ((v,),))
    return snap


def max_diff(before, after, skip):
    worst, bad = 0.0, []
    for name, (r0, c0, vb) in before.items():
        ra, ca, va = after[name]
        for i, rowb in enumerate(vb):
            for j, b in enumerate(rowb):
                r, c = r0 + i, c0 + j
                if skip(name, r, c):
                    continue
                ia, ja = r - ra, c - ca
                a = va[ia][ja] if 0 <= ia < len(va) and 0 <= ja < len(va[0]) else None
                if isinstance(b, (int, float)) and isinstance(a, (int, float)):
                    d = abs(a - b)
                    worst = max(worst, d)
                    if d > 0 and len(bad) < 10:
                        bad.append((name, r, c, b, a))
                elif a != b:
                    worst = max(worst, 1.0)
                    if len(bad) < 10:
                        bad.append((name, r, c, b, a))
    return worst, bad


def set_formula(ws, addr, expected, new):
    cur = ws.Range(addr).Formula
    assert cur.replace(" ", "") == expected.replace(" ", ""), "%s!%s is %r, expected %r" % (ws.Name, addr, cur, expected)
    ws.Range(addr).Formula = new


# ---------------------------------------------------------------- Manual inputs: the switch
def add_switch(wb):
    mi = wb.Worksheets(MI)
    for r in range(110, 115):
        assert all(c is None for c in mi.Range("A%d:AE%d" % (r, r)).Value[0]), "Manual inputs row %d not empty" % r
    mi.Range("B111").Value = "Final Table 2 conventions (v1.5)"
    mi.Range("B111").Font.Bold = True
    mi.Range("D112").Value = ("3B output-based rebate also covers non-CBAM covered industry (1 = all covered industry; "
                              "0 = CBAM block only)")
    mi.Range("E112").Value = 1
    mi.Range("E112").Interior.Color = TAN
    mi.Range("F112").Value = ("Decision 2026-10-04 (CAVEATS): rebate to all covered industry; same switch as ThetaOther in the "
                              "AdHoc rebuild. Used by CarveOut_Table2 section G (3B only).")
    mi.Range("G112").Value = "Medium"
    mi.Range("H112").Value = "egy.in.thetao.---"
    mi.Range("D113").Value = ("CBAM obligation conventions (rows 76-82) stay a memo: FULL = 2030 phase-in (CBAM factor 0.485), "
                              "NOPHASE = no phase-in (factor 1). Table 2 row O is the CBAM-product intensity change and does "
                              "not depend on them.")


# ---------------------------------------------------------------- CarveOut_Table2: 3B rule
def carveout_changes(wb):
    ws = wb.Worksheets(CO)
    assert ws.Range("D6").HasFormula and ws.Range("B90").Value.startswith("Not re-solved"), "unexpected CarveOut layout"
    for r in range(91, 100):
        assert all(c is None for c in ws.Range("A%d:N%d" % (r, r)).Value[0]), "CarveOut row %d not empty" % r
    rows = [
        (92, "G. 3B rebate of all covered industry (decision 2026-10-04; switch 'Manual inputs'!E112); 0 for other scenarios", None, None),
        (93, "Active (3B and switch = 1)", "flag", "=IF(AND(D11=1,'Manual inputs'!$E$112=1),1,0)"),
        (94, "dIPPU of the run (CPAT)", "MtCO2e", "=Table2_Industry!$T$50"),
        (95, "NB: CPAT response of non-block industry (fuel + proportional IPPU) = dInd + dIPPU - r x (Bf + Bp)", "MtCO2e",
         "=(D15-D14)+D94-D30*(D44+D45)"),
        (96, "D_nb: output share removed = -(1 - s_int) x NB (added to K)", "MtCO2e", "=-(1-D20)*D95*D93"),
        (97, "D_nb_f: fuel part of D_nb = -(1 - s_int) x (dInd - r x Bf) (added to the energy-CO2 adjustment for deaths)",
         "MtCO2", "=-(1-D20)*((D15-D14)-D30*D44)*D93"),
        (98, "Rebate to non-block covered industry = tau x (kappa x IND1 - block post-policy fuel CO2) / 1000 (deducted from P)",
         "$bn", "=D93*D17*(D16*D15-(D44+D46))/1000"),
    ]
    for r, lab, unit, f in rows:
        ws.Range("B%d" % r).Value = lab
        if unit:
            ws.Range("C%d" % r).Value = unit
            ws.Range("D%d" % r).Formula = f
            ws.Range("D%d" % r).Interior.Color = TAN
    ws.Range("B92").Font.Bold = True
    set_formula(ws, "D67", "=D54+D49+D50", "=D54+D49+D50+D96")
    ws.Range("E67").Value = "CPAT dGHG + adj_f + adj_p (+ D_nb, 3B)"
    set_formula(ws, "D73", "=D55+D17*D49/1000+D58-D59-D60", "=D55+D17*D49/1000+D58-D59-D60-D98")
    ws.Range("E73").Value = "CPAT receipts + tau x adj_f + process fees - rebate - fund (- rebate to non-block industry, 3B)"
    set_formula(ws, "D74", "=D56*(D57+D49)/D57", "=D56*(D57+D49+D97)/D57")
    ws.Range("E74").Value = "CPAT deaths x (dEnergy + adj_f (+ D_nb_f, 3B))/dEnergy"
    t = str(ws.Range("B1").Value)
    ws.Range("B1").Value = t.replace("final v1.3", "final " + VER)
    ws.Range("B90").Value = ("Not re-solved: the 3C fund and the 3B rebate to the block are kernel values (block fuel-payment "
                             "change < $0.05bn); the 3B rebate to non-block industry is section G")
    ws.Range("B78").Value = ("F. Stored 2030 snapshot, all bundles (make_carveout_v0_4.py; row 84 (3B) refreshed in %s from the live "
                             "sheet; = EGYPT_CarveOut_Table2_%s)" % (VER, VER))


# ---------------------------------------------------------------- Table2_Final
def build_final_sheet(wb, printed):
    wb.Worksheets.Add(None, wb.Worksheets(CO))
    ws = wb.ActiveSheet
    ws.Name = T2F
    ws.Range("B1").Value = "Table 2 (2030), final %s: confirmation of the final numbers in one workbook" % VER
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 12
    ws.Range("B2").Value = ("Section A is the final Table 2 (stored 2030 snapshot of CarveOut_Table2, which is computed line by line "
                            "from CPAT runs and this kernel's CBAM block). B is the figure printed in the final documents. C is the "
                            "difference at the printed precision (must be 0). D is the deduction-based obligation memo. "
                            "E recomputes row O live for the scenario active in Settings!B10.")
    scen_cols = "DEFGHI"

    def header(r, title):
        ws.Range("B%d" % r).Value = title
        ws.Range("B%d" % r).Font.Bold = True
        ws.Range("B%d:J%d" % (r, r)).Interior.Color = SECTION
        for j, b in enumerate(BUNDLES):
            ws.Range("%s%d" % (scen_cols[j], r + 1)).Value = b
            ws.Range("%s%d" % (scen_cols[j], r + 1)).Font.Bold = True
        ws.Range("B%d" % (r + 1)).Value = "Metric"
        ws.Range("C%d" % (r + 1)).Value = "Unit"
        ws.Range("B%d:C%d" % (r + 1, r + 1)).Font.Bold = True

    # A. final values (formulas on the stored snapshot)
    header(4, "A. Final Table 2, 2030 (stored snapshot of CarveOut_Table2, rows 80-85)")
    ra = 6
    for i, (lab, unit, col, key, dp) in enumerate(METRICS):
        r = ra + i
        ws.Range("B%d" % r).Value = lab
        ws.Range("C%d" % r).Value = unit
        for j in range(6):
            c = scen_cols[j]
            ws.Range("%s%d" % (c, r)).Formula = (
                "=INDEX(%s!$%s$80:$%s$85,MATCH(%s$5,%s!$B$80:$B$85,0))" % (CO, col, col, c, CO))
            ws.Range("%s%d" % (c, r)).NumberFormat = "0.0" if dp else "#,##0"
            ws.Range("%s%d" % (c, r)).Interior.Color = TAN
    ws.Range("J6").Value = "Source"
    ws.Range("J7").Value = "CarveOut_Table2 stored snapshot (live-vs-stored checked below and on Check)"

    # B. printed values (typed)
    header(17, "B. As printed in the final documents (typed): EGYPT_CarveOut_Table2_%s, results text %s Table 2" % (VER, VER))
    rb = 19
    for i, (lab, unit, col, key, dp) in enumerate(METRICS):
        r = rb + i
        ws.Range("B%d" % r).Value = lab
        ws.Range("C%d" % r).Value = unit
        for j, b in enumerate(BUNDLES):
            c = ws.Range("%s%d" % (scen_cols[j], r))
            c.Value = printed[b][key]
            c.NumberFormat = "0.0" if dp else "#,##0"
            c.Interior.Color = REVIEW

    # C. differences
    header(30, "C. Difference at printed precision: ROUND(A, dp) - B (must be 0)")
    rc = 32
    for i, (lab, unit, col, key, dp) in enumerate(METRICS):
        r = rc + i
        ws.Range("B%d" % r).Value = lab
        ws.Range("C%d" % r).Value = "dp = %d" % dp
        for j in range(6):
            c = scen_cols[j]
            ws.Range("%s%d" % (c, r)).Formula = "=ROUND(%s%d,%d)-%s%d" % (c, ra + i, dp, c, rb + i)
            ws.Range("%s%d" % (c, r)).NumberFormat = "0.0"
    ws.Range("B42").Value = "Mismatches (count of non-zero differences; expected 0)"
    ws.Range("B42").Font.Bold = True
    ws.Range("D42").Formula = "=SUMPRODUCT(--(ABS(D32:I40)>0.00001))"
    ws.Range("D42").Interior.Color = REVIEW

    # D. memo
    header(45, "D. Memo, not Table 2: obligation net of the Egyptian carbon price (deduction-based, Task L), stored snapshot")
    memo = [("Obligation, 2030 phase-in (FULL: CBAM factor 0.485)", "H"), ("Obligation, no phase-in (NOPHASE: CBAM factor 1)", "I")]
    for i, (lab, col) in enumerate(memo):
        r = 47 + i
        ws.Range("B%d" % r).Value = lab
        ws.Range("C%d" % r).Value = "%"
        for j in range(6):
            c = scen_cols[j]
            ws.Range("%s%d" % (c, r)).Formula = "=INDEX(%s!$%s$80:$%s$85,MATCH(%s$46,%s!$B$80:$B$85,0))" % (CO, col, col, c, CO)
            ws.Range("%s%d" % (c, r)).NumberFormat = "0.0"
            ws.Range("%s%d" % (c, r)).Interior.Color = TAN
    ws.Range("B49").Value = ("Row O of Table 2 is the CBAM-product intensity change (section A); the EU price and phase-in cancel "
                             "out, so it does not depend on these conventions.")

    # E. live row O for the active scenario
    ws.Range("B52").Value = "E. Live recomputation of row O for the active scenario (Settings!B10 = %s)" % "bundle"
    ws.Range("B52").Font.Bold = True
    ws.Range("B52:L52").Interior.Color = SECTION
    heads = ["Product", "EU exports 2024 (kt)", "Output 2030 (kt)", "Baseline fuel EI (t/t)", "Baseline process EI (t/t)",
             "Extra fuel abatement (t/t)", "Policy process EI (t/t)", "Baseline EI", "Policy EI", "X x baseline EI",
             "X x policy EI"]
    for j, h in enumerate(heads):
        c = ws.Cells(53, 2 + j)
        c.Value = h
        c.Font.Bold = True
        c.WrapText = True
    for k in range(8):
        r = 54 + k
        ws.Range("B%d" % r).Formula = "=%s!D%d" % (MIT, 675 + k)
        ws.Range("C%d" % r).Formula = "=%s!I%d" % (MIT, 675 + k)
        ws.Range("D%d" % r).Formula = "=%s!T%d" % (MIT, 675 + k)
        ws.Range("E%d" % r).Formula = "=%s!F%d" % (MIT, 729 + k)
        ws.Range("F%d" % r).Formula = "=%s!G%d" % (MIT, 729 + k)
        ws.Range("G%d" % r).Formula = "=IF(D%d>0,%s!T%d*1000/D%d,0)" % (r, MIT, 751 + k, r)
        ws.Range("H%d" % r).Formula = "=IF(D%d>0,(%s!T%d+%s!T%d)*1000/D%d,0)" % (r, MIT, 740 + k, MIT, 762 + k, r)
        ws.Range("I%d" % r).Formula = "=E%d+F%d" % (r, r)
        ws.Range("J%d" % r).Formula = "=E%d*(1+%s!$D$31)+G%d+H%d" % (r, CO, r, r)
        ws.Range("K%d" % r).Formula = "=C%d*I%d" % (r, r)
        ws.Range("L%d" % r).Formula = "=C%d*J%d" % (r, r)
        ws.Range("E%d:L%d" % (r, r)).NumberFormat = "0.0000"
        ws.Range("B%d:L%d" % (r, r)).Interior.Color = TAN
    ws.Range("B63").Value = "O live = 100 x (sum X x policy EI / sum X x baseline EI - 1)"
    ws.Range("D63").Formula = '=IFERROR(100*(SUM(L54:L61)/SUM(K54:K61)-1),"n/a")'
    ws.Range("B64").Value = "O stored (active scenario)"
    ws.Range("D64").Formula = '=IFERROR(INDEX(%s!$M$80:$M$85,MATCH(%s!$D$6,%s!$B$80:$B$85,0)),"n/a")' % (CO, CO, CO)
    ws.Range("B65").Value = "Live - stored (stored value is rounded to 2 decimals; |diff| < 0.02 expected)"
    ws.Range("D65").Formula = '=IFERROR(D63-D64,"n/a")'
    for a in ("D63", "D64", "D65"):
        ws.Range(a).NumberFormat = "0.000"
        ws.Range(a).Interior.Color = REVIEW
    ws.Columns("B").ColumnWidth = 62
    ws.Columns("C").ColumnWidth = 14
    ws.Range("D:L").ColumnWidth = 13
    ws.Range("B53:L53").RowHeight = 45
    return ws


def add_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "%s: final Table 2 confirmation (sheet %s) and 3B rebate decision" % (VER, T2F)
    ws.Range("A%d" % r0).Font.Bold = True
    items = [
        ("Table2_Final: figures printed in the final documents differ from the workbook (count; expected 0)", "=%s!D42" % T2F),
        ("Table2_Final: live row O - stored row O, active scenario (n/a for LEGACY; |diff| < 0.02 expected)", "=%s!D65" % T2F),
        ("CarveOut_Table2: max |live - stored|, active scenario (n/a for LEGACY; < 0.01 expected)", "=%s!D88" % CO),
        ("3B rebate switch 'Manual inputs'!E112 (1 = all covered industry)", "='%s'!E112" % MI),
    ]
    for i, (lab, f) in enumerate(items):
        ws.Range("A%d" % (r0 + 1 + i)).Value = lab
        ws.Range("D%d" % (r0 + 1 + i)).Formula = f
        ws.Range("D%d" % (r0 + 1 + i)).Interior.Color = REVIEW
    return r0 + 1


def settings(wb, diff, chk):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt, final " + VER
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "3B rebate to all covered industry (Manual inputs E112; CarveOut_Table2 section G rows 92-98; K, P, Q for 3B only). "
        "New sheet %s: final Table 2, figures as printed in the documents, differences (count 0), obligation memo, live row O. "
        "Max abs diff over all other pre-existing cells = %.3g. Check row %d. v1.1-v1.3 were relabels of the Egypt deliverable "
        "set (no kernel change)." % (T2F, diff, chk))
    ws.Range("C%d" % (last + 1)).WrapText = False


def main():
    assert os.path.exists(SRC), "kernel v1.4 (build_v1_4.py) is missing: " + SRC
    res, printed = printed_values()
    if os.path.exists(DST):
        os.remove(DST)
    if not os.path.exists(SRC_OLD):
        shutil.copyfile(SRC, SRC_OLD)
    shutil.copyfile(SRC, DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    ok = False
    try:
        wb = xl.Workbooks.Open(DST)
        xl.CalculateFull()
        before = snapshot(wb)
        chk_last_before = before["Check"][0] + len(before["Check"][2]) - 1
        co_before = {name: before[name] for name in (CO,)}

        add_switch(wb)
        carveout_changes(wb)
        build_final_sheet(wb, printed)
        chk_row = add_check(wb)
        xl.CalculateFull()
        wsS = wb.Worksheets("Settings")
        active0 = wsS.Range("B10").Value

        # ---- loop the six scenarios: verify live vs mirror / stored, refresh the 3B stored row
        co, t2f = wb.Worksheets(CO), wb.Worksheets(T2F)
        problems, live3b = [], None
        for i, b in enumerate(BUNDLES):
            wsS.Range("B10").Value = b
            xl.CalculateFull()
            assert co.Range("D6").Value == b, "active scenario is %r, expected %s" % (co.Range("D6").Value, b)
            live = [co.Range("%s86" % c).Value for c in "CDEFGHIJKL"]          # J K L M N O Onp P Q T
            o_live = t2f.Range("D63").Value
            row = 80 + i
            stored = [co.Range("%s%d" % (c, row)).Value for c in "CDEFGHIJKL"]
            exp = res[b]
            mirror = {"J": exp["J"], "K": exp["K"], "P": exp["P"], "Q": exp["Q"], "N": exp["N"], "T": exp["T"]}
            got = dict(zip("J K L M N O Onp P Q T".split(), live))
            for k, tol in (("J", 0.02), ("K", 0.02), ("P", 0.02), ("Q", 1.0), ("N", 0.02), ("T", 0.02)):
                if abs(got[k] - mirror[k]) > tol:
                    problems.append("%s %s: live %.4f vs Python mirror %.4f" % (b, k, got[k], mirror[k]))
            if abs(o_live - exp["O_final"]) > 0.02:
                problems.append("%s O: live %.4f vs expected %.4f" % (b, o_live, exp["O_final"]))
            if abs(o_live - co.Range("M%d" % row).Value) > 0.02:
                problems.append("%s O: live %.4f vs stored %.4f" % (b, o_live, co.Range("M%d" % row).Value))
            if b != "3B":
                for k, lv, sv in zip("J K L M N O Onp P Q T".split(), live, stored):
                    if abs(lv - sv) > 0.01 and not (k == "Q" and abs(lv - sv) < 1.0):
                        problems.append("%s %s: live %.4f vs stored %.4f" % (b, k, lv, sv))
            else:
                live3b = live
            print(b, ["%.3f" % x for x in live], "O live %.3f" % o_live)
        wsS.Range("B10").Value = active0
        xl.CalculateFull()
        if problems:
            print("NOT saved. Problems:")
            for p in problems:
                print("  ", p)
            wb.Close(False)
            sys.exit(1)
        for c, v in zip("CDEFGHIJKL", live3b):
            co.Range("%s84" % c).Value = v
        xl.CalculateFull()

        after = snapshot(wb)

        def skip(name, r, c):
            if name == "Settings":
                return True
            if name == MI and r >= 110:
                return True
            if name == "Check" and r > chk_last_before:
                return True
            if name == CO and (r >= 78 or r == 1 or (r in (67, 73, 74) and c == 5)):
                return True                                  # title, notes, stored section F and new section G
            return False

        diff, bad = max_diff(before, after, skip)
        print("max abs diff over pre-existing cells (intended areas excluded): %.6g" % diff)
        for x in bad:
            print("  DIFF", x)
        mism = t2f.Range("D42").Value
        print("printed-vs-workbook mismatches (expected 0):", mism)
        if diff != 0 or mism != 0:
            print("NOT saved: regression or document check failed.")
            wb.Close(False)
            sys.exit(1)
        settings(wb, diff, chk_row)
        wb.Save()
        wb.Close(False)
        ok = True
    finally:
        xl.Quit()
    if ok:
        os.replace(SRC, SRC_OLD)
        os.makedirs(FINAL, exist_ok=True)
        old_final = os.path.join(FINAL, "CPAT_Industry_Kernel_Egypt_v1.3.xlsx")
        if os.path.exists(old_final):
            shutil.move(old_final, os.path.join(ARCHIVE, "CPAT_Industry_Kernel_Egypt_v1.3.xlsx"))
        shutil.copyfile(DST, os.path.join(FINAL, os.path.basename(DST)))
        print("saved", DST, "; copied to egypt/final; v1.4 moved to Old/; v1.3 final moved to egypt/archive")


if __name__ == "__main__":
    main()
