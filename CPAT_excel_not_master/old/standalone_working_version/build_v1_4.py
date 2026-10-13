"""Build CPAT_Industry_Kernel_Egypt_v1.4.xlsx from v1.3: T3 - CBAM product market data moved from Mitigation_Industry
into 'Manual inputs' (new section at rows 98-108) and linked from both scenario blocks.

DRAFTED ON LINUX, NOT RUN. Needs Excel + pywin32 on Windows. Run from this folder in a normal (non-sandboxed) shell:
    python build_v1_4.py

What it does (values are read from the workbook itself, so the move cannot change a number):
  * New 'Manual inputs' section "CBAM product market data (Task T3)": title row 98, note row 99, header row 100,
    8 product rows 101-108 (same product order as rows 53-60 / 66-73): E 2024 production kt, F growth %/yr,
    G 2024 total exports kt (blank placeholder, not linked: Mitigation_Industry column H is an unused placeholder),
    H 2024 EU exports kt, I pre-policy price USD/t, J codes, K source, L confidence. TAN fill (assumption, Low).
  * Mitigation_Industry, baseline block and policy block: F:G and I of rows 276-283 / 675-682 and G of rows
    288-295 / 687-694 become direct links (='Manual inputs'!E101 etc.), TAN fill. The N country switch and L:AI
    path formulas are untouched.
  * Check sheet: new section with one row = count of non-formula cells in those ranges (expected 0).
  * Settings: title and version-log row v1.4 (with the measured regression difference).
  * Regression: every sheet is read after CalculateFull before and after; the max abs difference over all
    pre-existing cells (excluding the Settings log and the new Manual inputs / Check rows) is reported and written to
    the version-log row. Expected 0.
  * Old/ gets a copy of v1.3; the root v1.3 file is moved there at the end (NORMS section 5).
Bookkeeping after a clean run (NORMS section 6): CAVEATS entry, instructions-egypt.yaml (TASK-T3 + TASK-1 notes),
EgyptTaskReference.md, context-egypt.md key-files table, tick T3 in TODO.md.
"""
import datetime
import os
import shutil
import sys

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "Old"))
from build_v0_4 import REVIEW, TAN, copy_formats  # noqa: E402

SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v1.3.xlsx")
SRC_OLD = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v1.3.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v1.4.xlsx")
VER = "v1.4"
MI = "Manual inputs"
MIT = "Mitigation_Industry"

R_TITLE, R_NOTE, R_HEAD, R0 = 98, 99, 100, 101        # new Manual inputs section; products R0..R0+7
NPROD = 8
# (production/growth/EU exports first row, price first row) per scenario block
BLOCKS = {"baseline": (276, 288), "policy": (675, 687)}
PRODUCTS = ["DRI-EAF steel", "Scrap-EAF steel", "BF-BOF steel (reference)", "Grey clinker (dry-process)",
            "Ammonia (net merchant)", "Urea", "Ammonium nitrate", "Primary aluminium"]
SOURCE = ("v0.7 ad hoc assumption (EgyptResultsInitial / AdHocCalculations), unsourced. Replace with CAPMAS / UN Comtrade "
          "(CN codes in Mitigation_Industry rows 288-295 labels).")


def values(ws, addr):
    return ws.Range(addr).Value


def snapshot(wb):
    """All used-range values per sheet (2-D tuples)."""
    snap = {}
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        v = ur.Value
        snap[ws.Name] = (ur.Row, ur.Column, v if isinstance(v, tuple) else ((v,),))
    return snap


def max_diff(before, after, skip):
    """Max abs numeric difference over cells present in the old used range; non-numeric must be equal.
    skip(sheet, row, col) -> True for cells to ignore (new rows, version log)."""
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
                    if d > worst:
                        worst = d
                    if d > 0 and len(bad) < 10:
                        bad.append((name, r, c, b, a))
                elif a != b:
                    worst = max(worst, 1.0)
                    if len(bad) < 10:
                        bad.append((name, r, c, b, a))
    return worst, bad


def build_section(wb):
    mi = wb.Worksheets(MI)
    mit = wb.Worksheets(MIT)
    assert mi.Cells(mi.Rows.Count, 4).End(-4162).Row <= 95, "Manual inputs already extended; revise R_TITLE/R0"
    for r in range(R_TITLE - 1, R0 + NPROD):
        assert all(c is None for c in mi.Range("A%d:AE%d" % (r, r)).Value[0]), "row %d not empty" % r

    # layout and value assertions on the calc sheet
    base0, price0 = BLOCKS["baseline"]
    for blk, (p0, q0) in BLOCKS.items():
        for k in range(NPROD):
            assert mit.Cells(p0 + k, 4).Value == (mit.Cells(base0 + k, 4).Value), "label mismatch %s %d" % (blk, k)
            for col in (6, 7, 9):   # F, G, I
                assert mit.Cells(p0 + k, col).Value == mit.Cells(base0 + k, col).Value, \
                    "baseline and policy differ at product %d col %d" % (k, col)
            assert mit.Cells(q0 + k, 7).Value == mit.Cells(price0 + k, 7).Value, "price differs, product %d" % k
            assert not mit.Cells(p0 + k, 6).HasFormula, "F%d already a formula" % (p0 + k)
            assert not mit.Cells(q0 + k, 7).HasFormula, "G%d already a formula" % (q0 + k)
        assert mit.Cells(p0 + 2, 4).Value.startswith("BF-BOF"), "unexpected block layout"

    # new section
    mi.Range("B%d" % R_TITLE).Value = "CBAM product market data (Task T3, v1.4)"
    mi.Range("B%d" % R_TITLE).Font.Bold = True
    mi.Range("D%d" % R_NOTE).Value = ("2024 production, growth, EU exports and pre-policy price, entered once and linked from "
                                      "both scenario blocks of Mitigation_Industry (F:G, I of rows 276-283 / 675-682; G of rows "
                                      "288-295 / 687-694). Values are the v0.7 ad hoc assumptions (Low confidence).")
    heads = {4: "Product", 5: "2024 production (kt)", 6: "Production growth (%/yr)",
             7: "2024 total exports (kt) - placeholder, not linked", 8: "2024 EU exports (kt)",
             9: "Pre-policy price (USD/t)", 10: "Codes", 11: "Source / rationale", 12: "Confidence"}
    for c, t in heads.items():
        mi.Cells(R_HEAD, c).Value = t
        mi.Cells(R_HEAD, c).Font.Bold = True
    fmts = {5: "#,##0", 6: "0.0%", 7: "#,##0", 8: "#,##0", 9: "#,##0"}
    for k in range(NPROD):
        r = R0 + k
        src = base0 + k
        stem = str(mit.Cells(price0 + k, 6).Value)          # e.g. stl.drg (product code in column F of the price rows)
        mi.Cells(r, 4).Value = PRODUCTS[k]
        mi.Cells(r, 5).Value = mit.Cells(src, 6).Value
        mi.Cells(r, 6).Value = mit.Cells(src, 7).Value
        mi.Cells(r, 8).Value = mit.Cells(src, 9).Value
        mi.Cells(r, 9).Value = mit.Cells(price0 + k, 7).Value
        mi.Cells(r, 10).Value = ("egy.in.prod.{0} | egy.in.prodg.{0} | egy.in.expeu.{0} | all.in.ppx.{0}.---").format(stem)
        mi.Cells(r, 11).Value = SOURCE
        mi.Cells(r, 12).Value = "Low"
        for c, f in fmts.items():
            mi.Cells(r, c).NumberFormat = f
        for c in (5, 6, 7, 8, 9):
            mi.Cells(r, c).Interior.Color = TAN
        mi.Cells(r, 12).Interior.Color = REVIEW

    # links in both blocks
    for blk, (p0, q0) in BLOCKS.items():
        for k in range(NPROD):
            r = R0 + k
            for col, mcol in ((6, "E"), (7, "F"), (9, "H")):
                cell = mit.Cells(p0 + k, col)
                cell.Formula = "='%s'!%s%d" % (MI, mcol, r)
                cell.Interior.Color = TAN
            cell = mit.Cells(q0 + k, 7)
            cell.Formula = "='%s'!I%d" % (MI, r)
            cell.Interior.Color = TAN


def add_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    parts = []
    for p0, q0 in BLOCKS.values():
        for rng in ("F%d:G%d" % (p0, p0 + 7), "I%d:I%d" % (p0, p0 + 7), "G%d:G%d" % (q0, q0 + 7)):
            parts.append("SUMPRODUCT(--NOT(ISFORMULA(%s!%s)))" % (MIT, rng))
    ws.Range("A%d" % r0).Value = "v1.4: CBAM market data moved to Manual inputs (Task T3)"
    ws.Range("A%d" % r0).Font.Bold = True
    ws.Range("A%d" % (r0 + 1)).Value = ("T3: numeric literals left in Mitigation_Industry F:G, I rows 276-283 / 675-682 and G rows "
                                        "288-295 / 687-694 (count; expected 0)")
    ws.Range("D%d" % (r0 + 1)).Formula = "=" + "+".join(parts)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    return r0, r0 + 1


def settings(wb, diff, chk_row):
    ws = wb.Worksheets("Settings")
    assert ws.Range("A1").Value.startswith("CPAT industry kernel - Egypt"), "unexpected Settings A1"
    ws.Range("A1").Value = "CPAT industry kernel - Egypt, final " + VER
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "T3: CBAM 2024 production, growth, EU exports and pre-policy prices moved from Mitigation_Industry (both blocks) to "
        "Manual inputs rows %d-%d and linked. No value changes: max abs difference over all pre-existing cells = %.3g. "
        "Check row %d (literal count, expected 0)." % (R_TITLE, R0 + NPROD - 1, diff, chk_row))
    ws.Range("C%d" % (last + 1)).WrapText = False


def main():
    for p in (SRC,):
        assert os.path.exists(p), "missing " + p
    if os.path.exists(DST):
        os.remove(DST)
    os.makedirs(os.path.join(HERE, "Old"), exist_ok=True)
    if not os.path.exists(SRC_OLD):
        shutil.copyfile(SRC, SRC_OLD)
    shutil.copyfile(SRC, DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(DST)
        xl.CalculateFull()
        before = snapshot(wb)
        mi_last_before = before[MI][0] + len(before[MI][2]) - 1
        chk_last_before = before["Check"][0] + len(before["Check"][2]) - 1

        build_section(wb)
        r0, chk = add_check(wb)
        xl.CalculateFull()
        after = snapshot(wb)

        def skip(name, r, c):
            if name == "Settings":
                return True
            if name == MI and r >= R_TITLE - 1:
                return True
            if name == "Check" and r > chk_last_before:
                return True
            return False

        diff, bad = max_diff(before, after, skip)
        # the T3 cells themselves: values must equal the old literals (they were read from them)
        print("max abs diff over pre-existing cells (Settings, new rows excluded): %.6g" % diff)
        for b in bad:
            print("  DIFF", b)
        count = wb.Worksheets("Check").Range("D%d" % chk).Value
        print("Check literal count (expected 0):", count)
        assert mi_last_before <= 95
        if diff != 0 or count != 0:
            print("NOT saved: regression or check failed. Inspect the output above (workbook left unsaved).")
            wb.Close(False)
            sys.exit(1)
        settings(wb, diff, chk)
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    os.replace(SRC, SRC_OLD)   # v1.3 now lives in Old/ (a copy was made above; this removes the root copy)
    print("saved", DST, "; v1.3 moved to Old/")


if __name__ == "__main__":
    main()
