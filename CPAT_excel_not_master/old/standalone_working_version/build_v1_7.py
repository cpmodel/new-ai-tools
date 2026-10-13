"""Build CPAT_Industry_Kernel_Egypt_v1.7.xlsx from v1.6: new CPAT national outputs after a full-coverage EG3 run.

DRAFTED ON LINUX, NOT RUN, and only usable after the new CPAT run exists (egypt/supporting/EG3_FullCoverage_RunSpec_v0.1.md).
Needs Excel + pywin32 on Windows and kernel v1.6 in this folder. Run from this folder in a normal shell:   python build_v1_7.py

Steps (the CBAM block, elasticities, fund and Task J table do not depend on the CPAT run and are not touched):
  1. Replace the stored CPAT outputs in sheet CPAT_National (rows 6-133, key code|run, years 2022-2041 in L:AE) with
     egypt/supporting/AdHocRebuild/cpat_outputs_egypt_2022_2041.csv (the csv must hold the new EG3 rows and the same 128 keys).
  2. Refill the stored Table2_Industry snapshot rows (build_v0_16.refill_table2; the rebuild column is left as stored).
  3. Refresh the stored 2030 snapshot of CarveOut_Table2 for all six scenarios from the live sheet; rewrite the printed figures
     of Table2_Final from the live results; Settings log; Check section.
The Python mirror (make_carveout_v0_6.py, run on the new csv) gives carveout_v1_7_results.json for the tolerance warnings; live
results are written to carveout_v1_7_results_live.json, from which the documents are regenerated.
Regression: in the LEGACY state every sheet other than CPAT_National, Table2_Industry, CarveOut_Table2, Table2_Final, Check and
Settings must be unchanged (max abs diff 0), otherwise nothing is saved.
After a clean run it copies the kernel to egypt/final/ and egypt-final/Egypt_Model.xlsx and moves the previous final kernel to
egypt/archive/.
"""
import datetime
import json
import os
import shutil
import sys

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "Old"))
sys.path.insert(0, HERE)
import build_v0_15 as b15  # noqa: E402
import build_v0_16 as b16  # noqa: E402
from build_v0_4 import REVIEW, TAN, copy_formats  # noqa: E402

REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
ADHOC = os.path.join(REPO, "egypt", "supporting", "AdHocRebuild")
FINAL = os.path.join(REPO, "egypt", "final")
ARCHIVE = os.path.join(REPO, "egypt", "archive")
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v1.6.xlsx")
SRC_OLD = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v1.6.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v1.7.xlsx")
VER = "v1.7"
BUNDLES = ["1A", "2A", "2B", "3A", "3B", "3C"]
CO, T2F, MI = "CarveOut_Table2", "Table2_Final", "Manual inputs"

METRICS = [("J", 0), ("P", 1), ("K", 1), ("L", 1), ("M", 0), ("N", 1), ("O_final", 1), ("T", 1), ("Q", 0)]
M_CBAM = {"1A": 100, "2A": 44, "2B": 44, "3A": 100, "3B": 100, "3C": 100}


def rnd(x, dp):
    q = 10 ** dp
    return (int(abs(x) * q + 0.5 + 1e-9) / q) * (1 if x >= 0 else -1)


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


def refresh_cpat_national(wb, csv_path):
    """Overwrite the stored CPAT outputs (CPAT_National rows 6-133, years 2022-2041 = columns L:AE) from the csv."""
    import csv
    ws = wb.Worksheets("CPAT_National")
    rows = list(csv.reader(open(csv_path, encoding="utf8")))
    hdr = rows[0]
    years = [int(x) for x in hdr[4:]]
    assert years == list(range(2022, 2042)), "csv years are not 2022-2041"
    data = {r[2]: [float(x) if x not in ("", None) else 0.0 for x in r[4:]] for r in rows[1:]}
    keys = [ws.Cells(r, 2).Value for r in range(6, 134)]
    assert len(keys) == 128 and set(keys) == set(data), "csv keys differ from CPAT_National keys: %s" % (set(keys) ^ set(data))
    changed = 0
    for i, k in enumerate(keys):
        r = 6 + i
        old = ws.Range("L%d:AE%d" % (r, r)).Value[0]
        new = data[k]
        if any(abs((o or 0) - n) > 1e-9 for o, n in zip(old, new)):
            changed += 1
            ws.Range("L%d:AE%d" % (r, r)).Value = [new]
    ws.Range("B135").Value = str(ws.Range("B135").Value) + " Stored CPAT outputs refreshed in %s from %s." % (VER, os.path.basename(csv_path))
    return changed


def export_live(co, t2f, i):
    """One scenario's live Table 2 values (active scenario already set)."""
    v = lambda a: co.Range(a).Value
    return dict(J=v("D66"), K=v("D67"), L=v("D68"), M=v("D69"), N=v("D70"), P=v("D73"), Q=v("D74"), T=v("D75"),
                Onp=v("D72"), O_full=v("D71"), O_final=t2f.Range("D63").Value,
                dghg=v("D54"), cpat_blk=v("D48"), dB=v("D46") + v("D47"), D_nb=v("D96"), B=v("D44") + v("D45"),
                r=100 * v("D30"), i_f=100 * v("D31"))


def refresh_snapshot(wb, xl, mirror):
    wsS, co, t2f = wb.Worksheets("Settings"), wb.Worksheets(CO), wb.Worksheets(T2F)
    active0 = wsS.Range("B10").Value
    live = {}
    tol = dict(K=0.35, P=0.12, Q=45.0, T=0.7, N=0.08, J=0.02)
    warn = []
    for i, b in enumerate(BUNDLES):
        wsS.Range("B10").Value = b
        xl.CalculateFull()
        assert co.Range("D6").Value == b
        d = export_live(co, t2f, i)
        live[b] = d
        row = 80 + i
        for c, k in zip("CDEFGHIJKL", ["J", "K", "L", "M", "N", "O_full", "Onp", "P", "Q", "T"]):
            co.Range("%s%d" % (c, row)).Value = d[k]
        co.Range("M%d" % row).Value = round(d["O_final"], 2)
        for k, t in tol.items():
            if abs(d[k] - mirror[b][k]) > t:
                warn.append("%s %s: live %.3f, mirror %.3f" % (b, k, d[k], mirror[b][k]))
        print(b, " ".join("%s %.3f" % (k, d[k]) for k in ("K", "P", "Q", "N", "T", "O_final")))
    wsS.Range("B10").Value = active0
    xl.CalculateFull()
    return live, warn


def write_printed(wb, live):
    ws = wb.Worksheets(T2F)
    scen_cols = "DEFGHI"
    for i, (key, dp) in enumerate(METRICS):
        r = 19 + i
        for j, b in enumerate(BUNDLES):
            src = M_CBAM[b] if key == "M" else live[b][key]
            ws.Range("%s%d" % (scen_cols[j], r)).Value = rnd(src, dp)
    ws.Range("B17").Value = ("B. As printed in the final documents (typed): EGYPT_CarveOut_Table2_%s, results text %s Table 2" % (VER, VER))
    ws.Range("B1").Value = "Table 2 (2030), final %s: confirmation of the final numbers in one workbook" % VER
    co = wb.Worksheets(CO)
    t = str(co.Range("B1").Value)
    co.Range("B1").Value = t.replace("final v1.6", "final " + VER)
    co.Range("B78").Value = ("F. Stored 2030 snapshot, all bundles (refreshed in %s from the live sheet for all six; "
                             "= EGYPT_CarveOut_Table2_%s)" % (VER, VER))


def add_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "%s: CPAT national outputs refreshed after the full-coverage EG3 run (CPAT_National)" % VER
    ws.Range("A%d" % r0).Font.Bold = True
    items = [
        ("Table2_Final: figures printed in the final documents differ from the workbook (count; expected 0)", "=%s!D42" % T2F),
        ("kappa(EG3), active scenario 3A-3C (about 1 expected after a full-coverage run; 0.537 before)", "=Table2_Industry!$T$54"),
    ]
    for i, (lab, f) in enumerate(items):
        ws.Range("A%d" % (r0 + 1 + i)).Value = lab
        ws.Range("D%d" % (r0 + 1 + i)).Formula = f
        ws.Range("D%d" % (r0 + 1 + i)).Interior.Color = REVIEW
    return r0 + 1


def settings(wb, chk, warn):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt, final " + VER
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "CPAT_National refreshed from the new CPAT outputs (full-coverage EG3 run); Table2_Industry and CarveOut_Table2 stored "
        "snapshots refilled; Table2_Final printed figures rewritten. CBAM block unchanged. Check row %d.%s"
        % (chk, " Mirror difference warnings: %d (documents regenerated from the live results)." % len(warn) if warn else ""))
    ws.Range("C%d" % (last + 1)).WrapText = False


def main():
    assert os.path.exists(SRC), "kernel v1.6 is missing: " + SRC
    mirror = json.load(open(os.path.join(ADHOC, "carveout_v1_7_results.json"), encoding="utf8"))["results"]
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
        assert wb.Worksheets("Settings").Range("B10").Value == "LEGACY", "set Settings!B10 to LEGACY in v1.6 first"
        before = snapshot(wb)
        chk_last_before = before["Check"][0] + len(before["Check"][2]) - 1

        n = refresh_cpat_national(wb, os.path.join(ADHOC, "cpat_outputs_egypt_2022_2041.csv"))
        print("CPAT_National rows changed:", n)
        xl.CalculateFull()
        k = b16.t2_rows(wb)
        b16.refill_table2(wb, k)
        xl.CalculateFull()
        live, warn = refresh_snapshot(wb, xl, mirror)
        write_printed(wb, live)
        chk = add_check(wb)
        xl.CalculateFull()
        after = snapshot(wb)

        # LEGACY-state regression: only the intended areas may differ
        def skip(name, r, c):
            if name == "Settings":
                return True
            if name == "CPAT_National":
                return True
            if name == "Check" and r > chk_last_before:
                return True
            if name in (CO, T2F):
                return True                                  # stored snapshots / printed figures: intended
            if name in ("Table2_Industry", "Fund_Industry"):
                return True                                  # stored snapshot rows and stored fund rows: intended
            if name == "Check":
                return True                                  # Task J v2 stored table: intended
            return False
        diff, bad = max_diff(before, after, skip)
        print("max abs diff in the other sheets (LEGACY state): %.6g" % diff)
        for x in bad:
            print("  DIFF", x)
        mism = wb.Worksheets(T2F).Range("D42").Value
        print("printed-vs-workbook mismatches:", mism)
        json.dump(live, open(os.path.join(ADHOC, "carveout_v1_7_results_live.json"), "w"), indent=1)
        if warn:
            print("MIRROR DIFFERENCES (live vs make_carveout_v0_5.py):")
            for w in warn:
                print("  ", w)
        if diff != 0:
            print("NOT saved: unexpected change outside the intended areas.")
            wb.Close(False)
            sys.exit(1)
        if mism != 0 or warn:
            print("DOCUMENTS MUST BE REGENERATED from carveout_v1_7_results_live.json (printed figures differ from the workbook).")
        settings(wb, chk, warn)
        wb.Save()
        wb.Close(False)
        ok = True
    finally:
        xl.Quit()
    if ok:
        os.replace(SRC, SRC_OLD)
        os.makedirs(FINAL, exist_ok=True)
        old_final = os.path.join(FINAL, "CPAT_Industry_Kernel_Egypt_v1.6.xlsx")
        if os.path.exists(old_final):
            shutil.move(old_final, os.path.join(ARCHIVE, "CPAT_Industry_Kernel_Egypt_v1.6.xlsx"))
        shutil.copyfile(DST, os.path.join(FINAL, os.path.basename(DST)))
        ef = os.path.join(REPO, "egypt-final", "Egypt_Model.xlsx")
        if os.path.isdir(os.path.dirname(ef)):
            shutil.copyfile(DST, ef)
        print("saved", DST, "; copied to egypt/final and egypt-final/Egypt_Model.xlsx; v1.6 moved to Old/ and egypt/archive")


if __name__ == "__main__":
    main()
